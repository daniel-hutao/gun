// Copyright 2016 - 2020 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.
//
// Package excelize providing a set of functions that allow you to write to
// and read from XLSX files. Support reads and writes XLSX file generated by
// Microsoft Excel™ 2007 and later. Support save file without losing original
// charts of XLSX. This library needs Go version 1.10 or later.

package excelize

import (
	"errors"
	"strings"
)

type adjustDirection bool

const (
	columns adjustDirection = false
	rows    adjustDirection = true
)

// adjustHelper provides a function to adjust rows and columns dimensions,
// hyperlinks, merged cells and auto filter when inserting or deleting rows or
// columns.
//
// sheet: Worksheet name that we're editing
// column: Index number of the column we're inserting/deleting before
// row: Index number of the row we're inserting/deleting before
// offset: Number of rows/column to insert/delete negative values indicate deletion
//
// TODO: adjustPageBreaks, adjustComments, adjustDataValidations, adjustProtectedCells
//
func (f *File) adjustHelper(sheet string, dir adjustDirection, num, offset int) error {
	xlsx, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	if dir == rows {
		f.adjustRowDimensions(xlsx, num, offset)
	} else {
		f.adjustColDimensions(xlsx, num, offset)
	}
	f.adjustHyperlinks(xlsx, sheet, dir, num, offset)
	if err = f.adjustMergeCells(xlsx, dir, num, offset); err != nil {
		return err
	}
	if err = f.adjustAutoFilter(xlsx, dir, num, offset); err != nil {
		return err
	}
	if err = f.adjustCalcChain(dir, num, offset); err != nil {
		return err
	}
	checkSheet(xlsx)
	_ = checkRow(xlsx)

	if xlsx.MergeCells != nil && len(xlsx.MergeCells.Cells) == 0 {
		xlsx.MergeCells = nil
	}

	return nil
}

// adjustColDimensions provides a function to update column dimensions when
// inserting or deleting rows or columns.
func (f *File) adjustColDimensions(xlsx *xlsxWorksheet, col, offset int) {
	for rowIdx := range xlsx.SheetData.Row {
		for colIdx, v := range xlsx.SheetData.Row[rowIdx].C {
			cellCol, cellRow, _ := CellNameToCoordinates(v.R)
			if col <= cellCol {
				if newCol := cellCol + offset; newCol > 0 {
					xlsx.SheetData.Row[rowIdx].C[colIdx].R, _ = CoordinatesToCellName(newCol, cellRow)
				}
			}
		}
	}
}

// adjustRowDimensions provides a function to update row dimensions when
// inserting or deleting rows or columns.
func (f *File) adjustRowDimensions(xlsx *xlsxWorksheet, row, offset int) {
	for i, r := range xlsx.SheetData.Row {
		if newRow := r.R + offset; r.R >= row && newRow > 0 {
			f.ajustSingleRowDimensions(&xlsx.SheetData.Row[i], newRow)
		}
	}
}

// ajustSingleRowDimensions provides a function to ajust single row dimensions.
func (f *File) ajustSingleRowDimensions(r *xlsxRow, num int) {
	r.R = num
	for i, col := range r.C {
		colName, _, _ := SplitCellName(col.R)
		r.C[i].R, _ = JoinCellName(colName, num)
	}
}

// adjustHyperlinks provides a function to update hyperlinks when inserting or
// deleting rows or columns.
func (f *File) adjustHyperlinks(xlsx *xlsxWorksheet, sheet string, dir adjustDirection, num, offset int) {
	// short path
	if xlsx.Hyperlinks == nil || len(xlsx.Hyperlinks.Hyperlink) == 0 {
		return
	}

	// order is important
	if offset < 0 {
		for rowIdx, linkData := range xlsx.Hyperlinks.Hyperlink {
			colNum, rowNum, _ := CellNameToCoordinates(linkData.Ref)

			if (dir == rows && num == rowNum) || (dir == columns && num == colNum) {
				f.deleteSheetRelationships(sheet, linkData.RID)
				if len(xlsx.Hyperlinks.Hyperlink) > 1 {
					xlsx.Hyperlinks.Hyperlink = append(xlsx.Hyperlinks.Hyperlink[:rowIdx],
						xlsx.Hyperlinks.Hyperlink[rowIdx+1:]...)
				} else {
					xlsx.Hyperlinks = nil
				}
			}
		}
	}

	if xlsx.Hyperlinks == nil {
		return
	}

	for i := range xlsx.Hyperlinks.Hyperlink {
		link := &xlsx.Hyperlinks.Hyperlink[i] // get reference
		colNum, rowNum, _ := CellNameToCoordinates(link.Ref)

		if dir == rows {
			if rowNum >= num {
				link.Ref, _ = CoordinatesToCellName(colNum, rowNum+offset)
			}
		} else {
			if colNum >= num {
				link.Ref, _ = CoordinatesToCellName(colNum+offset, rowNum)
			}
		}
	}
}

// adjustAutoFilter provides a function to update the auto filter when
// inserting or deleting rows or columns.
func (f *File) adjustAutoFilter(xlsx *xlsxWorksheet, dir adjustDirection, num, offset int) error {
	if xlsx.AutoFilter == nil {
		return nil
	}

	coordinates, err := f.areaRefToCoordinates(xlsx.AutoFilter.Ref)
	if err != nil {
		return err
	}
	x1, y1, x2, y2 := coordinates[0], coordinates[1], coordinates[2], coordinates[3]

	if (dir == rows && y1 == num && offset < 0) || (dir == columns && x1 == num && x2 == num) {
		xlsx.AutoFilter = nil
		for rowIdx := range xlsx.SheetData.Row {
			rowData := &xlsx.SheetData.Row[rowIdx]
			if rowData.R > y1 && rowData.R <= y2 {
				rowData.Hidden = false
			}
		}
		return nil
	}

	coordinates = f.adjustAutoFilterHelper(dir, coordinates, num, offset)
	x1, y1, x2, y2 = coordinates[0], coordinates[1], coordinates[2], coordinates[3]

	if xlsx.AutoFilter.Ref, err = f.coordinatesToAreaRef([]int{x1, y1, x2, y2}); err != nil {
		return err
	}
	return nil
}

// adjustAutoFilterHelper provides a function for adjusting auto filter to
// compare and calculate cell axis by the given adjust direction, operation
// axis and offset.
func (f *File) adjustAutoFilterHelper(dir adjustDirection, coordinates []int, num, offset int) []int {
	if dir == rows {
		if coordinates[1] >= num {
			coordinates[1] += offset
		}
		if coordinates[3] >= num {
			coordinates[3] += offset
		}
	} else {
		if coordinates[2] >= num {
			coordinates[2] += offset
		}
	}
	return coordinates
}

// areaRefToCoordinates provides a function to convert area reference to a
// pair of coordinates.
func (f *File) areaRefToCoordinates(ref string) ([]int, error) {
	rng := strings.Split(ref, ":")
	return areaRangeToCoordinates(rng[0], rng[1])
}

// areaRangeToCoordinates provides a function to convert cell range to a
// pair of coordinates.
func areaRangeToCoordinates(firstCell, lastCell string) ([]int, error) {
	coordinates := make([]int, 4)
	var err error
	coordinates[0], coordinates[1], err = CellNameToCoordinates(firstCell)
	if err != nil {
		return coordinates, err
	}
	coordinates[2], coordinates[3], err = CellNameToCoordinates(lastCell)
	return coordinates, err
}

// sortCoordinates provides a function to correct the coordinate area, such
// correct C1:B3 to B1:C3.
func sortCoordinates(coordinates []int) error {
	if len(coordinates) != 4 {
		return errors.New("coordinates length must be 4")
	}
	if coordinates[2] < coordinates[0] {
		coordinates[2], coordinates[0] = coordinates[0], coordinates[2]
	}
	if coordinates[3] < coordinates[1] {
		coordinates[3], coordinates[1] = coordinates[1], coordinates[3]
	}
	return nil
}

// coordinatesToAreaRef provides a function to convert a pair of coordinates
// to area reference.
func (f *File) coordinatesToAreaRef(coordinates []int) (string, error) {
	if len(coordinates) != 4 {
		return "", errors.New("coordinates length must be 4")
	}
	firstCell, err := CoordinatesToCellName(coordinates[0], coordinates[1])
	if err != nil {
		return "", err
	}
	lastCell, err := CoordinatesToCellName(coordinates[2], coordinates[3])
	if err != nil {
		return "", err
	}
	return firstCell + ":" + lastCell, err
}

// adjustMergeCells provides a function to update merged cells when inserting
// or deleting rows or columns.
func (f *File) adjustMergeCells(xlsx *xlsxWorksheet, dir adjustDirection, num, offset int) error {
	if xlsx.MergeCells == nil {
		return nil
	}

	for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
		areaData := xlsx.MergeCells.Cells[i]
		coordinates, err := f.areaRefToCoordinates(areaData.Ref)
		if err != nil {
			return err
		}
		x1, y1, x2, y2 := coordinates[0], coordinates[1], coordinates[2], coordinates[3]
		if dir == rows {
			if y1 == num && y2 == num && offset < 0 {
				f.deleteMergeCell(xlsx, i)
				i--
			}
			y1 = f.adjustMergeCellsHelper(y1, num, offset)
			y2 = f.adjustMergeCellsHelper(y2, num, offset)
		} else {
			if x1 == num && x2 == num && offset < 0 {
				f.deleteMergeCell(xlsx, i)
				i--
			}
			x1 = f.adjustMergeCellsHelper(x1, num, offset)
			x2 = f.adjustMergeCellsHelper(x2, num, offset)
		}
		if x1 == x2 && y1 == y2 {
			f.deleteMergeCell(xlsx, i)
			i--
		}
		if areaData.Ref, err = f.coordinatesToAreaRef([]int{x1, y1, x2, y2}); err != nil {
			return err
		}
	}
	return nil
}

// adjustMergeCellsHelper provides a function for adjusting merge cells to
// compare and calculate cell axis by the given pivot, operation axis and
// offset.
func (f *File) adjustMergeCellsHelper(pivot, num, offset int) int {
	if pivot >= num {
		pivot += offset
		if pivot < 1 {
			return 1
		}
		return pivot
	}
	return pivot
}

// deleteMergeCell provides a function to delete merged cell by given index.
func (f *File) deleteMergeCell(sheet *xlsxWorksheet, idx int) {
	if len(sheet.MergeCells.Cells) > idx {
		sheet.MergeCells.Cells = append(sheet.MergeCells.Cells[:idx], sheet.MergeCells.Cells[idx+1:]...)
		sheet.MergeCells.Count = len(sheet.MergeCells.Cells)
	}
}

// adjustCalcChain provides a function to update the calculation chain when
// inserting or deleting rows or columns.
func (f *File) adjustCalcChain(dir adjustDirection, num, offset int) error {
	if f.CalcChain == nil {
		return nil
	}
	for index, c := range f.CalcChain.C {
		colNum, rowNum, err := CellNameToCoordinates(c.R)
		if err != nil {
			return err
		}
		if dir == rows && num <= rowNum {
			if newRow := rowNum + offset; newRow > 0 {
				f.CalcChain.C[index].R, _ = CoordinatesToCellName(colNum, newRow)
			}
		}
		if dir == columns && num <= colNum {
			if newCol := colNum + offset; newCol > 0 {
				f.CalcChain.C[index].R, _ = CoordinatesToCellName(newCol, rowNum)
			}
		}
	}
	return nil
}
