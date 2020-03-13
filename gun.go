package main

import (
	. "github.com/daniel-hutao/gun/excel"
)

func main() {
	e := NewExcel("Book1.xlsx")
	e.File.SetCellInt("Sheet1", "A1", 2)
	e.File.Save()
}
