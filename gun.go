package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylarn/excelize"
	fundx "github.com/daniel-hutao/fund/fund"
	"github.com/daniel-hutao/gun/fund"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/daniel-hutao/gun/excel"
)

// 需求1：
// 第一期：表格中有N个以基金代码开头命名的工作簿，第一列有已经存在的日期（从第三行开始）；
// 程序遍历第一列，看第二列内容是否为空，如果第二列为空，则判断这一行需要新数据，
// 从这一行开始往后所有第一列存在内容的行都尝试构造数据，直到获取数据失败，
// 然后将构造好的多行数据填充到工作簿中

func main() {
	e := excel.NewExcel("test.xlsx")
	sheetMap := e.File.GetSheetMap()

	for _, sheetName := range sheetMap {
		shot(e, sheetName)
		// get jz from server
		//jz, err := fundx.GetRain().GetOneFundJZ(fundCode, date)
		//if err != nil {
		//	log.Fatal(err)
		//}

		// get the coordinate with date
		//cors, err := e.File.SearchSheet(sheetName, strconv.Itoa(excel.DateToInt(date)))
		//if err != nil {
		//	log.Fatal(err)
		//}
		//if len(cors) != 1 {
		//	log.Fatal("cores != 1")
		//}
		//cor := cors[0]
		//c, r, err := excelize.CellNameToCoordinates(cor)
		//if err != nil {
		//	log.Fatal(err)
		//}

		// set value
		//cor1, err := excelize.CoordinatesToCellName(c+1, r)
		//if err != nil {
		//	log.Fatal(err)
		//}

		//e.File.SetCellFloat(sheetName, cor1, jz, 4, 64)

	}

	e.File.Save()
}

// deal one sheet
func shot(e *excel.Excel, sheetName string) {
	// get code
	fundCode := strings.Split(sheetName, "-")[0]
	if len(fundCode) != 6 {
		log.Fatal("err code!")
	}
	log.Printf("Got fund code: %s", fundCode)

	rows, err := e.File.Rows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// get start date
	var startDateFloatStr string
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			log.Fatal(err)
		}
		// only date exist; 1-10 => it's a date format
		if len(row) == 1 && len(row[0]) > 1 && len(row[0]) < 10 {
			startDateFloatStr = row[0]
			break
			// the date exist and T日 is ""
		} else if len(row) > 1 && len(row[0]) > 1 && len(row[0]) < 10 && len(row[1]) == 0 {
			startDateFloatStr = row[0]
			break
		}
	}
	if len(startDateFloatStr) == 0 {
		log.Printf("Failed to get start date!")
	}
	log.Printf("Start date float is: %s", startDateFloatStr)

	startDateFloat, err := strconv.ParseFloat(startDateFloatStr, 64)
	if err != nil {
		log.Fatal(err)
	}
	startDate, err := excelize.ExcelDateToTime(startDateFloat, false)
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("Start date is: %s", startDate)

	dateA := startDate.Format("2006-01-02")
	dateB := time.Now().Format("2006-01-02")
	if dateA > dateB {
		log.Fatal("dateA > dateB")
	}

	log.Printf("Prepare to get jz from %s to %s", dateA, dateB)
	jzs, err := fundx.GetRain().GetOneFundJZFromDateAtoDateB(fundCode, dateA, dateB)
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("Got jzs: %v", jzs)

	if len(jzs) == 0 {
		log.Printf("Got jzs is []")
		return
	}

	var fundRows = make([]fund.FundRow, 0)
	for _, jz := range jzs { // TODO 在获取到开始日期的时候同时需要拿到起始行号，可以在迭代的时候顺带记录行号
		fundRows = append(fundRows, StructureFundRow(jz, 1))
	}
}

func StructureFundRow(jz float64, row int) fund.FundRow {
	rowBefore := row - 1
	return fund.FundRow{
		CyTr:  fmt.Sprintf("B%d+1", rowBefore),
		Dwjz:  jz,
		Cyfe:  fmt.Sprintf("D%d+G%d", rowBefore, rowBefore),
		Zczz:  fmt.Sprintf("C%d*D%d", row, row),
		Cccb:  fmt.Sprintf("I%d+F%d+H%d", rowBefore, rowBefore, rowBefore),
		Febh:  "0.00",
		Sxf:   "0.00",
		Qrje:  fmt.Sprintf("G%d*C%d", row, row),
		Ccsy:  fmt.Sprintf("D%d*C%d-F%d", row, row, row),
		Ccsyl: fmt.Sprintf("J%d/F%d", row, row),
	}
}
