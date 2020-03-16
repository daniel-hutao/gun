package main

import (
	"log"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"

	"github.com/daniel-hutao/fund/fund"
	"github.com/daniel-hutao/gun/excel"
)

// 需求1：
// 第一期：表格中有N个以基金名称命名的工作簿，第一行有基金代码，第一列有已经存在的日期；
// 程序接受指定的日期，然后拉取该日期的各个基金净值信息，填充进 Excel 中相应的所有工作簿。
// 并且将剩下每一格按照无交易的默认值(公式)填充好。

func main() {
	var date = "2020-03-13"

	e := excel.NewExcel("fund.xlsx")
	sheetMap := e.File.GetSheetMap()

	for _, sheetName := range sheetMap {
		// get code from the sheet
		title, err := e.File.GetCellValue(sheetName, "A1")
		if err != nil {
			log.Fatal(err)
		}
		fundCode := strings.Split(title, "-")[1]

		// get jz from server
		jz, err := fund.GetRain().GetOneFundJZ(fundCode, date)
		if err != nil {
			log.Fatal(err)
		}

		// get the coordinate with date
		cors, err := e.File.SearchSheet(sheetName, strconv.Itoa(excel.DateToInt(date)))
		if err != nil {
			log.Fatal(err)
		}
		if len(cors) != 1 {
			log.Fatal("cores != 1")
		}
		cor := cors[0]
		c, r, err := excelize.CellNameToCoordinates(cor)
		if err != nil {
			log.Fatal(err)
		}

		// set value
		cor1, err := excelize.CoordinatesToCellName(c+1, r)
		if err != nil {
			log.Fatal(err)
		}

		e.File.SetCellFloat(sheetName, cor1, jz, 4, 64)

	}

	e.File.Save()
}
