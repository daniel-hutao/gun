package main

import (
	"fmt"
	"github.com/daniel-hutao/fund/fund"
	. "github.com/daniel-hutao/gun/excel"
	"log"
	"strconv"
)

// 需求1：
// 第一期：表格中有N个以基金名称命名的工作簿，第一行有基金代码，第一列有已经存在的日期；
// 程序接受指定的日期，然后拉取该日期的各个基金净值信息，填充进 Excel 中相应的所有工作簿。
// 并且将剩下每一格按照无交易的默认值(公式)填充好。

func main() {
	e := NewExcel("Book1.xlsx")
	intStr, err := e.File.GetCellValue("Sheet1", "A1")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(intStr)
	fmt.Println(DateToInt("2020-03-15"))
	fmt.Println(IntToDate(43905))

	jz := fund.GetRain().GetOneFundJZ("001717", "2020-03-13")
	fmt.Println(strconv.FormatFloat(jz, 64))
	//fmt.Println(time.UTC)
}
