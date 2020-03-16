package excel

import (
	"log"
	"os"

	"github.com/360EntSecGroup-Skylarn/excelize"
)

type Excel struct {
	File *excelize.File
}

// 存在则打开，不存在则新建
func NewExcel(filename string) *Excel {
	e := Excel{}
	if _, err := os.Stat(filename); err != nil {
		if os.IsNotExist(err) {
			e.File = excelize.NewFile()
			err1 := e.File.SaveAs(filename)
			if err1 != nil {
				log.Fatal(err1)
			}
			e.openFile(filename)
		} else {
			log.Fatal(err)
		}
	} else {
		e.openFile(filename)
	}
	return &e
}

func (e *Excel) openFile(filename string) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	e.File = f
}
