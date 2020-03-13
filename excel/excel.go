package excel

import (
	"github.com/360EntSecGroup-Skylarn/excelize"
	"log"
	"os"
)

type Excel struct {
	File *excelize.File
}

func NewExcel(filename string) *Excel {
	e := Excel{}
	if _, err := os.Stat(filename); err != nil {
		if os.IsNotExist(err) {
			e.File = excelize.NewFile()
			err1 := e.File.SaveAs(filename)
			if err1 != nil {
				log.Fatal(err1)
			}
			e.OpenFile(filename)
		} else {
			log.Fatal(err)
		}
	} else {
		e.OpenFile(filename)
	}
	return &e
}

func (e *Excel) OpenFile(filename string) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	e.File = f
}
