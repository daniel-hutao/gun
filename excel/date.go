package excel

import (
	"log"
	"time"
)

// 2020-03-15 -> 43905
func DateToInt(dateStr string) int {
	date, err := time.Parse("2006-01-02", dateStr)
	if err != nil {
		log.Fatal(err)
	}
	beginDate := time.Date(1900, 1, 1, 0, 0, 0, 0, time.Now().Location())
	sub := date.Sub(beginDate).Hours() / 24
	// 1900-01-01 -> 1
	// 1900-02-29 -> +1
	return int(sub) + 1 + 1
}

// 43905 -> 2020-03-15
func IntToDate(dateInt int) string {
	beginDate := time.Date(1900, 1, 1, 0, 0, 0, 0, time.Now().Location())
	addDur := time.Duration(time.Duration(24*dateInt) * time.Hour)
	date := beginDate.Add(addDur - (1+1)*24*time.Hour)
	return date.Format("2006-01-02")
}
