// test_excel project main.go
// author: plmsuper8
/*
This CLT program transfer excel to plain txt.
	Support .xlsx on all platform, and .xls on windows (maybe buggy).
	So far, only extract Sheet1.
	working list:
		Done!.support both XLS and XLSX
		Done!.support flag (see -help)
		Done!.fix utf8 on windows by adding BOM head (-bom)
		Done!.Check xls binary head <D0 CF 11 E0 A1 B1 1A E1>
		.Multi-threads, only print finished file
		.flag sheet id, cols, rows etc.
*/

package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"

	"github.com/Luxurioust/excelize"
	"github.com/aswjh/excel"
)

const (
	IRETURN = "\n"
)

var (
	isep    string       = "\t"
	dirname string       = "./"
	isbom   bool         = false
	optxls  excel.Option = excel.Option{
		"Visible":       false,
		"DisplayAlerts": true,
		"Readonly":      true}
)

func init() {
	//read flags
	flag.StringVar(&dirname, "dirname", "./", "the target directory or xlsx file")
	flag.StringVar(&isep, "sep", "\t", "seperator of output")
	flag.BoolVar(&isbom, "bom", false, `add byte sequence <EF BB BF> in head
		of utf8 file. Required by Microsoft, but not for Linux`)
	flag.Parse()
}

func check(e error) {
	if e != nil {
		log.Fatal(e)
	}
}

func main() {
	//check files
	if isbom == true {
		fmt.Print("\xef\xbb\xbf")
	}
	if st, err := os.Stat(dirname); st.IsDir() {
		check(err)
		//if dir is a directory
		dinfo, _ := os.Open(dirname)
		dinfo.Chdir()
		list, err2 := dinfo.Readdirnames(-1)
		check(err2)
		for _, fname := range list {
			if ism, _ := regexp.MatchString(".xlsx$", fname); ism == true {
				printXLSX(fname)
			} else if ism, _ := regexp.MatchString(".xls$", fname); ism == true {
				printXLS(fname)
			} else {
				fmt.Println("#Not Excel:", fname)
			}
		}
	} else if ism, _ := regexp.MatchString(".xlsx$", dirname); ism == true {
		printXLSX(dirname)
	} else if ism, _ := regexp.MatchString(".xls$", dirname); ism == true {
		printXLS(dirname)
	}
	/*DEBUG: wait to exit
	fmt.Println("#Scan finished! Return to exit...")
	buf := bufio.NewReader(os.Stdin)
	buf.ReadBytes('\n')
	*/
	os.Exit(0)
}

func printXLSX(f string) (res int) {
	fmt.Println("#Excel:", f)
	xlsx, err := excelize.OpenFile(f)
	check(err)
	//xlsx.SetActiveSheet(1)
	rows := xlsx.GetRows("Sheet1")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, isep)
		}
		fmt.Print(IRETURN)
	}
	fmt.Println("#Excel End:", f)
	return 0
}

func printXLS(f string) (res int) {
	//check is real excel
	//
	defer func() {
		if r := recover(); r != nil {
			fmt.Println("#Excel Skipped:", r)
		}
	}()
	f, _ = filepath.Abs(f)
	t_f, _ := os.Open(f)
	ib := make([]byte, 8)
	var IB_HEAD_XLS = []byte{208, 207, 17, 224, 161, 177, 26, 225}
	_, err1 := t_f.Read(ib)
	check(err1)
	for i, _ := range IB_HEAD_XLS {
		if IB_HEAD_XLS[i] != ib[i] {
			fmt.Println("#Not Excel:", f)
			return 1
		}
	}
	xls, err := excel.Open(f, optxls)
	check(err)
	fmt.Println("#Excel:", f)
	defer xls.Quit()
	sheet, _ := xls.Sheet("Sheet1")
	defer sheet.Release()
	sheet.ReadRow("A", 1, func(row []interface{}) (rc int) {
		for _, cell := range row {
			fmt.Print(cell, isep)
		}
		fmt.Print(IRETURN)
		//fmt.Println(row)
		return 0
	})
	fmt.Println("#Excel End:", f)
	return 0
}
