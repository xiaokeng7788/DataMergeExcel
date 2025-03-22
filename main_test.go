package dataMergeExcel

import (
	"fmt"
	"testing"
)

const dir string = "D:\\桌面\\main\\2024年服务总次数.xlsx"
const sheetName string = "Sheet3"
const title string = "1"
const titleNum int = 1
const out string = "D:\\桌面\\watch"
const outFileName string = "out.xlsx"

func TestMainTest(t *testing.T) {
	excelFile, err := OpenExcelFile(dir)
	if err != nil {
		t.Error(err)
	}
	excelFile.SheetName = sheetName
	err = excelFile.IsExitSheetName(false)
	if err != nil {
		t.Error(err)
	}
	data, err := excelFile.GetExcelSheetData()
	if err != nil {
		t.Error(err)
	}
	fmt.Println(data)
}
