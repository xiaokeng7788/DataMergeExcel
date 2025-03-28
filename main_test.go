package dataMergeExcel

import (
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/excelUtils"
	"testing"
)

const dir string = "D:\\桌面\\test"
const sheetName string = "Sheet1"
const title string = "1"
const titleNum int = 1
const out string = "D:\\桌面"
const outFileName string = "out.xlsx"

func TestMainTest(t *testing.T) {
	excelFile, err := OpenExcelFile(dir + "\\13.xlsx")
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
	firstNum, appointNum, err := GetExcelTitleInfo(data, title, titleNum)
	if err != nil {
		t.Error(err)
	}
	res, _, err := ConvertToOneDimension(data, firstNum, titleNum, appointNum, true)
	if err != nil {
		t.Error(err)
	}
	excelFile2, err := OpenExcelFile(dir + "\\12.xlsx")
	if err != nil {
		t.Error(err)
	}
	excelFile2.SheetName = sheetName
	err = excelFile2.IsExitSheetName(false)
	if err != nil {
		t.Error(err)
	}
	data2, err := excelFile2.GetExcelSheetData()
	if err != nil {
		t.Error(err)
	}
	firstNum2, appointNum2, err := GetExcelTitleInfo(data, title, titleNum)
	if err != nil {
		t.Error(err)
	}
	res2, _, err := ConvertToOneDimension(data2, firstNum2, titleNum, appointNum2, true)
	if err != nil {
		t.Error(err)
	}
	excel := excelUtils.MergeMuchExcelOneIndexExcel(res, res2)
	fmt.Println(excel)
}

func TestCreatedExcel(t *testing.T) {
	res := make([][]string, 0)
	row := []string{"1", "2", "3"}
	res = append(res, row)
	res = append(res, row)
	res = append(res, row)
	excel := NewCreateExcel()
	excel.SheetName = sheetName
	excel.OutPath = out
	excel.OutFile = outFileName
	excel.TitleNum = titleNum
	err := excel.IsExitSheetName(true)
	if err != nil {
		t.Error(err)
	}
	err = excel.WriteExcelSheet(res)
	if err != nil {
		t.Error(err)
	}
	err = excel.CreatedExcelPath()
	if err != nil {
		t.Error(err)
	}
}
