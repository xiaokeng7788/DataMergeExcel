package dataMergeExcel

import (
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/excelUtils"
	"testing"
)

const dir string = "D:\\桌面\\test.xlsx"
const sheetName string = "Sheet1"
const title string = "通话id"
const titleNum uint = 1
const out string = "D:\\桌面"
const outFileName string = "out.xlsx"

func Test(t *testing.T) {
	res, err := excelUtils.GetExcelSheetData(dir, sheetName)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(res)
}
