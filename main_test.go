package dataMergeExcel

import (
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/excelUtils"
	"testing"
)

const dir string = "D:\\桌面\\test.xlsx"
const sheetName string = "244"
const title string = "通话id"
const titleNum int = 1
const out string = "D:\\桌面\\watch"
const outFileName string = "out.xlsx"

func Test(t *testing.T) {
	res, err := excelUtils.GetExcelSheetData(dir, sheetName)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(res)
}

func TestCreatedExcel(t *testing.T) {
	res := map[string][][]string{}
	res["244"] = [][]string{[]string{"直播ID", "用户ID", "用户姓名", "观看次数", "观看时长(秒)", "入会时间", "退会时间"}}
	res["255"] = [][]string{[]string{"直播ID1", "用户ID1", "用户姓名1", "观看次数1", "观看时长(秒)1", "入会时间1", "退会时间1"}}
	result := []string{"244", "255"}
	err := excelUtils.BatchCreateExcel(out, outFileName, result, res, titleNum)
	if err != nil {
		t.Error(err)
	}
}
