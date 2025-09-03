package dataMergeExcel

import (
	"bytes"
	"encoding/json"
	"fmt"
	"testing"
)

const dir string = "D:\\桌面\\各种统计结果"
const sheetName string = "Sheet2"
const title string = "1"
const titleNum int = 1
const out string = "D:\\桌面\\main"
const outFileName string = "out.xlsx"

// 测试获取数据是否正常
func TestCreatedExcel(t *testing.T) {
	excel := NewCreateExcel()
	excel.SetImportConfig(dir, "年龄段统计.xlsx")
	excel.SetSheetConfig(sheetName, title)
	err := excel.OpenExcelFile()
	if err != nil {
		t.Error(err)
		return
	}
	data, err := excel.GetExcelSheetData()
	if err != nil {
		t.Error(err)
		return
	}
	for i, datum := range data {
		fmt.Println(i, datum)
	}
	//	输出文件流
	buffer, err := excel.ExportExcelBuffer()
	if err != nil {
		t.Error(err)
		return
	}
	stream, err := OpenExcelStream(bytes.NewReader(buffer))
	if err != nil {
		t.Error(err)
		return
	}
	stream.SetSheetConfig(sheetName, title)
	data, err = stream.GetExcelSheetData()
	if err != nil {
		t.Error(err)
		return
	}
	for i, datum := range data {
		fmt.Println(i, datum)
	}
}

func TestGetExcelAppointIndexRepeatData2(t *testing.T) {
	var rr = make([][]string, 0)
	err := json.Unmarshal([]byte(`[["出生日期","家庭医生工号","居住地址","居民姓名","联系电话","身份证号","本人手机号","签约机构编码","性别","签约时间","签约状态","续约时间","续约告知方式","签约开始时间(续约时间)"],["1977-12-18","112","新建一路1599弄15号703室","卢荣刚","","342622197712681216","13858077252","42505437031011411B1001","男","2024-09-20","1","","223","2025-09-21 00:00:00"]]`), &rr)
	if err != nil {
		t.Error(err)
		return
	}
	excel := NewCreateExcel()
	excel.SetExportConfig(out, outFileName)
	excel.SetSheetConfig(sheetName, title)
	err = excel.WriteExcelSheet(rr)
	err = excel.ExportExcel()
	if err != nil {
		t.Error(err)
		return
	}
}
