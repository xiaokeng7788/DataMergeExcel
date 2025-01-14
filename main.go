package main

import (
	"DataMergeExcel/utils"
)

func main() {
	//dir := "D:\\桌面\\测试表格"
	//utils.MergeMuchExcelOneIndexExcel(dir, "测试表格.xlsx", "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)
	//
	//x := "D:\\桌面\\测试表格\\重复表格\\测试表格-1227.xlsx"
	//y := "D:\\桌面\\测试表格\\重复表格\\测试表格.xlsx"
	//utils.MergeMuchExcelOneRepeatExcel(y, x, "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)

	dir1 := "D:\\桌面\\测试表格\\测试表格.xlsx"
	out := "D:\\桌面\\测试表格\\生成表格"
	utils.MergeWorkSheetData(dir1, "Sheet1", "姓名", out, 2)

}
