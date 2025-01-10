package main

import (
	"DataMergeExcel/common"
	"DataMergeExcel/utils"
	"fmt"
)

func main() {
	dir := "D:\\桌面\\测试表格"
	utils.MergeMuchExcelOneIndexExcel(dir, "测试表格.xlsx", "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)

	x := "D:\\桌面\\测试表格\\重复表格\\测试表格-1227.xlsx"
	y := "D:\\桌面\\测试表格\\重复表格\\测试表格.xlsx"
	utils.MergeMuchExcelOneRepeatExcel(y, x, "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)

	dir1 := "D:\\桌面\\测试表格\\测试表格.xlsx"
	res, err := common.GetExcelAppointData(dir1, "Sheet1", "姓名", 2)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(res)
}
