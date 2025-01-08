package main

import "DataMergeExcel/utils"

func main() {
	dir := "D:\\桌面\\测试表格"
	utils.MergeMuchExcelOneExcel(dir, "测试表格.xlsx", "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)
	//dir := "D:\\桌面\\测试表格\\测试表格 - 副本 (2).xlsx"
	//res, err := common.GetExcelIndexData(dir, "Sheet1", 1)
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//fmt.Println(res)

}
