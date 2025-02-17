package main

import (
	"DataMergeExcel/utils"
)

func main() {
	dir := "D:\\桌面\\test\\test.xlsx"
	//dir := "D:\\桌面\\测试表格"
	//utils.MergeMuchExcelOneIndexExcel(dir, "测试表格.xlsx", "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)
	//
	//x := "D:\\桌面\\测试表格\\重复表格\\测试表格-1227.xlsx"
	//y := "D:\\桌面\\测试表格\\重复表格\\测试表格.xlsx"
	//utils.MergeMuchExcelOneRepeatExcel(y, x, "Sheet1", "D:\\桌面\\测试表格\\生成表格", 1)

	utils.MergeSameDataIntoNewTable(dir, "Sheet1", "标题", "D:\\桌面\\test", 1)

	//f := "D:\\桌面\\新建 文本文档.txt"
	////	读取文件
	//file, err := os.ReadFile(f)
	//if err != nil {
	//	panic(err)
	//}
	//split := strings.Split(string(file), "\n")
	//for _, v := range split {
	//	compare := strings.Index(v, "$")
	//	fmt.Println(v[:compare], v[compare+1:])
	//}
}
