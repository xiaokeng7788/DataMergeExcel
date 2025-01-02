package main

import (
	"DataMergeExcel/common"
	"fmt"
)

func main() {
	dir := "D:\\桌面\\测试表格.xlsx"
	data, err := common.GetExcelIndexData(dir, "Sheet1", 2)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	fmt.Println(data)

}
