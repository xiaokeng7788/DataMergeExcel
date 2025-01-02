package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
)

func main() {

}

func GetExcelData(filePaths string) (map[string][]string, error) {
	fmt.Println(fmt.Sprintf("---开始读取 %v 数据文件表---\n", filePaths))
	f, err := excelize.OpenFile(filePaths)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			return
		}
	}()
	// 获取所有表格名称
	sheetNames := f.GetSheetList()
	if len(sheetNames) == 0 {
		return nil, errors.New("不存在工作表")
	}
	// 选择第一个表格名称
	firstSheet := sheetNames[0]
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows(firstSheet)
	if err != nil {
		return nil, err
	}
	res := make(map[string][]string) // key 序号 value 数据
	if len(rows) < 2 {
		return nil, errors.New("表内数据最少为三行")
	}
	firstNum := len(rows[0]) // 确定表头的数量
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}
		if i < 3 {
			//	 规定默认不处理前两行
			continue
		}
		v := make([]string, firstNum)
		for k, s := range row {
			v[k] = s
		}
		res[row[0]] = v
	}
	fmt.Printf("---%v 数据文件表读取完成---\n", filePaths)
	return res, nil
}
