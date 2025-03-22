package excelUtils

import (
	"strconv"
)

// region 通用函数

// 创建一个Excel文件 包含多个工作表
//
// out 输出路径
// sheetName 工作表名称
// data 数据
// titleNum 表头数量
//func (e *HandleExcel) BatchCreateExcel(out, outFileName string, sheetName []string, data map[string][][]string, titleNum int) error {
//	if exists, _ := PathExists(out); !exists {
//		return errors.New(out + "路径不存在")
//	}
//	// 创建一个新的Excel文件
//	err := new(HandleExcel).NewCreateExcel()
//	if err != nil {
//		return err
//	}
//
//	if len(sheetName) != len(data) {
//		return errors.New("指定表头和生成数据长度不一致")
//	}
//	// 判断表头和生成数据表头是否一致
//	for k := range data {
//		var exit bool
//		for _, v := range sheetName {
//			if k == v {
//				exit = true
//				break
//			}
//		}
//		if !exit {
//			return errors.New("生成数据中有不存在的表头数据")
//		}
//	}
//	// 判断工作表是否存在
//	for _, v := range sheetName {
//		err := e.SheetNameExists(v, true)
//		if err != nil {
//			return err
//		}
//	}
//	// 将数据写入Excel文件
//	for k, item := range data {
//		for i, v := range item {
//			cell, err := excelize.CoordinatesToCellName(1, i+titleNum) // 从第titleNum行开始写入数据
//			if err != nil {
//				return err
//			}
//			row := make([]string, 0)
//			row = v
//			if err := e.File.SetSheetRow(k, cell, &row); err != nil {
//				return err
//			}
//		}
//	}
//	// 根据指定路径保存文件
//	if err = e.SaveExcel(out, outFileName); err != nil {
//		return err
//	}
//	return nil
//}

// 从给定的数据表格中读取数据 找到表头的最大长度 以及返回指定表头的下标
//func GetExcelTitleInfo(data [][]string, title string, titleNum int) (max, index int, err error) {
//	if titleNum != 0 {
//		if len(data) < titleNum {
//			return 0, 0, fmt.Errorf("表内数据最少为 %v 行", titleNum)
//		}
//	} else {
//		titleNum = 3
//		if len(data) < titleNum {
//			return 0, 0, fmt.Errorf("表内数据最少为 %v 行", titleNum)
//		}
//	}
//	// 寻找表头中最大的一行值 并循环寻找表头
//	firstNum := len(data[0])
//	appointNum := 0 // 指定表头的下标
//	flag := false   // 标记是否已经找到了表头
//	if title == "" {
//		// 如果给定的表头为空 则默认为第一行
//		flag = true
//	}
//	for i := 0; i < titleNum; i++ {
//		if len(data[i]) > firstNum {
//			firstNum = len(data[i])
//		}
//		if !flag {
//			for k, v := range data[i] {
//				if v == title {
//					appointNum = k
//					flag = true
//				}
//			}
//		}
//	}
//	if !flag {
//		return 0, 0, errors.New("表头不存在")
//	}
//	return firstNum, appointNum, nil
//}

// 将字符串数字相加
func AddStringToInt(a, b string) string {
	aa, _ := strconv.Atoi(a)
	bb, _ := strconv.Atoi(b)
	return strconv.Itoa(aa + bb)
}
