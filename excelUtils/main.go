package excelUtils

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
)

// region 通用函数

// 判断路径是否存在
func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

// 读取规定工作表的数据
//
// filePaths excel文件路径
//
// sheetName 工作表名称
func GetExcelSheetData(filePaths, sheetName string) ([][]string, error) {
	if exists, _ := PathExists(filePaths); !exists {
		return nil, errors.New(filePaths + "文件不存在")
	}
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
	err = SheetNameExists(f, sheetName, false)
	if err != nil {
		return nil, err
	}
	// 获取 sheetName 上所有单元格
	return f.GetRows(sheetName)
}

// 创建一个Excel文件
//
// out 输出路径
// sheetName 工作表名称
// data 数据
// titleNum 表头数量
func CreateExcel(out, outFileName, sheetName string, data [][]string, titleNum int) error {
	if exists, _ := PathExists(out); !exists {
		return errors.New(out + "路径不存在")
	}
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	err := SheetNameExists(f, sheetName, true)
	if err != nil {
		return err
	}
	// 将数据写入Excel文件
	for k, item := range data {
		cell, err := excelize.CoordinatesToCellName(1, k+titleNum) // 从第titleNum行开始写入数据
		if err != nil {
			return err
		}
		row := make([]string, 0)
		row = item
		if err := f.SetSheetRow(sheetName, cell, &row); err != nil {
			return err
		}
	}
	// 根据指定路径保存文件
	if err := f.SaveAs(out + "\\" + outFileName); err != nil {
		return fmt.Errorf("文件保存失败，错误原因为: %v, 请重试", err.Error())
	}
	return nil
}

// 创建一个Excel文件 包含多个工作表
//
// out 输出路径
// sheetName 工作表名称
// data 数据
// titleNum 表头数量
func BatchCreateExcel(out, outFileName string, sheetName []string, data map[string][][]string, titleNum int) error {
	if exists, _ := PathExists(out); !exists {
		return errors.New(out + "路径不存在")
	}
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	if len(sheetName) != len(data) {
		return errors.New("指定表头和生成数据长度不一致")
	}
	// 判断表头和生成数据表头是否一致
	for k := range data {
		var exit bool
		for _, v := range sheetName {
			if k == v {
				exit = true
				break
			}
		}
		if !exit {
			return errors.New("生成数据中有不存在的表头数据")
		}
	}
	// 判断工作表是否存在
	for _, v := range sheetName {
		err := SheetNameExists(f, v, true)
		if err != nil {
			return err
		}
	}
	// 将数据写入Excel文件
	for k, item := range data {
		for i, v := range item {
			cell, err := excelize.CoordinatesToCellName(1, i+titleNum) // 从第titleNum行开始写入数据
			if err != nil {
				return err
			}
			row := make([]string, 0)
			row = v
			if err := f.SetSheetRow(k, cell, &row); err != nil {
				return err
			}
		}
	}
	// 根据指定路径保存文件
	if err := f.SaveAs(out + "\\" + outFileName); err != nil {
		return fmt.Errorf("文件保存失败，错误原因为: %v, 请重试", err.Error())
	}
	return nil
}

// 从给定的数据表格中读取数据 找到表头的最大长度 以及返回指定表头的下标
func GetExcelTitleInfo(data [][]string, title string, titleNum int) (max, index int, err error) {
	if titleNum != 0 {
		if len(data) < titleNum {
			return 0, 0, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	} else {
		titleNum = 3
		if len(data) < titleNum {
			return 0, 0, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	}
	// 寻找表头中最大的一行值 并循环寻找表头
	firstNum := len(data[0])
	appointNum := 0 // 指定表头的下标
	flag := false   // 标记是否已经找到了表头
	if title == "" {
		// 如果给定的表头为空 则默认为第一行
		flag = true
	}
	for i := 0; i < titleNum; i++ {
		if len(data[i]) > firstNum {
			firstNum = len(data[i])
		}
		if !flag {
			for k, v := range data[i] {
				if v == title {
					appointNum = k
					flag = true
				}
			}
		}
	}
	if !flag {
		return 0, 0, errors.New("表头不存在")
	}
	return firstNum, appointNum, nil
}

// 将字符串数字相加
func AddStringToInt(a, b string) string {
	aa, _ := strconv.Atoi(a)
	bb, _ := strconv.Atoi(b)
	return strconv.Itoa(aa + bb)
}

// 判断文件中是否包含指定的工作表
// sheetName : 工作表名称
// force: 是否强制创建工作表(false时如果不存在则返回错误, true时如果不存在则会创建, 如果创建失败则会返回错误)
func SheetNameExists(f *excelize.File, sheetName string, force bool) error {
	// 获取所有表格名称
	sheetNames := f.GetSheetList()
	for _, name := range sheetNames {
		if sheetName == name {
			return nil
		}
	}
	if !force {
		return fmt.Errorf("该文件中不存在工作表: %v", sheetName)
	}
	// 强制创建新的工作表
	_, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	return nil
}
