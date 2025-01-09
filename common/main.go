package common

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
)

// 读取带有唯一索引的数据
//
// 按照序号为key key行数据为value 存储到数组切片中
//
// filePaths excel文件路径
//
// sheetName 工作表名称 如果为空则默认读取第一个工作表
//
// titleNum 表头数量 表头不能为0 默认为3
func GetExcelIndexData(filePaths, sheetName string, titleNum uint) (res map[string][]string, err error) {
	fmt.Println(fmt.Sprintf("---开始读取 %v 数据文件表---\n", filePaths))
	rows, err := getExcelSheetData(filePaths, sheetName)
	if err != nil {
		return nil, err
	}
	if titleNum != 0 {
		if len(rows) < int(titleNum) {
			return nil, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	} else {
		titleNum = 3
		if len(rows) < int(titleNum) {
			return nil, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	}
	// 寻找表头中最大的一行值
	firstNum := len(rows[0])
	for i := 0; i < int(titleNum); i++ {
		if len(rows[i]) > firstNum {
			firstNum = len(rows[i])
		}
	}
	response := make(map[string][]string)
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}
		if i < int(titleNum) {
			continue
		}
		v := make([]string, firstNum)
		for k, s := range row {
			if k >= firstNum {
				break
			}
			v[k] = s
		}
		response[row[0]] = v
	}
	fmt.Printf("---%v 数据文件表读取完成---\n", filePaths)
	return response, nil
}

// 读取带有重复的数据
//
// 按照第一列为key key行数据为value 存储到数组二维切片中
//
// filePaths excel文件路径
//
// sheetName 工作表名称 如果为空则默认读取第一个工作表
//
// titleNum 表头数量 表头不能为0 默认为3
func GetExcelRepeatData(filePaths, sheetName string, titleNum uint) (res map[string][][]string, err error) {
	fmt.Println(fmt.Sprintf("---开始读取 %v 数据文件表---\n", filePaths))
	rows, err := getExcelSheetData(filePaths, sheetName)
	if err != nil {
		return nil, err
	}
	if titleNum != 0 {
		if len(rows) < int(titleNum) {
			return nil, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	} else {
		titleNum = 3
		if len(rows) < int(titleNum) {
			return nil, fmt.Errorf("表内数据最少为 %v 行", titleNum)
		}
	}
	// 寻找表头中最大的一行值
	firstNum := len(rows[0])
	for i := 0; i < int(titleNum); i++ {
		if len(rows[i]) > firstNum {
			firstNum = len(rows[i])
		}
	}
	response := make(map[string][][]string) // key 序号 value 数据
	for i, row := range rows {
		if len(row) == 0 {
			continue
		}
		if i < int(titleNum) {
			continue
		}
		v := make([]string, firstNum)
		for k, s := range row {
			if k >= firstNum {
				break
			}
			v[k] = s
		}
		response[row[0]] = append(response[row[0]], v)
	}
	fmt.Printf("---%v 数据文件表读取完成---\n", filePaths)
	return response, nil
}

// 读取规定工作表的数据
//
// filePaths excel文件路径
//
// sheetName 工作表名称
func getExcelSheetData(filePaths, sheetName string) ([][]string, error) {
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
		return nil, errors.New("该文件中不存在工作表")
	}
	if sheetName == "" {
		sheetName = sheetNames[0]
	} else {
		var exit bool
		for _, name := range sheetNames {
			if sheetName == name {
				exit = true
				break
			}
		}
		if !exit {
			return nil, fmt.Errorf("该文件中不存在工作表: %v", sheetName)
		}
	}

	// 获取 sheetName 上所有单元格
	return f.GetRows(sheetName)
}

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

// 创建一个Excel文件
//
// out 输出路径
// sheetName 工作表名称
// data 数据
// titleNum 表头数量
func CreateExcel(out, sheetName string, data [][]string, titleNum int) error {
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
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
	if err := f.SaveAs(out + "\\整合数据.xlsx"); err != nil {
		return fmt.Errorf("文件保存失败，错误原因为: %v, 请重试", err.Error())
	}
	fmt.Println("↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓")
	fmt.Println("\n---写入数据完成---")
	fmt.Println("文件已经导出，请查看----> 整合数据.xlsx")
	fmt.Println("\n↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑")
	return nil
}
