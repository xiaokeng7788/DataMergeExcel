package dataMergeExcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

const (
	DefaultSheetName = "Sheet1" // 默认工作簿名称
)

type Excel struct {
	File      *excelize.File
	SheetName string // 工作表名
	Title     string // 表头 确定以那一列为索引
	TitleNum  int    // 表头长度
	FilePath  string // 原始文件路径
	OutPath   string // 输出文件路径
	OutFile   string // 输出文件名
}

// 判断表中是否有指定的工作簿 如果没有是否强制创建一个默认工作簿 默认工作簿名称为Sheet1
func (e *Excel) IsExitSheetName(force bool) error {
	sheetList := e.File.GetSheetList()
	for _, v := range sheetList {
		if v == e.SheetName {
			return nil
		}
	}
	if !force {
		return fmt.Errorf("该文件中不存在工作表: %v", e.SheetName)
	}
	// 默认强制创建的工作簿名称为Sheet1
	if e.SheetName == "" {
		e.SheetName = DefaultSheetName
	}
	_, err := e.File.NewSheet(e.SheetName)
	if err != nil {
		return err
	}
	return nil
}

// 创建一个新的Excel文件
func NewCreateExcel() *Excel {
	// 创建一个新的Excel文件
	return &Excel{File: excelize.NewFile()}
}

// 读取工作表的数据
func (e *Excel) GetExcelSheetData() ([][]string, error) {
	if e.File == nil {
		return nil, errors.New("暂无可执行文件")
	}
	// 获取所有表格名称
	err := e.IsExitSheetName(false)
	if err != nil {
		return nil, err
	}
	// 获取 sheetName 上所有单元格
	rows, err := e.File.GetRows(e.SheetName)
	if err := e.File.Close(); err != nil {
		return nil, err
	}
	return rows, nil
}

// 将数据写入表格
func (e *Excel) WriteExcelSheet(data [][]string) error {
	if e.File == nil {
		return errors.New("暂无可执行文件")
	}
	for rowIndex, item := range data {
		cell, err := excelize.CoordinatesToCellName(1, rowIndex+e.TitleNum) // 后续表头也需要写入
		if err != nil {
			return err
		}
		row := make([]string, 0)
		row = item
		if err := e.File.SetSheetRow(e.SheetName, cell, &row); err != nil {
			return err
		}
	}
	return nil
}

// 将数据流式写入表格
func (e *Excel) WriteExcelSheetStream(data [][]string) error {
	if e.File == nil {
		return errors.New("暂无可执行文件")
	}
	// 创建一个流式写入器
	streamWriter, err := e.File.NewStreamWriter(e.SheetName)
	if err != nil {
		return err
	}
	for rowIndex, item := range data {
		cell, err := excelize.CoordinatesToCellName(1, rowIndex)
		if err != nil {
			return err
		}
		interfaceItem := make([]interface{}, len(item))
		for i, v := range item {
			interfaceItem[i] = v
		}
		if err := streamWriter.SetRow(cell, interfaceItem); err != nil {
			return err
		}
	}
	// 关闭流式写入器
	if err := streamWriter.Flush(); err != nil {
		return err
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

// 处理表中二维数据转换成直接可用map一维数组
//
// needHeader 是否需要表头数据返回
func ConvertToOneDimension(data [][]string, firstNum, titleNum, appointNum int, needHeader bool) (res map[string][]string, header [][]string, err error) {
	response := make(map[string][]string)
	for i, row := range data {
		if len(row) == 0 {
			continue
		}
		if i < titleNum {
			if needHeader {
				header = append(header, row)
			}
			continue
		}
		if len(row) <= appointNum {
			return nil, nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头长度比该行数据长度还要多")
		}
		if row[appointNum] == "" {
			return nil, nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头该列数据不存在数据")
		}
		v := make([]string, firstNum)
		for k, s := range row {
			if k >= firstNum {
				break
			}
			v[k] = s
		}
		response[row[appointNum]] = v
	}
	return response, header, nil
}

// 处理表中二维数据转换成直接可用map二维数组
func ConvertToMultipleDimensions(data [][]string, firstNum, titleNum, appointNum int, needHeader bool) (res map[string][][]string, header [][]string, err error) {
	response := make(map[string][][]string)
	for i, row := range data {
		if len(row) == 0 {
			continue
		}
		if i < titleNum {
			if needHeader {
				header = append(header, row)
			}
			continue
		}
		if len(row) <= appointNum {
			return nil, nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头长度比该行数据长度还要多")
		}
		if row[appointNum] == "" {
			return nil, nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头该列数据不存在数据")
		}
		v := make([]string, firstNum)
		for k, s := range row {
			if k >= firstNum {
				break
			}
			v[k] = s
		}
		response[row[appointNum]] = append(response[row[appointNum]], v)
	}
	return response, header, nil
}

// 方便快速使用提供的指定文件路径获取表格数据
func GetExcelAppointIndexRepeatData(filePaths, sheetName, title string, titleNum int) (res map[string][][]string, err error) {
	excelFile, err := OpenExcelFile(filePaths)
	if err != nil {
		return nil, err
	}
	excelFile.SheetName = sheetName
	err = excelFile.IsExitSheetName(false)
	if err != nil {
		return nil, err
	}
	data, err := excelFile.GetExcelSheetData()
	if err != nil {
		return nil, err
	}
	firstNum, appointNum, err := GetExcelTitleInfo(data, title, titleNum)
	if err != nil {
		return nil, err
	}
	dimension, header, err := ConvertToMultipleDimensions(data, firstNum, titleNum, appointNum, true)
	if err != nil {
		return nil, err
	}
	dimension["header"] = header
	return dimension, nil
}
