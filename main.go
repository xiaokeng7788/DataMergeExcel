package dataMergeExcel

import (
	"errors"
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/excelUtils"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"sort"
	"strconv"
)

const (
	DefaultSheetName = "Sheet1" // 默认工作簿名称
)

type Excel struct {
	File      *excelize.File
	SheetName string // 工作表名
	Title     string // 表头 确定以那一列为索引
	FilePath  string // 原始文件路径
	FileName  string // 原始文件名
	OutPath   string // 输出文件路径
	OutFile   string // 输出文件名
	TitleNum  int    // 表头长度 以此确定表头的最大长度
}

// 判断表中是否有指定的工作簿 如果没有是否强制创建一个默认工作簿 默认工作簿名称为Sheet1
//
// force 是否强制创建工作簿
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

// 为创建新的可执行文件配置一些导出地址信息
//
// OutPath: 输出文件路径
// OutFileName: 输出文件名
func (e *Excel) SetExportConfig(OutPath, OutFileName string) {
	if OutPath != "" && e.OutPath != OutPath {
		e.OutPath = OutPath
	}
	if OutFileName != "" && e.OutFile != OutFileName {
		e.OutFile = OutFileName
	}
}

// 为创建新的可执行文件配置一些导入地址信息
//
// FilePath: 原始文件路径
// FileName: 原始文件名
func (e *Excel) SetImportConfig(FilePath, FileName string) {
	if FilePath != "" && e.FilePath != FilePath {
		e.FilePath = FilePath
	}
	if FileName != "" && e.FileName != FileName {
		e.FileName = FileName
	}
}

// 为创建新的可执行文件配置一些表格信息
//
// SheetName: 工作表名称
// Title: 表头名称
func (e *Excel) SetSheetConfig(SheetName, Title string) {
	if SheetName != "" && e.SheetName != SheetName && SheetName != DefaultSheetName {
		e.File.NewSheet(SheetName)
		e.SheetName = SheetName
	} else {
		e.SheetName = DefaultSheetName
	}
	if Title != "" && e.Title != Title {
		e.Title = Title
	}
}

// 判断导出的文件所需的参数是否满足
func (e *Excel) IsExportConfig() error {
	if e.File == nil {
		return errors.New("暂未初始化可执行文件")
	}
	if e.OutPath == "" {
		return errors.New("导出文件路径不能为空")
	}
	if e.OutFile == "" {
		return errors.New("导出文件名不能为空")
	}
	if e.SheetName == "" {
		return errors.New("工作表名称不能为空")
	}
	if !PathExists(filepath.Join(e.OutPath, e.OutFile)) {
		return errors.New("导出文件路径不存在，请先创建该路径")
	}
	return nil
}

// 读取工作表的数据
func (e *Excel) GetExcelSheetData() ([][]string, error) {
	if e.File == nil {
		return nil, errors.New("暂无可执行文件")
	}
	// 判断工作簿中是否有
	err := e.IsExitSheetName(false)
	if err != nil {
		return nil, err
	}
	// 获取 sheetName 上所有单元格
	rows, err := e.File.GetRows(e.SheetName)
	if err != nil {
		return nil, err
	}
	if err = e.File.Close(); err != nil {
		return nil, err
	}
	return rows, nil
}

// 将数据写入表格并导出文件
func (e *Excel) WriteExcelSheet(data [][]string) error {
	if e.File == nil {
		return errors.New("暂无可执行文件")
	}
	for rowIndex, item := range data {
		cell, err := excelize.CoordinatesToCellName(1, rowIndex+1)
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
		cell, err := excelize.CoordinatesToCellName(1, rowIndex+1)
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
	if err = streamWriter.Flush(); err != nil {
		return err
	}
	return nil
}

// 将表格信息导出到文件
func (e *Excel) ExportExcel() error {
	if err := e.IsExportConfig(); err != nil {
		return err
	}
	if err := e.File.SaveAs(filepath.Join(e.OutPath, e.OutFile)); err != nil {
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
func GetExcelAppointIndexRepeatData(filePaths, fileName, sheetName, title string, titleNum int) (res map[string][][]string, err error) {
	excelFile := NewCreateExcel()
	excelFile.SetImportConfig(filePaths, fileName)
	err = excelFile.OpenExcelFile()
	if err != nil {
		return nil, err
	}
	excelFile.SetSheetConfig(sheetName, title)
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

// 多个数据表合并成一个表格 只能处理以数字为唯一索引的表格
//
// 1. 需要有个总表 就是原始数据表 里面应该包含分表的所有初始数据
//
// 2. 需要知道原始数据表和分表数据的位置
//
// 3. 需要知道操作的工作表位置和表头长度
//
// dir 文件夹路径 all原始数据表格文件名带有.xlsx sheetName 工作表名		out导出文件路径		 titleNum 表头长度
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneIndexExcel(dir, all, sheetName, out string, titleNum int) error {
	// 检查是否存在文件夹
	if exist, _ := excelUtils.PathExists(dir); !exist {
		return errors.New("系统找不到指定文件，请先确定excel文件夹是否存在，并重试")
	}
	if exist, _ := excelUtils.PathExists(out); !exist {
		return errors.New("系统找不到导出指定文件，请先确定导出文件路径是否存在")
	}
	var allData map[string][]string
	pathMap := map[string]bool{}
	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			// 只允许遍历到dir目录下的文件，不允许深入到子目录
			if path != dir {
				return filepath.SkipDir
			}
			return nil
		}
		if filepath.Ext(path) != ".xlsx" {
			return nil
		}
		pathMap[path] = true
		return nil
	})
	if err != nil {
		return errors.New("文件读取失败，请重试")
	}
	if len(pathMap) < 2 {
		return errors.New("并未发现需要合并的文件，请检查文件夹内是否有文件")
	}
	path := dir + "\\" + all // 原始数据表路径
	if _, ok := pathMap[path]; !ok {
		return fmt.Errorf("未发现原始数据文件表，请检查文件夹内是否有--> %v 文件", all)
	} else {
		// 首先确定主表数据 然后和分表数据匹配
		data, err := excelUtils.GetExcelIndexData(path, sheetName, titleNum)
		if err != nil {
			return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
		}
		allData = data
		delete(pathMap, path) // 删除原始数据 确保不再数据读取
	}
	for k := range pathMap {
		data, err := excelUtils.GetExcelIndexData(k, sheetName, titleNum)
		if err != nil {
			return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
		}
		for k := range allData {
			if _, ok := data[k]; ok {
				for i, v := range allData[k] {
					if v != data[k][i] && data[k][i] != "" {
						allData[k][i] = data[k][i]
					}
				}
			}
		}
	}
	// 排序  因为map 是无序的
	var index []int
	for k := range allData {
		atoi, _ := strconv.Atoi(k)
		index = append(index, atoi)
	}
	sort.Ints(index)
	result := make([][]string, 0)
	for _, item := range index {
		result = append(result, allData[strconv.Itoa(item)])
	}
	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, int(titleNum)); err != nil {
		return errors.New("导出表格失败" + err.Error())
	}
	return nil
}

// 两个表数据融合 可以处理非数字唯一索引的表格
//
// 可以是两个相互独立的表 只要有相同索引就行
//
// 把两个表拥有相同索引的数据进行融合，产生新表 新表数据就会是两个表拥有共同列的拼接到一起 自行分辨左右两表数据
//
// x y 需要处理的文件地址	sheetName 工作表名	out导出文件路径		 titleNum 表头长度
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneRepeatExcel(x, y, sheetName, title, out string, titleNum int) error {
	// 检查是否存在文件夹
	if exist, _ := excelUtils.PathExists(out); !exist {
		return errors.New("系统找不到指定文件->out，请先确定excel文件夹是否存在，并重试")
	}
	_, xData, err := excelUtils.GetExcelRepeatData(x, sheetName, title, titleNum)
	if err != nil {
		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
	}
	_, yData, err := excelUtils.GetExcelRepeatData(y, sheetName, title, titleNum)
	if err != nil {
		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
	}
	result := make([][]string, 0)
	for k, v := range xData {
		if _, ok := yData[k]; ok {
			row := make([]string, 0)
			for _, value := range v {
				row = append(row, value...)
			}
			for _, value := range yData[k] {
				row = append(row, value...)
			}
			result = append(result, row)
		}
	}
	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, int(titleNum)); err != nil {
		return fmt.Errorf("导出表格失败" + err.Error())
	}
	return nil
}

// 将同一工作表中的具有相同索引的数据合并到一起 指的是表格中数值类型相加
//
// 除索引表头列数据可以是任意类型 其他的数据类型只能是数字类型
//
// filePaths 需要处理的文件地址		sheetName 工作表名	title 以哪个标题为索引		out导出文件路径		 titleNum 表头长度
func MergeWorkSheetData(filePaths, sheetName, title, out string, titleNum int) error {
	index, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
	if err != nil {
		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
	}
	result := make([][]string, 0)
	for k, v := range res {
		if len(v) > 1 {
			r := make([]string, len(v[0]))
			for _, s := range v {
				for i, s1 := range s {
					if i == index {
						continue
					}
					r[i] = excelUtils.AddStringToInt(r[i], s1)
				}
			}
			r[index] = k
			result = append(result, r)
		} else {
			result = append(result, v...)
		}
	}
	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, titleNum); err != nil {
		return errors.New("导出表格失败" + err.Error())
	}
	return nil
}

// 将同一工作表中的按照固定列拆分后 将相同列名的数据单独合并成一个表格
func MergeSameDataIntoNewTable(filePaths, sheetName, title, out string, titleNum int) error {
	_, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
	if err != nil {
		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
	}
	for s, v := range res {
		if err = excelUtils.CreateExcel(out, s+".xlsx", sheetName, v, titleNum); err != nil {
			return errors.New("导出表格失败" + err.Error())
		}
	}
	return nil
}
