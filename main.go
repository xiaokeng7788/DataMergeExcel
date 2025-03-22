package dataMergeExcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
)

// 多个数据表合并成一个表格 只能处理以数字为唯一索引的表格
//
// 1. 需要有个总表 就是原始数据表 里面应该包含分表的所有初始数据
//
// 2. 需要知道原始数据表和分表数据的位置
//
// 3. 需要知道操作的工作表位置和表头长度
//
// dir 文件夹路径 all原始数据表格文件名带有.xlsx sheetName 工作表名		out导出文件路径		 titleNum 表头长度 flush 是否使用流式写入
//
// 正确处理会得到一个不含表头的合并之后的数据表
//func MergeMuchExcelOneIndexExcel(dir, all, sheetName, out string, titleNum int, flush bool) error {
//	// 检查是否存在文件夹
//	if exist, _ := excelUtils.PathExists(dir); !exist {
//		return errors.New("系统找不到指定文件，请先确定excel文件夹是否存在，并重试")
//	}
//	if exist, _ := excelUtils.PathExists(out); !exist {
//		return errors.New("系统找不到导出指定文件，请先确定导出文件路径是否存在")
//	}
//	var allData map[string][]string
//	pathMap := map[string]bool{}
//	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
//		if err != nil {
//			return err
//		}
//		if info.IsDir() {
//			// 只允许遍历到dir目录下的文件，不允许深入到子目录
//			if path != dir {
//				return filepath.SkipDir
//			}
//			return nil
//		}
//		if filepath.Ext(path) != ".xlsx" {
//			return nil
//		}
//		pathMap[path] = true
//		return nil
//	})
//	if err != nil {
//		return errors.New("文件读取失败，请重试")
//	}
//	if len(pathMap) < 2 {
//		return errors.New("并未发现需要合并的文件，请检查文件夹内是否有文件")
//	}
//	path := dir + "\\" + all // 原始数据表路径
//	if _, ok := pathMap[path]; !ok {
//		return fmt.Errorf("未发现原始数据文件表，请检查文件夹内是否有--> %v 文件", all)
//	} else {
//		// 首先确定主表数据 然后和分表数据匹配
//		data, err := excelUtils.GetExcelAppointIndexData(path, sheetName, "", titleNum)
//		if err != nil {
//			return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//		}
//		allData = data
//		delete(pathMap, path) // 删除原始数据 确保不再数据读取
//	}
//	for k := range pathMap {
//		data, err := excelUtils.GetExcelAppointIndexData(k, sheetName, "", titleNum)
//		if err != nil {
//			return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//		}
//		for k := range allData {
//			if _, ok := data[k]; ok {
//				for i, v := range allData[k] {
//					if v != data[k][i] && data[k][i] != "" {
//						allData[k][i] = data[k][i]
//					}
//				}
//			}
//		}
//	}
//	// 排序  因为map 是无序的
//	var index []int
//	for k := range allData {
//		atoi, _ := strconv.Atoi(k)
//		index = append(index, atoi)
//	}
//	sort.Ints(index)
//	result := make([][]string, 0)
//	for _, item := range index {
//		result = append(result, allData[strconv.Itoa(item)])
//	}
//	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, int(titleNum), flush); err != nil {
//		return errors.New("导出表格失败" + err.Error())
//	}
//	return nil
//}

// 两个表数据融合 可以处理非数字唯一索引的表格
//
// 可以是两个相互独立的表 只要有相同索引就行
//
// 把两个表拥有相同索引的数据进行融合，产生新表 新表数据就会是两个表拥有共同列的拼接到一起 自行分辨左右两表数据
//
// x y 需要处理的文件地址	sheetName 工作表名	out导出文件路径		 titleNum 表头长度  flush 是否使用流式写入
//
// 正确处理会得到一个不含表头的合并之后的数据表
//func MergeMuchExcelOneRepeatExcel(x, y, sheetName, title, out string, titleNum int, flush bool) error {
//	// 检查是否存在文件夹
//	if exist, _ := excelUtils.PathExists(out); !exist {
//		return errors.New("系统找不到指定文件->out，请先确定excel文件夹是否存在，并重试")
//	}
//	_, xData, err := excelUtils.GetExcelRepeatData(x, sheetName, title, titleNum)
//	if err != nil {
//		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//	}
//	_, yData, err := excelUtils.GetExcelRepeatData(y, sheetName, title, titleNum)
//	if err != nil {
//		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//	}
//	result := make([][]string, 0)
//	for k, v := range xData {
//		if _, ok := yData[k]; ok {
//			row := make([]string, 0)
//			for _, value := range v {
//				row = append(row, value...)
//			}
//			for _, value := range yData[k] {
//				row = append(row, value...)
//			}
//			result = append(result, row)
//		}
//	}
//	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, int(titleNum), flush); err != nil {
//		return fmt.Errorf("导出表格失败" + err.Error())
//	}
//	return nil
//}

// 将同一工作表中的具有相同索引的数据合并到一起 指的是表格中数值类型相加
//
// 除索引表头列数据可以是任意类型 其他的数据类型只能是数字类型
//
// filePaths 需要处理的文件地址		sheetName 工作表名	title 以哪个标题为索引		out导出文件路径		 titleNum 表头长度 flush 是否使用流式写入
//func MergeWorkSheetData(filePaths, sheetName, title, out string, titleNum int, flush bool) error {
//	index, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
//	if err != nil {
//		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//	}
//	result := make([][]string, 0)
//	for k, v := range res {
//		if len(v) > 1 {
//			r := make([]string, len(v[0]))
//			for _, s := range v {
//				for i, s1 := range s {
//					if i == index {
//						continue
//					}
//					r[i] = excelUtils.AddStringToInt(r[i], s1)
//				}
//			}
//			r[index] = k
//			result = append(result, r)
//		} else {
//			result = append(result, v...)
//		}
//	}
//	if err = excelUtils.CreateExcel(out, "整合数据.xlsx", sheetName, result, titleNum, flush); err != nil {
//		return errors.New("导出表格失败" + err.Error())
//	}
//	return nil
//}

// 将同一工作表中的按照固定列拆分后 将相同列名的数据单独合并成一个表格 flush 是否使用流式写入
//func MergeSameDataIntoNewTable(filePaths, sheetName, title, out string, titleNum int, flush bool) error {
//	_, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
//	if err != nil {
//		return fmt.Errorf("表格数据处理错误，错误原因为: %v\n", err.Error())
//	}
//	for s, v := range res {
//		if err = excelUtils.CreateExcel(out, s+".xlsx", sheetName, v, titleNum, flush); err != nil {
//			return errors.New("导出表格失败" + err.Error())
//		}
//	}
//	return nil
//}

// 重建 excel

type Excel struct {
	File      *excelize.File
	SheetName string // 工作表名
	Title     string // 表头 确定以那一列为索引
	TitleNum  int    // 表头长度
	FilePath  string // 原始文件路径
	OutPath   string // 输出文件路径
	OutFile   string // 输出文件名
}

const (
	// 默认工作簿名称
	DefaultSheetName = "Sheet1"
)

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
	e.SheetName = DefaultSheetName
	_, err := e.File.NewSheet(e.SheetName)
	if err != nil {
		return err
	}
	return nil
}

// 创建一个新的Excel文件
func (e *Excel) NewCreateExcel() error {
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	e.File = f
	return nil
}

// 读取规定工作表的数据
//
// filePaths excel文件路径
//
// sheetName 工作表名称
func (e *Excel) GetExcelSheetData() ([][]string, error) {
	if e.File == nil {
		return nil, errors.New("暂无可执行操作")
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
