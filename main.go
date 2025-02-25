package dataMergeExcel

import (
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/excelUtils"
	"os"
	"path/filepath"
	"sort"
	"strconv"
)

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
func MergeMuchExcelOneIndexExcel(dir, all, sheetName, out string, titleNum int) {
	// 检查是否存在文件夹
	if exist, _ := excelUtils.PathExists(dir); !exist {
		fmt.Println("系统找不到指定文件，请先确定excel文件夹是否存在，并重试")
		return
	}
	if exist, _ := excelUtils.PathExists(out); !exist {
		fmt.Println("系统找不到导出指定文件，请先确定导出文件路径是否存在")
		return
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
		fmt.Println("文件读取失败，关闭此窗口重试")
		return
	}
	if len(pathMap) < 2 {
		fmt.Println("并未发现需要合并的文件，请检查文件夹内是否有文件")
		return
	}
	path := dir + "\\" + all // 原始数据表路径
	if _, ok := pathMap[path]; !ok {
		fmt.Printf("未发现原始数据文件表，请检查文件夹内是否有--> %v 文件", all)
		return
	} else {
		// 首先确定主表数据 然后和分表数据匹配
		data, err := excelUtils.GetExcelIndexData(path, sheetName, titleNum)
		if err != nil {
			fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
			return
		}
		allData = data
		delete(pathMap, path) // 删除原始数据 确保不再数据读取
	}
	for k := range pathMap {
		data, err := excelUtils.GetExcelIndexData(k, sheetName, titleNum)
		if err != nil {
			fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
			return
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
		fmt.Println("导出表格失败" + err.Error())
		return
	}
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
func MergeMuchExcelOneRepeatExcel(x, y, sheetName, title, out string, titleNum int) {
	// 检查是否存在文件夹
	if exist, _ := excelUtils.PathExists(out); !exist {
		fmt.Println("系统找不到指定文件->out，请先确定excel文件夹是否存在，并重试")
		return
	}
	_, xData, err := excelUtils.GetExcelRepeatData(x, sheetName, title, titleNum)
	if err != nil {
		fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
		return
	}
	_, yData, err := excelUtils.GetExcelRepeatData(y, sheetName, title, titleNum)
	if err != nil {
		fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
		return
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
		fmt.Println("导出表格失败" + err.Error())
		return
	}
}

// 将同一工作表中的具有相同索引的数据合并到一起 指的是表格中数值类型相加
//
// 除索引表头列数据可以是任意类型 其他的数据类型只能是数字类型
//
// filePaths 需要处理的文件地址		sheetName 工作表名	title 以哪个标题为索引		out导出文件路径		 titleNum 表头长度
func MergeWorkSheetData(filePaths, sheetName, title, out string, titleNum int) {
	index, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
	if err != nil {
		fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
		return
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
		fmt.Println("导出表格失败" + err.Error())
		return
	}
}

// 将同一工作表中的按照固定列拆分后 将相同列名的数据单独合并成一个表格
func MergeSameDataIntoNewTable(filePaths, sheetName, title, out string, titleNum int) {
	_, res, err := excelUtils.GetExcelRepeatData(filePaths, sheetName, title, titleNum)
	if err != nil {
		fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
		return
	}
	for s, v := range res {
		if err = excelUtils.CreateExcel(out, s+".xlsx", sheetName, v, titleNum); err != nil {
			fmt.Println("导出表格失败" + err.Error())
			return
		}
	}
}
