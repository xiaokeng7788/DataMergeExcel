package utils

import (
	"DataMergeExcel/common"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"sort"
	"strconv"
)

// 多个数据表合并成一个表格
//
// 1. 需要有个总表 就是原始数据表 里面应该包含分表的所有初始数据
//
// 2. 需要知道原始数据表和分表数据的位置
//
// 3. 需要知道操作的工作表位置和表头长度
//
// dir 文件夹路径 all 	原始数据表格文件名带有.xlsx		 titleNum 表头长度
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneExcel(dir, all, sheetName string, titleNum uint) {
	var allData map[string][]string
	pathMap := map[string]bool{}
	// 检查是否存在文件夹
	_, err2 := os.Stat(dir)
	if err2 != nil {
		fmt.Println("系统找不到指定文件，请先确定excel文件夹是否存在，并重试")
		return
	}
	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if info.IsDir() {
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
	if _, ok := pathMap[dir+"\\"+all]; !ok {
		fmt.Printf("未发现原始数据文件表，请检查文件夹内是否有--> %v 文件", all)
		return
	} else {
		fmt.Println("原始数据文件表已找到")
		// 首先确定主表数据 然后和分表数据匹配
		data, err := common.GetExcelIndexData(dir+"\\"+all, sheetName, titleNum)
		if err != nil {
			fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
			return
		}
		allData = data
		delete(pathMap, dir+"\\"+all) // 删除原始数据 确保不再数据读取
	}
	fmt.Println("---开始读取分表数据文件表---")
	for k := range pathMap {
		fmt.Println(k)
		data, err := common.GetExcelIndexData(dir+"\\"+all, sheetName, titleNum)
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
	fmt.Println("---分表数据文件表读取完成---")
	fmt.Println("---开始写入数据---")
	// 排序  因为map 是无序的
	var index []int
	for k := range allData {
		atoi, _ := strconv.Atoi(k)
		index = append(index, atoi)
	}
	sort.Ints(index)
	// 创建一个新的Excel文件
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 将数据写入Excel文件
	for k, item := range index {
		cell, err := excelize.CoordinatesToCellName(1, k+2) // 从第二行开始写入数据
		if err != nil {
			fmt.Println(err)
			return
		}
		row := make([]string, 0)
		row = append(row, allData[strconv.Itoa(item)]...)
		if err := f.SetSheetRow("Sheet1", cell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// 根据指定路径保存文件
	if err := f.SaveAs("./整合数据.xlsx"); err != nil {
		fmt.Printf("文件保存失败，错误原因为: %v, 请重试", err.Error())
		return
	}
	fmt.Println("↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓")
	fmt.Println("\n\n---写入数据完成---")
	fmt.Println("文件已经导出，请查看----> 整合数据.xlsx")
	fmt.Println("\n\n↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑")
}
