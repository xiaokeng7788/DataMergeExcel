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

// 多个数据表合并成一个表格 只能处理以数字为唯一索引的表格
//
// 1. 需要有个总表 就是原始数据表 里面应该包含分表的所有初始数据
//
// 2. 需要知道原始数据表和分表数据的位置
//
// 3. 需要知道操作的工作表位置和表头长度
//
// dir 文件夹路径 all 	原始数据表格文件名带有.xlsx		out导出文件路径		 titleNum 表头长度
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneExcel(dir, all, sheetName, out string, titleNum uint) {
	var allData map[string][]string
	pathMap := map[string]bool{}
	// 检查是否存在文件夹
	if exist, _ := common.PathExists(dir); !exist {
		fmt.Println("系统找不到指定文件，请先确定excel文件夹是否存在，并重试")
		return
	}
	if exist, _ := common.PathExists(out); !exist {
		fmt.Println("系统找不到导出指定文件，请先确定导出文件路径是否存在")
		return
	}
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
		data, err := common.GetExcelIndexData(path, sheetName, titleNum)
		if err != nil {
			fmt.Printf("表格数据处理错误，错误原因为: %v\n", err.Error())
			return
		}
		allData = data
		delete(pathMap, path) // 删除原始数据 确保不再数据读取
	}
	for k := range pathMap {
		data, err := common.GetExcelIndexData(k, sheetName, titleNum)
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
		if err := f.SetSheetRow(sheetName, cell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// 根据指定路径保存文件
	if err := f.SaveAs(out + "\\整合数据.xlsx"); err != nil {
		fmt.Printf("文件保存失败，错误原因为: %v, 请重试", err.Error())
		return
	}
	fmt.Println("↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓")
	fmt.Println("\n---写入数据完成---")
	fmt.Println("文件已经导出，请查看----> 整合数据.xlsx")
	fmt.Println("\n↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑")
}
