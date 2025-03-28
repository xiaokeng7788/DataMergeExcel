package excelUtils

import (
	"strconv"
)

// 多个数据表合并成一个表格 只能处理以数字为唯一索引的表格
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneIndexExcel(allData map[string][]string, otherData map[string][]string) map[string][]string {
	for k := range allData {
		if _, ok := otherData[k]; ok {
			for i, v := range allData[k] {
				if v != otherData[k][i] && otherData[k][i] != "" && v == "" { // 保证以主表数据不变融合数据
					allData[k][i] = otherData[k][i]
				}
			}
		}
	}
	return allData
}

// 两个表数据融合 可以处理非数字唯一索引的表格
//
// 可以是两个相互独立的表 只要有相同索引就行
//
// 把两个表拥有相同索引的数据进行融合，产生新表 新表数据就会是两个表拥有共同列的拼接到一起 自行分辨左右两表数据
//
// x y 需要处理的文件地址	sheetName 工作表名	out导出文件路径		 titleNum 表头长度  flush 是否使用流式写入
//
// 正确处理会得到一个不含表头的合并之后的数据表
func MergeMuchExcelOneRepeatExcel(xData, yData map[string][][]string) [][]string {
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
	return result
}

// 将同一工作表中的具有相同索引的数据合并到一起 指的是表格中数值类型相加
//
// 除索引表头列数据可以是任意类型 其他的数据类型只能是数字类型
//
// filePaths 需要处理的文件地址		sheetName 工作表名	title 以哪个标题为索引		out导出文件路径		 titleNum 表头长度 flush 是否使用流式写入
func MergeWorkSheetData(index int, data map[string][][]string) [][]string {
	result := make([][]string, 0)
	for k, v := range data {
		if len(v) > 1 {
			r := make([]string, len(v[0]))
			for _, s := range v {
				for i, s1 := range s {
					if i == index {
						continue
					}
					r[i] = AddStringToInt(r[i], s1)
				}
			}
			r[index] = k
			result = append(result, r)
		} else {
			result = append(result, v...)
		}
	}
	return result
}

// 将同一工作表中的按照固定列拆分后 将相同列名的数据单独合并成一个表格 flush 是否使用流式写入

// 创建一个Excel文件 包含多个工作表
//
// out 输出路径
// sheetName 工作表名称
// data 数据
// titleNum 表头数量

// 将字符串数字相加
func AddStringToInt(a, b string) string {
	aa, _ := strconv.Atoi(a)
	bb, _ := strconv.Atoi(b)
	return strconv.Itoa(aa + bb)
}
