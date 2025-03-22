package excelUtils

// 处理唯一索引数据
// 唯一索引的最终数据结构是 map[string][]string{}

// region 处理唯一索引数据

// 读取自定义的数据表格数据 指定表头必须是唯一键
//
// 给定一个数据表 指定以那一列为固定表头长度 指定以哪一个表头名称为索引构建数据
//func GetExcelAppointIndexData(filePaths, sheetName, title string, titleNum int) (res map[string][]string, err error) {
//	rows, err := GetExcelSheetData(filePaths, sheetName)
//	if err != nil {
//		return nil, err
//	}
//
//	// 寻找表头中最大的一行值 并循环寻找表头
//	firstNum, appointNum, err := GetExcelTitleInfo(rows, title, titleNum)
//	if err != nil {
//		return nil, err
//	}
//	return ConvertToMapOne(rows, firstNum, titleNum, appointNum)
//}

// 读取自定义数据表格 指定表头可以不是唯一键
//
// 给定一个数据表 指定以那一列为固定表头长度 指定以哪一个表头名称为索引构建数据
//func GetExcelAppointIndexRepeatData(filePaths, sheetName, title string, titleNum int) (res map[string][][]string, err error) {
//	rows, err := GetExcelSheetData(filePaths, sheetName)
//	if err != nil {
//		return nil, err
//	}
//	// 寻找表头中最大的一行值 并循环寻找表头
//	firstNum, appointNum, err := GetExcelTitleInfo(rows, title, titleNum)
//	if err != nil {
//		return nil, err
//	}
//	return ConvertToMap(rows, firstNum, titleNum, appointNum)
//}

// 处理表中二维数据转换成直接可用map一维数组
//func ConvertToMapOne(data [][]string, firstNum, titleNum, appointNum int) (res map[string][]string, err error) {
//	response := make(map[string][]string)
//	for i, row := range data {
//		if len(row) == 0 {
//			continue
//		}
//		if i < titleNum {
//			continue
//		}
//		if len(row) <= appointNum {
//			return nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头长度比该行数据长度还要多")
//		}
//		if row[appointNum] == "" {
//			return nil, errors.New("第" + strconv.Itoa(i+1) + "行数据存在问题,指定表头该列数据不存在数据")
//		}
//		v := make([]string, firstNum)
//		for k, s := range row {
//			if k >= firstNum {
//				break
//			}
//			v[k] = s
//		}
//		response[row[appointNum]] = v
//	}
//	return response, nil
//}
