package excelUtils

// 处理多维索引数据
// 数据最终结构是 map[string][][]string{}

// region 处理多维索引数据

// 读取带有重复的数据
//
// 按照第一列为key key行数据为value 存储到数组二维切片中
//
// filePaths excel文件路径
//
// sheetName 工作表名称 如果为空则默认读取第一个工作表
//
// title 指定表头名称 以下标为key
//
// titleNum 表头数量 表头不能为0 默认为3
//func GetExcelRepeatData(filePaths, sheetName, title string, titleNum int) (index int, res map[string][][]string, err error) {
//	rows, err := GetExcelSheetData(filePaths, sheetName)
//	if err != nil {
//		return 0, nil, err
//	}
//	// 寻找表头中最大的一行值 并循环寻找表头
//	firstNum, appointNum, err := GetExcelTitleInfo(rows, title, titleNum)
//	if err != nil {
//		return 0, nil, err
//	}
//	toMap, err := ConvertToMap(rows, firstNum, titleNum, appointNum)
//	if err != nil {
//		return 0, nil, err
//	}
//	return appointNum, toMap, nil
//}

// 处理表中二维数据转换成直接可用map二维数组
//func ConvertToMap(data [][]string, firstNum, titleNum, appointNum int) (res map[string][][]string, err error) {
//	response := make(map[string][][]string)
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
//		response[row[appointNum]] = append(response[row[appointNum]], v)
//	}
//	return response, nil
//}
