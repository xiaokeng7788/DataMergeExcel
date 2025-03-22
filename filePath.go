package dataMergeExcel

import (
	"github.com/xuri/excelize/v2"
	"os"
)

// 判断路径是否存在
func PathExists(path string) bool {
	_, err := os.Stat(path) //os.Stat获取文件信息
	if err != nil {
		return os.IsExist(err)
	}
	return true
}

// 根据文件地址操作excel
func OpenExcelFile(filePath string) (*Excel, error) {
	if !PathExists(filePath) {
		return nil, os.ErrNotExist
	}
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	return &Excel{File: file}, nil
}

// 根据指定的工作簿找到对应的数
