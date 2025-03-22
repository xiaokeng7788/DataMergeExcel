package dataMergeExcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
)

// 判断路径是否存在
func PathExists(path string) bool {
	_, err := os.Stat(path)
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

// 创建指定的文件路径excel表格文件
func (e *Excel) CreatedExcelPath() error {
	if e.File == nil {
		return errors.New("暂无可执行文件")
	}
	if !PathExists(e.OutPath) {
		return os.ErrNotExist
	}
	if err := e.File.SaveAs(e.OutPath + "\\" + e.OutFile); err != nil {
		return fmt.Errorf("文件保存失败，错误原因为: %v, 请重试", err.Error())
	}
	if err := e.File.Close(); err != nil {
		return err
	}
	return nil
}
