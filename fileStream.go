package dataMergeExcel

import (
	"github.com/xuri/excelize/v2"
	"io"
)

// 根据文件流操作excel
func OpenExcelStream(fileStream io.Reader) (*Excel, error) {
	file, err := excelize.OpenReader(fileStream)
	if err != nil {
		return nil, err
	}
	return &Excel{File: file}, nil
}
