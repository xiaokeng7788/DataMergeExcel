package dataMergeExcel

import (
	"bytes"
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

// 将数据保存到缓冲区 返回缓冲区数据
func (e *Excel) CreatedExcelBuffer() (int64, error) {
	// 将数据保存到缓冲区
	buf := new(bytes.Buffer)
	return e.File.WriteTo(buf)
}
