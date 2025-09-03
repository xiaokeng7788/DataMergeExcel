package dataMergeExcel

import (
	"bytes"
	"errors"
	"fmt"
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

// 将表格信息转换成二进制文件返回 供用户自己操作
func (e *Excel) ExportExcelBuffer() ([]byte, error) {
	if e.File == nil {
		return nil, errors.New("暂无可执行文件")
	}
	// 将数据保存到缓冲区
	buf := new(bytes.Buffer)
	if _, err := e.File.WriteTo(buf); err != nil {
		return nil, fmt.Errorf("文件保存失败，错误原因为: %v, 请重试", err.Error())
	}
	data := buf.Bytes()
	return data, nil
}
