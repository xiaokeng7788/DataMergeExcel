package controller

import (
	"DataMergeExcel/utils"
	"errors"
	"fmt"
	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
	"mime/multipart"
)

// 通用处理文件
func HandlerExcelFile(c *gin.Context) {
	// 接收上传的文件
	file, err := c.FormFile("file")
	if err != nil {
		c.JSON(utils.UnprocessableEntityHttpResponse(err.Error()))
		return
	}
	ger, err := GetExcelSheetData(file, "")
	if err != nil {
		c.JSON(utils.AccessDeniedHttpResponse(err.Error()))
		return
	}
	c.JSON(utils.OkHttpResponse(ger))
}

// 将文件直接转换成可操作的数据格式
func GetExcelSheetData(files *multipart.FileHeader, sheetName string) ([][]string, error) {
	// 判断文件大小是否小于10MB
	if files.Size > 10*1024*1024 {
		return nil, fmt.Errorf("文件大小不能超过10MB")
	}
	// 判断文件是不是excel文件
	if files.Header.Get("Content-Type") != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" {
		return nil, fmt.Errorf("文件不是excel文件")
	}
	src, err := files.Open()
	if err != nil {
		return nil, fmt.Errorf("文件读取失败," + err.Error())
	}
	defer func(src multipart.File) {
		err := src.Close()
		if err != nil {
			fmt.Println("关闭文件失败:", err)
			return
		}
	}(src)
	f, err := excelize.OpenReader(src)
	if err != nil {
		return nil, fmt.Errorf("文件读取失败," + err.Error())
	}
	err = f.Close()
	if err != nil {
		return nil, fmt.Errorf("文件读取失败," + err.Error())
	}
	// 获取所有表格名称
	sheetNames := f.GetSheetList()
	if len(sheetNames) == 0 {
		return nil, errors.New("该文件中不存在工作表")
	}
	if sheetName == "" {
		sheetName = sheetNames[0]
	} else {
		var exit bool
		for _, name := range sheetNames {
			if sheetName == name {
				exit = true
				break
			}
		}
		if !exit {
			return nil, fmt.Errorf("该文件中不存在工作表: %v", sheetName)
		}
	}
	return f.GetRows(sheetName)
}
