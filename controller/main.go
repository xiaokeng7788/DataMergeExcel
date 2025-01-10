package controller

import (
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
		c.String(500, "上传失败")
		return
	}
	// 判断文件大小是否小于10MB
	if file.Size > 10*1024*1024 {
		c.String(500, "上传的文件大小超过10MB")
		return
	}
	// 判断文件是不是excel文件
	if file.Header.Get("Content-Type") != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" {
		c.String(500, "上传的文件不是excel文件")
		return
	}
	src, err := file.Open()
	if err != nil {
		c.String(500, "读取文件失败")
		return
	}
	defer func(src multipart.File) {
		err := src.Close()
		if err != nil {
			fmt.Println("关闭文件失败:", err)
		}
	}(src)
	f, err := excelize.OpenReader(src)
	if err != nil {
		fmt.Println("读取文件失败:", err)
		return
	}
	err = f.Close()
	if err != nil {
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println("读取Sheet1失败:", err)
		return
	}
	c.JSON(200, rows)
}
