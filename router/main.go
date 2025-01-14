package router

import (
	"DataMergeExcel/controller"
	"github.com/gin-gonic/gin"
)

// 注册路由
func Init(r *gin.Engine) {
	r.GET("/", func(c *gin.Context) {
		c.String(200, "dev")
	})
	api := r.Group("/api")
	api.POST("excel/file", controller.HandlerExcelFile)
}
