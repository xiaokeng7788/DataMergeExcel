package bootstrap

import (
	"DataMergeExcel/router"
	"context"
	"errors"
	"fmt"
	"github.com/gin-gonic/gin"
	"log"
	"net/http"
	"os"
	"os/signal"
	"syscall"
	"time"
)

func RunServer() {
	var engine *gin.Engine
	gin.SetMode(gin.DebugMode)
	engine = gin.New()
	engine.Use(gin.Logger(), gin.Recovery())
	// 关闭控制台颜色
	gin.DisableConsoleColor()
	router.Init(engine)
	// 开启服务器
	srv := &http.Server{
		Addr:    ":" + "8099",
		Handler: engine,
	}
	go func() {
		if err := srv.ListenAndServe(); err != nil && !errors.Is(err, http.ErrServerClosed) {
			str := fmt.Sprintf("listen: %s\n", err) //拼接字符串
			fmt.Println(str)
		}
	}()
	log.Println("服务器启动成功")
	// 等待中断信号以优雅地关闭服务器（设置 5 秒的超时时间）
	quit := make(chan os.Signal)
	signal.Notify(quit, syscall.SIGINT, syscall.SIGTERM)
	<-quit
	fmt.Println("服务器正在关闭 ...")

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := srv.Shutdown(ctx); err != nil {
		str := fmt.Sprintf("服务器关闭： %s\n", err)
		fmt.Println(str)
	}
	fmt.Println("服务器已关闭")
}
