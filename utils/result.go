package utils

import "net/http"

// 错误返回值
type HttpResponse struct {
	Code    uint        `json:"code"`
	Message string      `json:"message"`
	Data    interface{} `json:"data"`
}

// 403响应 拒绝访问Http响应
func AccessDeniedHttpResponse(message string) (int, HttpResponse) {
	return http.StatusForbidden, HttpResponse{http.StatusForbidden, message, nil}
}

// 422响应 不可处理的实体Http响应
func UnprocessableEntityHttpResponse(message string) (int, HttpResponse) {
	return http.StatusUnprocessableEntity, HttpResponse{http.StatusUnprocessableEntity, message, nil}
}

// 200 响应 正确Http响应
func OkHttpResponse(data interface{}) (int, HttpResponse) {
	return http.StatusOK, HttpResponse{http.StatusOK, "响应成功", data}
}
