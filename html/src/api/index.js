//导入request.js 请求工具
import request from '@/utils/request'
import axios from 'axios'

//
export const getExcel = () => {
    return request.get('/excel')
}

export const uploadFile = (files) => {
    const data = new FormData()
    files.forEach(item => {
        data.append("file", item)
    });
    return axios.post("/api/excel/file", data, {
        headers: {
            "Content-Type": "multipart/form-data",
        }
    })
        .then(response => {
            //操作失败
            return response.data;
        })
        .catch(error => {
            //操作失败
            throw error;
        });
}