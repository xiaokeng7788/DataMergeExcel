//导入axios
import axios from 'axios'
import { ElMessage } from 'element-plus'
//定义一个变量,记录公共的前缀 baseURL
const baseURL = '/api';
const instance = axios.create({ baseURL })

//添加请求拦截器
instance.interceptors.request.use(
    (config) => {
        return config
    },
    (err) => {
        //如果请求错误做什么
        Promise.reject(err)
    }
)

//添加响应拦截器
instance.interceptors.response.use(
    result => {
        //判断业务状态码
        if (result.data.code === 200) {
            // console.log(result);
            return result.data;
        }
        //操作失败
        ElMessage.error(result.data.message ? result.data.message : '服务异常')
        //异步的状态转化成失败
        return Promise.reject(result.data)

    },
    err => {
        ElMessage.error('服务异常');
        return Promise.reject(err);//异步的状态转化成失败的状态
    }
)

export default instance;