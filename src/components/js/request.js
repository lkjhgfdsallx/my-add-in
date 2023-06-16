import axios from 'axios'

// 创建axios实例
const service = axios.create({
  // 公共接口
  baseURL: 'https://mptapi.fof99.com',
  // 超时时间,设置了30s的超时时间
  timeout: 30 * 1000,
})
// request拦截器
service.interceptors.request.use((config) => {
  // 发送请求前
  return config
}, (error) => {
  // 请求错误
  console.log(error) // 用于调试
  Promise.reject(error)
})
// respone拦截器
service.interceptors.response.use(
  (response) => {
    // 响应数据
    return response.data
  },
  (error) => {
    // 响应错误
    console.log(`err${error}`) // 用于调试
    return Promise.reject(error)
  },
)

// 导出
export default service
