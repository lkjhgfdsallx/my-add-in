import request from './request'

// 不带参示例
export function wxLogin() {
  return request({
    url: '/newgoapi/wx/qr/login',
    method: 'get',
  })
}
