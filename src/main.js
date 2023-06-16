import { createApp } from 'vue'
import App from './App.vue'

import router from './router' // 引入vue-router

import { dateFunction, flushedFunction, helpFunction, insertFunction, loginFunction, productsFunction } from './components/js/dialog'

window.Office.onReady((reason) => {
  createApp(App).use(router).mount('#app')
  console.log(reason)
})

window.Office.actions.associate('insertFunction', insertFunction)
window.Office.actions.associate('productsFunction', productsFunction)
window.Office.actions.associate('dateFunction', dateFunction)
window.Office.actions.associate('helpFunction', helpFunction)
window.Office.actions.associate('loginFunction', loginFunction)
window.Office.actions.associate('flushedFunction', flushedFunction)
