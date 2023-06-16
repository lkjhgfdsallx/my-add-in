import { createMemoryHistory, createRouter } from 'vue-router'
import Date from '../views/DateIndex.vue'
import Flushed from '../views/FlushedIndex.vue'
import Help from '../views/HelpIndex.vue'
import Insertion from '../views/InsertionIndex.vue'
import Login from '../views/LoginIndex.vue'
import Products from '../views/ProductsIndex.vue'

const routes = [
  {
    path: '/', // 当访问首页时，重定向跳转到login界面
    redirect: {
      name: 'Login',
    },
  },
  {
    path: '/date',
    name: 'Date',
    component: Date,
  },
  {
    path: '/flushed',
    name: 'Flushed',
    component: Flushed,
  },
  {
    path: '/help',
    name: 'Help',
    component: Help,
  },
  {
    path: '/insertion',
    name: 'Insertion',
    component: Insertion,
  },
  {
    path: '/login',
    name: 'Login',
    component: Login,
  },
  {
    path: '/products',
    name: 'Products',
    component: Products,
  },
]

const router = createRouter({
  history: createMemoryHistory(),
  routes,
})

export default router
