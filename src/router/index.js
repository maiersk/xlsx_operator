import { createRouter, createWebHashHistory } from 'vue-router'

import Layout from '../layout/'

export const Routers = [
  {
    path: '/',
    component: Layout,
    redirect: '/test',
    children: [
      {
        path: '/test',
        name: 'testXlsx',
        component: () => import('../views/xlsx/')
      }
    ]
  },
]

export default createRouter({
  history: createWebHashHistory(),
  routes: Routers
})