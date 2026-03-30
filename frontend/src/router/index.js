// frontend/src/router/index.js
import { createRouter, createWebHistory } from 'vue-router'
import Upload from '../views/Upload.vue'
import DataPreview from '../views/DataPreview.vue'
import Export from '../views/Export.vue'

export default createRouter({
  history: createWebHistory(),
  routes: [
    { path: '/', redirect: '/upload' },
    { path: '/upload', component: Upload },
    { path: '/preview', component: DataPreview },
    { path: '/export', component: Export },
  ]
})
