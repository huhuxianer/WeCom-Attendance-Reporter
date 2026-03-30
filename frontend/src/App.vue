<template>
  <el-container style="height: 100vh">
    <el-aside width="200px" style="background: #304156">
      <div style="color: #fff; padding: 20px 16px; font-size: 16px; font-weight: bold; border-bottom: 1px solid #3d4f60">
        考勤报表生成系统
      </div>
      <el-menu
        router
        background-color="#304156"
        text-color="#bfcbd9"
        active-text-color="#409EFF"
        :default-active="$route.path"
      >
        <el-menu-item index="/upload">
          <el-icon><Upload /></el-icon>
          <span>数据导入</span>
        </el-menu-item>
        <el-menu-item index="/preview" :disabled="!hasData">
          <el-icon><DataAnalysis /></el-icon>
          <span>数据预览</span>
        </el-menu-item>
        <el-menu-item index="/export" :disabled="!hasData">
          <el-icon><Download /></el-icon>
          <span>报表导出</span>
        </el-menu-item>
      </el-menu>
    </el-aside>
    <el-main style="padding: 24px; background: #f0f2f5; overflow-y: auto">
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
import { ref, provide } from 'vue'

const hasData = ref(false)
provide('hasData', hasData)
provide('setHasData', (val) => { hasData.value = val })
</script>
