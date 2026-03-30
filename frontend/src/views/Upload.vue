<template>
  <div>
    <el-card>
      <template #header><span style="font-size: 16px; font-weight: bold">数据导入</span></template>
      <el-upload
        drag
        accept=".xlsx"
        :http-request="handleUpload"
        :show-file-list="false"
        :disabled="loading"
        style="width: 100%"
      >
        <el-icon style="font-size: 48px; color: #409EFF; margin-top: 20px"><UploadFilled /></el-icon>
        <div style="margin: 12px 0; color: #606266">
          拖拽 xlsx 文件至此处，或 <em style="color: #409EFF">点击上传</em>
        </div>
        <template #tip>
          <div style="color: #909399; font-size: 12px; margin-top: 8px; text-align: center">
            仅支持企微导出的上下班打卡日报 .xlsx 文件
          </div>
        </template>
      </el-upload>
    </el-card>

    <el-card style="margin-top: 16px; border-left: 4px solid #67c23a" shadow="never">
      <el-alert title="系统使用说明" type="success" :closable="false" show-icon style="margin-bottom: 12px">
        <div style="color: #303133; font-size: 13px; line-height: 1.9">
          本系统支持企微管理端导出“上下班打卡日报”后上传，一键生成报表。
        </div>
      </el-alert>
      <el-steps :active="3" finish-status="success" align-center>
        <el-step title="导出日报" description="企微管理端导出 .xlsx" />
        <el-step title="上传解析" description="上传后自动解析数据" />
        <el-step title="生成报表" description="考勤统计 / 周内加班" />
      </el-steps>
      <div style="margin-top: 12px; color: #606266; font-size: 13px">
        <span style="font-weight: 600">建议文件名：</span>
        <el-tag size="small" type="info">上下班打卡_日报_20260201-20260228（部门简称）.xlsx</el-tag>
      </div>
    </el-card>

    <el-card v-if="summary" style="margin-top: 16px">
      <template #header><span style="font-size: 16px; font-weight: bold">解析结果</span></template>
      <el-descriptions :column="2" border>
        <el-descriptions-item label="总人数">{{ summary.total_persons }} 人</el-descriptions-item>
        <el-descriptions-item label="日期范围">{{ summary.date_range }}</el-descriptions-item>
        <el-descriptions-item label="概况统计记录">{{ summary.total_rows_overview }} 条</el-descriptions-item>
        <el-descriptions-item label="打卡详情记录">{{ summary.total_rows_details }} 条</el-descriptions-item>
        <el-descriptions-item label="包含部门" :span="2">
          <el-tag v-for="d in summary.departments" :key="d" style="margin: 2px 4px 2px 0" type="primary">
            {{ d }}
          </el-tag>
        </el-descriptions-item>
      </el-descriptions>
      <div style="margin-top: 16px">
        <el-button type="primary" @click="$router.push('/preview')">
          前往数据预览 →
        </el-button>
        <el-button @click="$router.push('/export')" style="margin-left: 8px">
          前往报表导出 →
        </el-button>
      </div>
    </el-card>
  </div>
</template>

<script setup>
import { ref, inject } from 'vue'
import { ElMessage } from 'element-plus'
import { uploadFile } from '../api'

const loading = ref(false)
const summary = ref(null)
const setHasData = inject('setHasData')

async function handleUpload({ file }) {
  loading.value = true
  try {
    const { data } = await uploadFile(file)
    summary.value = data
    setHasData(true)
    ElMessage.success('上传解析成功')
  } catch (e) {
    ElMessage.error('上传失败：' + (e.response?.data?.detail || e.message))
  } finally {
    loading.value = false
  }
}
</script>
