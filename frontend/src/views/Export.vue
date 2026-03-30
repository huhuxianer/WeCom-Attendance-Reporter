<template>
  <div>
    <el-row :gutter="24">
      <el-col :span="12">
        <el-card>
          <template #header>
            <span style="font-size: 16px; font-weight: bold">考勤数据统计报表</span>
          </template>
          <el-form label-width="60px" style="margin-bottom: 16px">
            <el-form-item label="部门">
              <el-input
                v-model="attendanceDept"
                placeholder="留空导出全部，或输入部门关键词"
                clearable
              />
            </el-form-item>
          </el-form>
          <p style="color: #909399; font-size: 13px; margin-bottom: 16px">
            格式参考：1月考勤数据统计模版，每人2行（上午/下午），包含31天符号及月度汇总
          </p>
          <el-button type="primary" :loading="loadingAttendance" @click="doExport('attendance')">
            <el-icon style="margin-right: 4px"><Download /></el-icon>
            导出 xlsx
          </el-button>
        </el-card>
      </el-col>

      <el-col :span="12">
        <el-card>
          <template #header>
            <span style="font-size: 16px; font-weight: bold">周内加班统计报表</span>
          </template>
          <el-form label-width="60px" style="margin-bottom: 16px">
            <el-form-item label="部门">
              <el-input
                v-model="overtimeDept"
                placeholder="留空导出全部，或输入部门关键词"
                clearable
              />
            </el-form-item>
          </el-form>
          <p style="color: #909399; font-size: 13px; margin-bottom: 16px">
            格式参考：1月周内加班统计模版，每人2行（20:00-22:00/22:00之后），包含月度次数统计
          </p>
          <el-button type="primary" :loading="loadingOvertime" @click="doExport('overtime')">
            <el-icon style="margin-right: 4px"><Download /></el-icon>
            导出 xlsx
          </el-button>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { exportAttendance, exportOvertime } from '../api'
import { ElMessage } from 'element-plus'

const attendanceDept = ref('')
const overtimeDept = ref('')
const loadingAttendance = ref(false)
const loadingOvertime = ref(false)

async function doExport(type) {
  const isAttendance = type === 'attendance'
  const loading = isAttendance ? loadingAttendance : loadingOvertime
  const dept = (isAttendance ? attendanceDept.value : overtimeDept.value) || null
  const fn = isAttendance ? exportAttendance : exportOvertime
  const label = isAttendance ? '考勤数据统计' : '周内加班统计'
  const deptLabel = dept || '全部'

  loading.value = true
  try {
    const resp = await fn(dept)
    
    // Extract filename from the content-disposition header if available
    let filename = `${deptLabel}_2026年1月${label}.xlsx` // fallback
    const disposition = resp.headers['content-disposition']
    if (disposition) {
      const utf8FilenameMatch = disposition.match(/filename\*=utf-8''([^;]+)/i)
      if (utf8FilenameMatch && utf8FilenameMatch[1]) {
        filename = decodeURIComponent(utf8FilenameMatch[1])
      } else {
        const filenameMatch = disposition.match(/filename="?([^";]+)"?/i)
        if (filenameMatch && filenameMatch[1]) {
          filename = decodeURIComponent(filenameMatch[1])
        }
      }
    }

    const blob = new Blob([resp.data], { type: resp.headers['content-type'] })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    URL.revokeObjectURL(url)
    ElMessage.success('导出成功')
  } catch (e) {
    ElMessage.error('导出失败：' + (e.response?.data?.detail || e.message))
  } finally {
    loading.value = false
  }
}
</script>
