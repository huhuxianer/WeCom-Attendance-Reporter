<template>
  <el-card>
    <template #header>
      <span style="font-size: 16px; font-weight: bold">数据预览</span>
    </template>

    <el-tabs v-model="activeTab" @tab-change="onTabChange">
      <el-tab-pane label="概况统计" name="overview" />
      <el-tab-pane label="打卡详情" name="details" />
    </el-tabs>

    <!-- 筛选栏 -->
    <el-row :gutter="12" style="margin-bottom: 16px">
      <el-col :span="6">
        <el-input
          v-model="filters.keyword"
          placeholder="搜索姓名"
          clearable
          @clear="fetchData"
        />
      </el-col>
      <el-col :span="6">
        <el-input
          v-model="filters.dept"
          placeholder="部门筛选"
          clearable
          @clear="fetchData"
        />
      </el-col>
      <el-col :span="6" v-if="activeTab === 'details'">
        <el-input
          v-model="filters.date"
          placeholder="日期（如 2026/01）"
          clearable
          @clear="fetchData"
        />
      </el-col>
      <el-col :span="3">
        <el-button type="primary" @click="doSearch">查询</el-button>
      </el-col>
    </el-row>

    <!-- 表格 -->
    <el-table
      :data="tableData"
      v-loading="loading"
      border
      stripe
      size="small"
      style="width: 100%"
      height="500"
    >
      <el-table-column
        v-for="col in columns"
        :key="col"
        :prop="col"
        :label="col"
        :width="colWidth(col)"
        show-overflow-tooltip
      />
    </el-table>

    <!-- 分页 -->
    <div style="margin-top: 12px; display: flex; justify-content: flex-end">
      <el-pagination
        layout="total, prev, pager, next"
        :total="total"
        :page-size="pageSize"
        v-model:current-page="currentPage"
        @current-change="fetchData"
      />
    </div>
  </el-card>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { getOverview, getDetails } from '../api'
import { ElMessage } from 'element-plus'

const activeTab = ref('overview')
const loading = ref(false)
const tableData = ref([])
const columns = ref([])
const total = ref(0)
const currentPage = ref(1)
const pageSize = 50
const filters = ref({ keyword: '', dept: '', date: '' })

async function fetchData() {
  loading.value = true
  try {
    const params = { page: currentPage.value, page_size: pageSize }
    if (filters.value.keyword) params.keyword = filters.value.keyword
    if (filters.value.dept) params.dept = filters.value.dept

    let resp
    if (activeTab.value === 'overview') {
      resp = await getOverview(params)
    } else {
      if (filters.value.date) params.date = filters.value.date
      resp = await getDetails(params)
    }

    tableData.value = resp.data.items
    total.value = resp.data.total
    if (resp.data.items.length > 0) {
      columns.value = Object.keys(resp.data.items[0])
    }
  } catch (e) {
    ElMessage.error('加载失败：' + (e.response?.data?.detail || e.message))
  } finally {
    loading.value = false
  }
}

function doSearch() {
  currentPage.value = 1
  fetchData()
}

function onTabChange() {
  currentPage.value = 1
  filters.value = { keyword: '', dept: '', date: '' }
  fetchData()
}

function colWidth(col) {
  if (col === '部门') return 240
  if (col === '日期') return 180
  if (col.includes('时间')) return 130
  if (col.includes('地点')) return 200
  return 100
}

onMounted(fetchData)
</script>
