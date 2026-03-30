// frontend/src/api/index.js
import axios from 'axios'

const apiBaseURL = import.meta.env.DEV ? 'http://localhost:8000/api' : '/api'
const api = axios.create({ baseURL: apiBaseURL })

export const uploadFile = (file) => {
  const form = new FormData()
  form.append('file', file)
  return api.post('/upload', form)
}

export const getOverview = (params) => api.get('/data/overview', { params })
export const getDetails = (params) => api.get('/data/details', { params })

export const exportAttendance = (dept) =>
  api.post('/export/attendance', { dept }, { responseType: 'blob' })

export const exportOvertime = (dept) =>
  api.post('/export/overtime', { dept }, { responseType: 'blob' })
