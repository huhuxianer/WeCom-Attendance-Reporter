# 考勤报表生成系统 | WeCom Attendance Automation

企业微信考勤导出数据的可视化与考勤统计报表自动生成工具。上传企微管理端导出的上下班打卡日报 `.xlsx` 文件，一键生成标准格式的考勤数据统计报表和周内加班统计报表，可作为加班费计算辅助工具。输入输出文件见 examples 中的示例文件。

> WeCom / DingTalk attendance report automation — upload punch-clock daily reports, generate formatted attendance & overtime Excel sheets in one click.

## 功能特性

- **数据导入** — 拖拽上传企微/钉钉导出的打卡日报 Excel 文件，自动解析
- **数据预览** — 分页查看考勤概况和打卡详情，支持按姓名、部门、日期筛选
- **考勤报表导出** — 基于模版生成考勤数据统计报表（含明细 + 汇总两个 Sheet）
- **加班报表导出** — 生成周内加班统计报表，按 20:00-22:00 / 22:00 之后两个时段统计
- **动态适配** — 自动提取数据中的年月，支持任意月份；人员数量动态扩展
- **部门过滤** — 支持按部门关键词筛选导出

## 技术栈

| 层 | 技术 |
|---|------|
| 前端 | Vue 3 + Element Plus + Vite |
| 后端 | FastAPI + Pandas + OpenPyXL |
| 部署 | Docker Compose + Nginx |

## 快速开始

### 环境要求

- Python 3.11+
- Node.js 20+ (推荐 22)

### 本地开发

```bash
# 启动后端
cd backend
pip install -r requirements.txt
python main.py
# API 运行在 http://localhost:8000
# API 文档: http://localhost:8000/docs

# 启动前端
cd frontend
npm install
npm run dev
# 前端运行在 http://localhost:5173
```

### Docker 部署

```bash
docker-compose up -d --build
# 访问 http://localhost:8082
```

## 使用流程

```
导出日报 (.xlsx) → 上传解析 → 数据预览 → 导出报表
```

1. 从企微/钉钉管理后台导出「上下班打卡日报」Excel 文件
2. 在系统「数据导入」页面上传文件
3. （可选）在「数据预览」页面查看和筛选数据
4. 在「报表导出」页面按需生成考勤统计或加班统计报表

## 数据源格式

上传的 `.xlsx` 文件需包含以下两个 Sheet：

| Sheet | 内容 |
|-------|------|
| 概况统计与打卡明细 | 每人每天一条记录：考勤结果、异常统计、假勤统计等 |
| 打卡详情 | 每次打卡一条记录：实际打卡时间、打卡状态、打卡地点等 |

`examples/` 目录下提供了示例数据文件供参考。

## 项目结构

```
├── backend/
│   ├── main.py              # FastAPI 入口
│   ├── config.json           # 考勤符号映射配置
│   ├── routers/              # API 路由
│   │   ├── upload.py         # 文件上传
│   │   ├── data.py           # 数据查询
│   │   └── export.py         # 报表导出
│   ├── services/             # 业务逻辑
│   │   ├── parser.py         # Excel 解析
│   │   ├── attendance.py     # 考勤报表生成
│   │   └── overtime.py       # 加班报表生成
│   └── tests/                # 单元测试
├── frontend/
│   ├── src/
│   │   ├── views/            # 页面组件
│   │   ├── api/              # API 调用
│   │   └── router/           # 路由配置
│   └── vite.config.js
├── 资料/                      # 报表模版文件
├── examples/                  # 示例数据
├── docker-compose.yml
└── deploy_report_project.sh   # 部署脚本
```

## 配置说明

### 考勤符号映射

编辑 `backend/config.json` 中的 `attendance_symbols` 字段可自定义考勤状态与符号的映射关系：

```json
{
  "attendance_symbols": {
    "正常": "√",
    "迟到": "※",
    "早退": "◇",
    "旷工": "×",
    "事假": "○",
    "病假": "☆"
  }
}
```

### 加班时段配置

```json
{
  "overtime_thresholds": {
    "period1_start": "20:00",
    "period1_end": "22:00",
    "period2_start": "22:00"
  }
}
```

## API 端点

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/upload` | 上传打卡日报文件 |
| GET | `/api/data/overview` | 分页查询考勤概况 |
| GET | `/api/data/details` | 分页查询打卡详情 |
| POST | `/api/export/attendance` | 导出考勤统计报表 |
| POST | `/api/export/overtime` | 导出加班统计报表 |

## License

MIT
