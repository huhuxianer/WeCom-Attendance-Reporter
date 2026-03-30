"""export router - 生成并下载报表文件。计划端点: POST /api/export/attendance, POST /api/export/overtime"""
import tempfile, os, re
from datetime import datetime
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from pathlib import Path
import pandas as pd
import cache
from services.attendance import generate_attendance_report
from services.overtime import generate_overtime_report

router = APIRouter()

TEMPLATE_DIR = Path(__file__).parent.parent.parent / "资料"


class ExportRequest(BaseModel):
    dept: Optional[str] = None


def _get_year_month(df: pd.DataFrame) -> tuple[int, int]:
    try:
        if "日期" in df.columns:
            valid_dates = df[~df["日期"].isin(["", "--", "nan"]) & df["日期"].notna()]["日期"]
            if not valid_dates.empty:
                first_date = str(valid_dates.iloc[0]).split()[0]
                for fmt in ["%Y/%m/%d", "%Y-%m-%d"]:
                    try:
                        dt = datetime.strptime(first_date, fmt)
                        return dt.year, dt.month
                    except Exception:
                        pass
    except Exception:
        pass
    now = datetime.now()
    return now.year, now.month


def _extract_file_tag(filename: Optional[str]) -> str:
    """从文件名括号中提取标签，例如 'xxx (部门A).xlsx' -> '部门A'"""
    if not filename:
        return "全部"
    # 匹配括号内容，支持半角和全角括号
    match = re.search(r'[\(\（](.*?)[\)\）]', filename)
    if match:
        return match.group(1).strip()
    return "全部"


@router.post("/export/attendance")
def export_attendance(req: ExportRequest):
    df = cache.get_overview()
    if df is None:
        raise HTTPException(400, "请先上传数据文件")

    year, month = _get_year_month(df)
    
    # 确定部门标签
    if req.dept:
        dept_label = req.dept
    else:
        source_name = cache.get_source_filename()
        dept_label = _extract_file_tag(source_name)

    filename = f"{dept_label}_{year}年{month}月考勤数据统计.xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()

    template = str(TEMPLATE_DIR / "1月考勤数据统计_模版.xlsx")
    details_df = cache.get_details()
    generate_attendance_report(df, template, tmp.name, year, month, dept=req.dept, details_df=details_df)

    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


@router.post("/export/overtime")
def export_overtime(req: ExportRequest):
    df = cache.get_details()
    overview_df = cache.get_overview()
    if df is None or overview_df is None:
        raise HTTPException(400, "请先上传数据文件")

    year, month = _get_year_month(df)

    # 确定部门标签
    if req.dept:
        dept_label = req.dept
    else:
        source_name = cache.get_source_filename()
        dept_label = _extract_file_tag(source_name)

    filename = f"{dept_label}_{year}年{month}月周内加班统计.xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()

    template = str(TEMPLATE_DIR / "1月周内加班统计_模版.xlsx")
    generate_overtime_report(df, template, tmp.name, year, month, dept=req.dept, overview_df=overview_df)

    return FileResponse(
        tmp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )
