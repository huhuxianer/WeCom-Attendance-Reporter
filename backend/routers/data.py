"""data router - 提供解析后数据的分页查询。计划端点: GET /api/data/overview, GET /api/data/details"""
from fastapi import APIRouter, HTTPException, Query
from typing import Optional
import cache

router = APIRouter()


def _df_to_records(df, page: int, page_size: int):
    total = len(df)
    start = (page - 1) * page_size
    end = start + page_size
    items = df.iloc[start:end].fillna("").to_dict(orient="records")
    return {"total": total, "page": page, "page_size": page_size, "items": items}


@router.get("/data/overview")
def get_overview(
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=200),
    keyword: Optional[str] = None,
    dept: Optional[str] = None,
):
    df = cache.get_overview()
    if df is None:
        raise HTTPException(400, "请先上传数据文件")
    if keyword:
        df = df[df["姓名"].str.contains(keyword, na=False)]
    if dept:
        df = df[df["部门"].str.contains(dept, na=False)]
    return _df_to_records(df, page, page_size)


@router.get("/data/details")
def get_details(
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=200),
    name: Optional[str] = None,
    date: Optional[str] = None,
    dept: Optional[str] = None,
):
    df = cache.get_details()
    if df is None:
        raise HTTPException(400, "请先上传数据文件")
    if name:
        df = df[df["姓名"].str.contains(name, na=False)]
    if date:
        df = df[df["日期"].str.contains(date, na=False)]
    if dept:
        df = df[df["部门"].str.contains(dept, na=False)]
    return _df_to_records(df, page, page_size)
