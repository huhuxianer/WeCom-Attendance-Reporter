# backend/services/parser.py
import pandas as pd
from dataclasses import dataclass


@dataclass
class ParseResult:
    overview: pd.DataFrame
    details: pd.DataFrame
    summary: dict


def parse_xlsx(filepath: str) -> ParseResult:
    # === 解析 Sheet1：概况统计 ===
    # 读取全部内容（不指定 header，按原始行处理）
    raw_overview = pd.read_excel(filepath, sheet_name=0, header=None, dtype=str)

    # Sheet1 实际行结构：
    #   index=0: sheet 名称行
    #   index=1: 统计时间说明行
    #   index=2: 分组名行（时间、姓名、账号、基础信息、考勤概况...）
    #   index=3: 字段名行（部门、职务、工号、班次...）
    #   index=4+: 数据行
    #
    # 合并分组名 + 字段名：优先取字段名行的值，字段名为空时取分组名行的值
    # 这样 col0=时间, col1=姓名, col2=账号, col3=部门, col4=职务 ...
    # col50=打卡时间(上班1分组), col52=打卡时间(下班1分组) 等重复列名保留分组信息
    grp_row = raw_overview.iloc[2].tolist()
    fld_row = raw_overview.iloc[3].tolist()
    col_headers = _merge_headers(grp_row, fld_row)

    # 数据从 index=4 开始
    overview_data = raw_overview.iloc[4:].copy()
    overview_data.columns = col_headers
    overview_data = overview_data.reset_index(drop=True)
    overview = _extract_overview_cols(overview_data, raw_overview)

    # 过滤掉空行（姓名为空）
    overview = overview[overview["姓名"].notna() & (overview["姓名"] != "") & (overview["姓名"] != "nan")]
    overview = overview.reset_index(drop=True)

    # === 解析 Sheet2：打卡详情 ===
    raw_details = pd.read_excel(filepath, sheet_name=1, header=None, dtype=str)

    # Sheet2 实际行结构：
    #   index=0: sheet 名称行
    #   index=1: 统计时间说明行
    #   index=2: 字段名行（日期、姓名、账号、部门...）
    #   index=3: 空行
    #   index=4+: 数据行
    detail_headers = raw_details.iloc[2].fillna("").tolist()

    # 数据从 index=4 开始（index=3 是空行）
    details_data = raw_details.iloc[4:].copy()
    details_data.columns = detail_headers
    details_data = details_data.reset_index(drop=True)
    details = _extract_details_cols(details_data)
    details = details[details["姓名"].notna() & (details["姓名"] != "") & (details["姓名"] != "nan")]
    details = details.reset_index(drop=True)

    # === 生成摘要 ===
    persons = overview["姓名"].unique().tolist()
    departments = sorted([d for d in overview["部门"].dropna().unique().tolist() if d and d != "nan"])
    dates = sorted([d for d in overview["日期"].dropna().unique().tolist() if d and d != "nan"])
    summary = {
        "total_persons": len(persons),
        "date_range": f"{dates[0]} ~ {dates[-1]}" if dates else "",
        "departments": departments,
        "total_rows_overview": len(overview),
        "total_rows_details": len(details),
    }

    return ParseResult(overview=overview, details=details, summary=summary)


def _merge_headers(grp_row: list, fld_row: list) -> list:
    """合并分组名行和字段名行，生成完整列名列表。

    优先取字段名行的值；字段名为空时，取分组名行的值。
    用于处理 Sheet1 的两级列头。
    """
    result = []
    for g, f in zip(grp_row, fld_row):
        g_val = str(g).strip() if pd.notna(g) else ""
        f_val = str(f).strip() if pd.notna(f) else ""
        if g_val in ("nan", "None"):
            g_val = ""
        if f_val in ("nan", "None"):
            f_val = ""
        if f_val:
            result.append(f_val)
        elif g_val:
            result.append(g_val)
        else:
            result.append("")
    return result


def _find_col(col_list, name):
    """在列名列表中找包含 name 的第一个列名"""
    for c in col_list:
        if name in str(c):
            return c
    return None


def _extract_overview_cols(df: pd.DataFrame, raw: pd.DataFrame) -> pd.DataFrame:
    """从原始 DataFrame 中提取需要的列，统一命名"""
    cols = df.columns.tolist()

    def get(name):
        c = _find_col(cols, name)
        return df[c].fillna("") if c else pd.Series([""] * len(df))

    result = pd.DataFrame()
    result["日期"] = get("时间")
    result["姓名"] = get("姓名")
    result["账号"] = get("账号")
    result["部门"] = get("部门")
    result["职务"] = get("职务")
    result["工号"] = get("工号")
    result["班次"] = get("班次")
    result["考勤结果"] = get("考勤结果")
    result["异常合计"] = get("异常合计(次)")
    result["迟到次数"] = get("迟到次数(次)")
    result["迟到时长"] = get("迟到时长(分钟)")
    result["早退次数"] = get("早退次数(次)")
    result["旷工次数"] = get("旷工次数(次)")
    result["缺卡次数"] = get("缺卡次数(次)")
    result["加班状态"] = get("加班状态")
    result["加班时长"] = get("加班时长(小时)")
    result["工作日加班费时长"] = get("工作日加班计为加班费(小时)")
    result["由于外勤次数"] = get("外勤次数(次)")
    result["外出小时"] = get("外出(小时)")
    result["出差天数"] = get("出差(天)")
    result["事假天数"] = get("事假(天)")
    result["病假天数"] = get("病假(天)")
    result["调休假天数"] = get("调休假(天)")
    result["年假天数"] = get("年假(天)")
    result["婚假天数"] = get("婚假(天)")
    result["产假天数"] = get("产假(天)")
    result["陪产假天数"] = get("陪产假(天)")
    result["丧假天数"] = get("丧假(天)")
    result["其他天数"] = get("其他(天)")

    # 上下班打卡时间：
    # Sheet1 分组名行中 col50=上班1, col52=下班1
    # 合并后列名 col50='打卡时间'（上班1分组）, col52='打卡时间'（下班1分组）
    # 直接按原始列位置 iloc[:, 50] / iloc[:, 52] 取值
    # 重置 index 保持与 df 对齐
    try:
        result["上班打卡时间"] = raw.iloc[4:, 50].fillna("").values[:len(df)]
    except Exception:
        result["上班打卡时间"] = pd.Series([""] * len(df))

    try:
        result["下班打卡时间"] = raw.iloc[4:, 52].fillna("").values[:len(df)]
    except Exception:
        result["下班打卡时间"] = pd.Series([""] * len(df))

    return result.reset_index(drop=True)


def _extract_details_cols(df: pd.DataFrame) -> pd.DataFrame:
    """从打卡详情原始 DataFrame 中提取需要的列"""
    cols = df.columns.tolist()

    def get(name):
        c = _find_col(cols, name)
        return df[c].fillna("") if c else pd.Series([""] * len(df))

    result = pd.DataFrame()
    result["日期"] = get("日期")
    result["姓名"] = get("姓名")
    result["账号"] = get("账号")
    result["部门"] = get("部门")
    result["职务"] = get("职务")
    result["所属规则"] = get("所属规则")
    result["打卡类型"] = get("打卡类型")
    result["应打卡时间"] = get("应打卡时间")
    result["实际打卡时间"] = get("实际打卡时间")
    result["打卡状态"] = get("打卡状态")
    result["打卡地点"] = get("打卡地点")
    result["假勤申请"] = get("假勤申请")

    return result.reset_index(drop=True)
