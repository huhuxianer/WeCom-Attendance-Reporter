# backend/tests/test_parser.py
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import pandas as pd
from services.parser import parse_xlsx, ParseResult

SAMPLE_FILE = os.path.join(os.path.dirname(__file__), "../../资料/上下班打卡_日报_20260101-20260131 .xlsx")

def test_parse_returns_result():
    result = parse_xlsx(SAMPLE_FILE)
    assert isinstance(result, ParseResult)

def test_overview_has_rows():
    result = parse_xlsx(SAMPLE_FILE)
    assert len(result.overview) > 100

def test_details_has_rows():
    result = parse_xlsx(SAMPLE_FILE)
    assert len(result.details) > 100

def test_summary_fields():
    result = parse_xlsx(SAMPLE_FILE)
    assert "total_persons" in result.summary
    assert "date_range" in result.summary
    assert "departments" in result.summary

def test_overview_key_columns():
    result = parse_xlsx(SAMPLE_FILE)
    cols = result.overview.columns.tolist()
    assert "姓名" in cols
    assert "部门" in cols
    assert "日期" in cols

def test_details_key_columns():
    result = parse_xlsx(SAMPLE_FILE)
    cols = result.details.columns.tolist()
    assert "姓名" in cols
    assert "打卡类型" in cols
    assert "实际打卡时间" in cols

def test_summary_persons_count():
    result = parse_xlsx(SAMPLE_FILE)
    assert result.summary["total_persons"] > 5

def test_summary_departments_not_empty():
    result = parse_xlsx(SAMPLE_FILE)
    assert len(result.summary["departments"]) > 0
