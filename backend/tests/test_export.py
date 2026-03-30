# backend/tests/test_export.py
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from fastapi.testclient import TestClient
from main import app
import cache

client = TestClient(app)

SAMPLE_FILE = os.path.join(os.path.dirname(__file__), "../../资料/上下班打卡_日报_20260101-20260131 .xlsx")


def _upload():
    with open(SAMPLE_FILE, "rb") as f:
        client.post(
            "/api/upload",
            files={"file": ("test.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        )


def test_export_attendance_no_data_returns_400():
    cache.clear_data()
    resp = client.post("/api/export/attendance", json={})
    assert resp.status_code == 400


def test_export_attendance_returns_xlsx():
    _upload()
    resp = client.post("/api/export/attendance", json={"dept": "示例部门"})
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers.get("content-type", "")


def test_export_overtime_returns_xlsx():
    _upload()
    resp = client.post("/api/export/overtime", json={"dept": "示例部门"})
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers.get("content-type", "")
