# backend/tests/test_upload.py
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

SAMPLE_FILE = os.path.join(os.path.dirname(__file__), "../../资料/上下班打卡_日报_20260101-20260131 .xlsx")

def test_upload_no_file_returns_422():
    resp = client.post("/api/upload")
    assert resp.status_code == 422

def test_upload_valid_file_returns_summary():
    with open(SAMPLE_FILE, "rb") as f:
        resp = client.post(
            "/api/upload",
            files={"file": ("test.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        )
    assert resp.status_code == 200
    data = resp.json()
    assert "total_persons" in data
    assert "departments" in data
    assert data["total_persons"] > 0
