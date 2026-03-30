# backend/tests/test_data.py
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


def test_overview_no_data_returns_400():
    cache.clear_data()
    resp = client.get("/api/data/overview")
    assert resp.status_code == 400


def test_overview_pagination():
    _upload()
    resp = client.get("/api/data/overview?page=1&page_size=10")
    assert resp.status_code == 200
    data = resp.json()
    assert "items" in data
    assert "total" in data
    assert len(data["items"]) <= 10


def test_overview_filter_by_name():
    _upload()
    resp = client.get("/api/data/overview?keyword=员工A")
    assert resp.status_code == 200
    items = resp.json()["items"]
    for item in items:
        assert "员工A" in item.get("姓名", "")


def test_details_pagination():
    _upload()
    resp = client.get("/api/data/details?page=1&page_size=20")
    assert resp.status_code == 200
    data = resp.json()
    assert len(data["items"]) <= 20


def test_details_filter_by_name():
    _upload()
    resp = client.get("/api/data/details?name=员工B")
    assert resp.status_code == 200
    items = resp.json()["items"]
    for item in items:
        assert "员工B" in item.get("姓名", "")
