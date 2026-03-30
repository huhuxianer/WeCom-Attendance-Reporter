# backend/services/config_loader.py
import json
from pathlib import Path

_config = None

def get_config() -> dict:
    global _config
    if _config is None:
        config_path = Path(__file__).parent.parent / "config.json"
        with open(config_path, encoding="utf-8") as f:
            _config = json.load(f)
    return _config

def get_attendance_symbol(status: str) -> str:
    config = get_config()
    symbols = config["attendance_symbols"]
    status_str = str(status)
    # 优先精确匹配
    if status_str in symbols:
        return symbols[status_str]
    # 再进行包含匹配
    for key, symbol in symbols.items():
        if key in status_str:
            return symbol
    return status_str if status and status_str != "nan" else ""
