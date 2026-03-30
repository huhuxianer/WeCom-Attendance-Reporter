# backend/cache.py
import pandas as pd
from typing import Optional, Dict, Any

_overview_df: Optional[pd.DataFrame] = None
_details_df: Optional[pd.DataFrame] = None
_summary: Optional[Dict[str, Any]] = None
_source_filename: Optional[str] = None

# NOTE: not thread-safe; designed for single-worker uvicorn only
def set_data(overview: pd.DataFrame, details: pd.DataFrame, summary: dict, source_filename: str = None):
    global _overview_df, _details_df, _summary, _source_filename
    _overview_df = overview
    _details_df = details
    _summary = summary
    _source_filename = source_filename

def get_overview() -> Optional[pd.DataFrame]:
    return _overview_df

def get_details() -> Optional[pd.DataFrame]:
    return _details_df

def get_summary() -> Optional[dict]:
    return _summary

def get_source_filename() -> Optional[str]:
    return _source_filename

def has_data() -> bool:
    return _overview_df is not None

def clear_data():
    global _overview_df, _details_df, _summary, _source_filename
    _overview_df = None
    _details_df = None
    _summary = None
    _source_filename = None
