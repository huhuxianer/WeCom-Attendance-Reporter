"""upload router - 处理 Excel 文件上传与解析。计划端点: POST /api/upload"""
import tempfile, os
from fastapi import APIRouter, UploadFile, File, HTTPException
from services.parser import parse_xlsx
import cache

router = APIRouter()

@router.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(400, "仅支持 .xlsx 文件")

    # 写临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        result = parse_xlsx(tmp_path)
        cache.set_data(result.overview, result.details, result.summary, source_filename=file.filename)
        return result.summary
    except Exception as e:
        raise HTTPException(500, f"解析失败: {str(e)}")
    finally:
        os.unlink(tmp_path)
