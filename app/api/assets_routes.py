from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse

from ..services.ASSETS.parse.asset_import_parser import parse_asset_import_file

router = APIRouter()


@router.post("/import/parse")
async def parse_asset_import(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls", ".xlsm")):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx, .xls o .xlsm")

    try:
        content = await file.read()
        rows = parse_asset_import_file(content)
        return JSONResponse(content={"rows": rows, "total": len(rows)})
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error al procesar el archivo: {str(e)}")
