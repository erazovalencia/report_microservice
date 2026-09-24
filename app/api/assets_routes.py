from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse, StreamingResponse

from ..services.ASSETS.parse.asset_import_parser import parse_asset_import_file
from ..services.ASSETS.parse.asset_bulk_edit_parser import parse_asset_bulk_edit_file
from ..services.ASSETS.xlsx.asset_export import AssetExportService
from ..models.AssetModel import AssetExportRequest

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


@router.post("/bulk-edit/parse")
async def parse_asset_bulk_edit(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx o .xlsm")

    try:
        content = await file.read()
        rows = parse_asset_bulk_edit_file(content)
        return JSONResponse(content={"rows": rows, "total": len(rows)})
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error al procesar el archivo: {str(e)}")


@router.post("/export")
async def export_assets(payload: AssetExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay activos para exportar")

    try:
        service = AssetExportService()
        buffer = service.generate_file(payload.rows)
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": 'attachment; filename="inventario_activos.xlsx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando el inventario: {str(e)}")
