from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse

from ..models.VacationsModel import VacationExportRequest
from ..services.VACATIONS.xlsx.report_export import VacationReportExportService
from ..services.VACATIONS.parse.balance_import_parser import parse_balance_import_file

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/report/export")
async def export_vacations_report(payload: VacationExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay filas para exportar")

    try:
        service = VacationReportExportService()
        buffer = service.generate_file(payload.rows)
        suffix = f"_{payload.period_from}_{payload.period_to}" if payload.period_from else ""
        fname  = f"vacaciones{suffix}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando reporte: {str(e)}")


@router.post("/balance-import/parse")
async def parse_vacation_balance_import(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx o .xls")

    try:
        content = await file.read()
        rows = parse_balance_import_file(content)
        return JSONResponse(content={"rows": rows, "total": len(rows)})
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error al procesar el archivo: {str(e)}")
