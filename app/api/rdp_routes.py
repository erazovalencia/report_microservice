import io
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

from ..models.RdpModel import RdpExportRequest, RdpImportTemplateRequest
from ..services.RDP.xlsx.report_export import RdpReportExportService
from ..services.RDP.xlsx.import_template import RdpImportTemplateService
from ..services.RDP.parse.import_parser import parse_import_file

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/report/export")
async def export_rdp_report(payload: RdpExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay filas para exportar")

    try:
        service = RdpReportExportService()
        buffer = service.generate_file(payload.rows)
        fname  = f"reporte_rdp_{payload.period_from}_{payload.period_to}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando reporte: {str(e)}")


@router.post("/import/template")
async def generate_import_template(payload: RdpImportTemplateRequest):
    try:
        service = RdpImportTemplateService()
        data = {
            "shifts":    [e.dict() for e in payload.shifts],
            "absences":  [e.dict() for e in payload.absences],
            "bonuses":   [e.dict() for e in payload.bonuses],
            "projects":  [e.dict() for e in payload.projects],
            "employees": [e.dict() for e in payload.employees],
        }
        buffer = service.generate_file(data)
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": 'attachment; filename="plantilla_rdp.xlsx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando plantilla: {str(e)}")


@router.post("/import/parse")
async def parse_rdp_import(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx o .xls")

    try:
        content = await file.read()
        rows = parse_import_file(content)
        return JSONResponse(content={"rows": rows, "total": len(rows)})
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error al procesar el archivo: {str(e)}")
