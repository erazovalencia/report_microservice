import io
import zipfile
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

from ..models.RdpModel import (
    RdpExportRequest,
    RdpImportTemplateRequest,
    RdpConsolidatedRequest,
    RdpConsolidatedBatchRequest,
)
from ..services.RDP.xlsx.report_export import RdpReportExportService
from ..services.RDP.xlsx.import_template import RdpImportTemplateService
from ..services.RDP.xlsx.consolidated_report import RdpConsolidatedService
from ..services.RDP.xlsx.consolidated_report_pdf import generate_consolidated_pdf
from ..services.RDP.parse.import_parser import parse_import_file

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Techo espejo del MAX_BULK_DOWNLOAD_EMPLOYEES de valera (bulk-download.ts) —
# validación redundante server-side, ninguno de los dos confía solo en el otro.
MAX_CONSOLIDATED_BATCH_SIZE = 60


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
            "shifts":      [e.dict() for e in payload.shifts],
            "absences":    [e.dict() for e in payload.absences],
            "bonuses":     [e.dict() for e in payload.bonuses],
            "workCenters": [e.dict() for e in payload.workCenters],
            "employees":   [e.dict() for e in payload.employees],
        }
        buffer = service.generate_file(data)
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": 'attachment; filename="plantilla_rdp.xlsx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando plantilla: {str(e)}")


@router.post("/report/consolidated")
async def generate_consolidated_report(payload: RdpConsolidatedRequest):
    try:
        service = RdpConsolidatedService()
        buffer  = service.generate_file(payload)
        emp     = payload.empleado
        fname   = f"consolidado_{emp.identificacion}_{payload.periodo.get('from','')}_{payload.periodo.get('to','')}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando consolidado: {str(e)}")


@router.post("/report/consolidated/pdf")
async def generate_consolidated_report_pdf(payload: RdpConsolidatedRequest):
    try:
        buffer = generate_consolidated_pdf(payload)
        emp    = payload.empleado
        fname  = f"consolidado_{emp.identificacion}_{payload.periodo.get('from','')}_{payload.periodo.get('to','')}.pdf"
        return StreamingResponse(
            buffer,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando PDF: {str(e)}")


@router.post("/report/consolidated/pdf/batch")
async def generate_consolidated_pdf_batch(payload: RdpConsolidatedBatchRequest):
    if not payload.items:
        raise HTTPException(status_code=400, detail="La lista de consolidados está vacía")
    if len(payload.items) > MAX_CONSOLIDATED_BATCH_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"Máximo {MAX_CONSOLIDATED_BATCH_SIZE} consolidados por solicitud",
        )
    try:
        first = payload.items[0]
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in payload.items:
                pdf_buffer = generate_consolidated_pdf(item)
                emp = item.empleado
                filename = f"consolidado_{emp.identificacion}_{item.periodo.get('from','')}_{item.periodo.get('to','')}.pdf"
                zf.writestr(filename, pdf_buffer.read())
        zip_buffer.seek(0)
        fname = f"consolidados_{first.periodo.get('from','')}_{first.periodo.get('to','')}.zip"
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando consolidados: {str(e)}")


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
