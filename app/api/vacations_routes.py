from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..models.VacationsModel import VacationExportRequest
from ..services.VACATIONS.xlsx.report_export import VacationReportExportService

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
