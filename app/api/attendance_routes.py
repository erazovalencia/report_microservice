from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..models.AttendanceModel import AttendanceHistoryExportRequest
from ..services.ATTENDANCE.xlsx.history_export import AttendanceHistoryExportService

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/history/export")
async def export_attendance_history(payload: AttendanceHistoryExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay filas para exportar")

    try:
        service = AttendanceHistoryExportService()
        buffer = service.generate_file(payload.rows)
        fname = f"historico_asistencia_{payload.period_from}_{payload.period_to}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando histórico: {str(e)}")
