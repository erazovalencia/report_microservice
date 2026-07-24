from datetime import datetime

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..models.EmployeeModel import EmployeeExportRequest
from ..services.EMPLOYEES.xlsx.employee_export import EmployeeExportService

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/export")
async def export_employees(payload: EmployeeExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay empleados para exportar")

    try:
        service = EmployeeExportService()
        buffer = service.generate_file(payload)
        fname = f"empleados_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando listado de empleados: {str(e)}")
