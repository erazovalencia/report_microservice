from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..models.AdvancesModel import AdvancesSapBookingExportRequest
from ..services.ADVANCES.xlsx.sap_booking_export import AdvancesSapBookingExportService

router = APIRouter()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/sap-booking/export")
async def export_sap_booking(payload: AdvancesSapBookingExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay facturas para exportar")

    try:
        service = AdvancesSapBookingExportService()
        buffer = service.generate_file(payload.rows)
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": 'attachment; filename="cargue_sap_anticipos.xlsx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando cargue SAP: {str(e)}")
