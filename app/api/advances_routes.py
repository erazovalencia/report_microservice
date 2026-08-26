from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..models.AdvancesModel import (
    AdvancesSapBookingExportRequest,
    AdvancesSupervisionExportRequest,
    AdvancesInvoiceExportRequest,
)
from ..services.ADVANCES.xlsx.sap_booking_export import AdvancesSapBookingExportService
from ..services.ADVANCES.xlsx.supervision_export import AdvancesSupervisionExportService
from ..services.ADVANCES.xlsx.invoice_export import AdvancesInvoiceExportService

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


@router.post("/supervision/export")
async def export_supervision(payload: AdvancesSupervisionExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay registros para exportar")

    try:
        service = AdvancesSupervisionExportService()
        buffer = service.generate_file(payload.rows, {"scope": payload.scope})
        fname = f"supervision_anticipos_{payload.scope or 'general'}.xlsx"
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando supervisión: {str(e)}")


@router.post("/invoices/export")
async def export_invoices(payload: AdvancesInvoiceExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay facturas para exportar")

    try:
        service = AdvancesInvoiceExportService()
        buffer = service.generate_file(payload.rows)
        return StreamingResponse(
            buffer,
            media_type=XLSX_MIME,
            headers={"Content-Disposition": 'attachment; filename="facturas_legalizacion.xlsx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando export de facturas: {str(e)}")
