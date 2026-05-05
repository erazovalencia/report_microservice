import io
import zipfile

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response, StreamingResponse

from ..models.HvModel import HvBatchRequest, HvGenerateRequest
from ..security import require_api_key
from ..services.HV.pdf.hv_single import HvPdfService

router = APIRouter()

MAX_BATCH_SIZE = 50


@router.post(
    "/generate",
    summary="Genera la Hoja de Vida de un empleado en PDF",
    dependencies=[Depends(require_api_key)],
)
def generate_hv(request: HvGenerateRequest):
    try:
        service = HvPdfService()
        pdf_bytes = service.generate_pdf(request.employee.dict())
        filename = f"hv_{request.employee.identification}_{request.employee.lastName}.pdf"
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando HV: {str(e)}")


@router.post(
    "/batch",
    summary="Genera Hojas de Vida de múltiples empleados en un ZIP",
    dependencies=[Depends(require_api_key)],
)
def generate_hv_batch(request: HvBatchRequest):
    if not request.employees:
        raise HTTPException(status_code=400, detail="La lista de empleados está vacía")
    if len(request.employees) > MAX_BATCH_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"Máximo {MAX_BATCH_SIZE} empleados por solicitud",
        )
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for emp in request.employees:
                service = HvPdfService()
                pdf_bytes = service.generate_pdf(emp.dict())
                filename = f"{emp.identification}_{emp.lastName}.pdf"
                zf.writestr(filename, pdf_bytes)
        zip_buffer.seek(0)
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": 'attachment; filename="hojas_de_vida.zip"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando HVs: {str(e)}")
