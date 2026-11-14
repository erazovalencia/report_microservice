from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse

from ..models.TreasuryModel import TreasuryBalancesExportRequest, TreasuryTransactionsExportRequest
from ..security import require_api_key
from ..services.TREASURY.xlsx.balances_export import TreasuryBalancesExportService
from ..services.TREASURY.xlsx.transactions_export import TreasuryTransactionsExportService

# Every treasury route carries payees, tax ids, amounts and account numbers, so
# the whole router requires the shared x-api-key.
router = APIRouter(dependencies=[Depends(require_api_key)])

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _xlsx(buffer, filename: str) -> StreamingResponse:
    return StreamingResponse(buffer, media_type=XLSX_MIME, headers={"Content-Disposition": f'attachment; filename="{filename}"'})


# Plain defs run in the thread pool: building a large workbook is CPU work
# that would otherwise block the event loop of this worker.
@router.post("/transactions/export")
def export_treasury_transactions(payload: TreasuryTransactionsExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay pagos para exportar")
    try:
        return _xlsx(TreasuryTransactionsExportService().generate_file(payload), "transacciones_tesoreria.xlsx")
    except Exception as error:
        raise HTTPException(status_code=500, detail=f"Error generando las transacciones de tesorería: {error}")


@router.post("/balances/export")
def export_treasury_balances(payload: TreasuryBalancesExportRequest):
    if not payload.rows:
        raise HTTPException(status_code=400, detail="No hay saldos para exportar")
    try:
        return _xlsx(TreasuryBalancesExportService().generate_file(payload), "saldos_tesoreria.xlsx")
    except Exception as error:
        raise HTTPException(status_code=500, detail=f"Error generando el histórico de saldos: {error}")
