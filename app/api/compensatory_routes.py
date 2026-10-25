from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import JSONResponse

from ..services.BALANCES.parse.balance_import_parser import parse_balance_import_file

router = APIRouter()

COMPENSATORY_BALANCE_HEADERS = {"saldo", "cantidad", "compensatorios", "compensatorio", "dias", "días"}


@router.post("/balance-import/parse")
async def parse_compensatory_balance_import(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx o .xls")

    try:
        content = await file.read()
        rows = parse_balance_import_file(content, COMPENSATORY_BALANCE_HEADERS)
        return JSONResponse(content={"rows": rows, "total": len(rows)})
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error al procesar el archivo: {str(e)}")
