from pydantic import BaseModel
from typing import List


class AdvancesSapBookingRow(BaseModel):
    invoiceId: int
    documentNumber: str  # anXXXX / cnXXXX del Anticipo/Caja Menor padre, o "—" si NINGUNA
    documentType: str    # ANTICIPO | CAJA_MENOR | NINGUNA
    invoiceNumber: str
    thirdPartyTaxId: str
    thirdPartyName: str
    amount: float
    concept: str
    project: str = ""
    invoiceDate: str = ""


class AdvancesSapBookingExportRequest(BaseModel):
    rows: List[AdvancesSapBookingRow]
