from pydantic import BaseModel
from typing import List, Optional


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


class AdvancesSupervisionExportRow(BaseModel):
    documentNumber: str = "—"
    documentType: str  # ANTICIPO | CAJA_MENOR
    employeeName: str
    employeeDocumentId: str
    requestedAmount: float
    approvedAmount: Optional[float] = None
    status: str
    createdAt: str = ""
    openedAt: str = ""
    erpAdvanceNumber: str = ""
    legalizationApprovedAt: str = ""


class AdvancesSupervisionExportRequest(BaseModel):
    rows: List[AdvancesSupervisionExportRow]
    scope: str = ""  # "accounting" | "treasury" — solo para el título del archivo


class AdvancesInvoiceExportRow(BaseModel):
    invoiceNumber: str
    chargeNature: str  # ANTICIPO | CAJA_MENOR | TC | NINGUNA
    documentNumber: str = "—"
    thirdPartyTaxId: str
    thirdPartyName: str
    amount: float
    concept: str
    status1: str
    status2: str
    registeredByName: str
    createdAt: str = ""


class AdvancesInvoiceExportRequest(BaseModel):
    rows: List[AdvancesInvoiceExportRow]
