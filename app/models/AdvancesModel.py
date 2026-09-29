from pydantic import BaseModel
from typing import List, Optional


class AdvancesSapBookingRow(BaseModel):
    """
    Una fila = una factura = un documento contable de una sola línea
    (bloque 1, columnas P-V). Nombres de campo = código técnico real de SAP
    — se mantienen así a propósito,
    no en inglés "de negocio", porque es exactamente lo que Contabilidad/SAP
    reconoce al validar el archivo contra la especificación real.

    Columnas L/M/N/O (retención) y los bloques 2-4 (W-AQ, multi-línea
    contable) quedan fuera de este modelo — siempre en blanco en v1, no hay
    nada que el caller pueda mandar ahí.
    """
    # Cabecera del documento (A-K)
    buscs: str = "R"            # A — constante
    accnt: str                  # B — código SAP del acreedor, match por NIT
    bldat: str                  # C — fecha de la factura, DD.MM.YYYY
    xblnr: str = ""             # D — en blanco
    budat: str                  # E — fecha de contabilización, constante POR LOTE
    blart: str = "DS"           # F — constante
    wrbtr: float                # G — valor total de la factura
    waers: str = "COP"          # H — constante
    xmwst: str = "X"            # I — constante
    mwskz: str = "RZ"           # J — constante
    sgtxt: str                  # K — texto integrado: "{documentNumber} {inicial+apellido legalizador}", máx 25

    # Línea contable 1 — único bloque en uso (P-V)
    hkont_01: str                # P — cuenta contable del Servicio elegido
    wrbtr_01: float               # Q — mismo valor que wrbtr (factura de una sola línea)
    mwskz_01: str = "RZ"          # R — constante, igual a J
    sgtxt_01: str                 # S — descripción/concepto de la factura, máx 50
    kostl_01: str                 # T — centro de costo del legalizador
    aufnr_01: str = ""            # U — número de orden (excluyente con V, uno de los dos obligatorio)
    projk_01: str = ""            # V — elemento PEP (excluyente con U)


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
