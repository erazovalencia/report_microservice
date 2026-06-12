from pydantic import BaseModel, Field
from typing import List


class VacationExportRow(BaseModel):
    fechaSolicitud: str = ""          # ISO datetime
    identificacion: str = ""
    nombre: str = ""
    tipo: str = ""                    # TIME | MONEY | BOTH
    inicio: str = ""                  # YYYY-MM-DD ("" para MONEY)
    fin: str = ""                     # YYYY-MM-DD ("" para MONEY)
    estado: str = ""                  # pending | approved | in_progress | taken | rejected | cancelled
    saldo: str = ""                   # accruedAtRequest
    diasTiempo: float = 0.0
    diasDinero: float = 0.0
    disponible: str = ""
    aprobador: str = ""
    fechaDecision: str = ""           # ISO datetime (approvedAt | rejectedAt)
    diasTomados: float = 0.0
    motivo: str = ""
    motivoRechazo: str = ""
    observaciones: str = ""


class VacationExportRequest(BaseModel):
    rows: List[VacationExportRow]
    period_from: str = Field(default="", alias="from")
    period_to: str = Field(default="", alias="to")

    class Config:
        allow_population_by_field_name = True
