from pydantic import BaseModel, Field
from typing import List


class AttendanceHistoryExportRow(BaseModel):
    cedula: str
    nombre: str
    fecha: str
    hora: str
    tipo: str
    dispositivo: str = ""


class AttendanceHistoryExportRequest(BaseModel):
    rows: List[AttendanceHistoryExportRow]
    period_from: str = Field(default="", alias="from")
    period_to: str = Field(default="", alias="to")

    class Config:
        allow_population_by_field_name = True
