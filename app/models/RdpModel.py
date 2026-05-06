from pydantic import BaseModel, Field
from typing import List, Optional


class RdpExportRow(BaseModel):
    fecha: str
    unidadOrganizativa: str
    identificacion: str
    nombre: str
    tipo: str
    turnoAusencia: str
    tipoBono: str
    proyecto: str
    heDiurna: float
    heDiurnaFestiva: float
    heNocturna: float
    heNocturnaFestiva: float
    heTotal: float
    rotacion: str
    impactoVacaciones: float
    impactoCompensatorios: float
    estadoReporte: str
    biostarEntrada: str
    biostarSalida: str


class RdpExportRequest(BaseModel):
    rows: List[RdpExportRow]
    period_from: str = Field(default="", alias="from")
    period_to: str = Field(default="", alias="to")

    class Config:
        allow_population_by_field_name = True


class RdpCatalogEntry(BaseModel):
    code: str
    name: str


class RdpEmployeeEntry(BaseModel):
    cedula: str
    nombre: str


class RdpImportTemplateRequest(BaseModel):
    shifts: List[RdpCatalogEntry] = []
    absences: List[RdpCatalogEntry] = []
    bonuses: List[RdpCatalogEntry] = []
    projects: List[RdpCatalogEntry] = []
    employees: List[RdpEmployeeEntry] = []


class RdpParsedImportRow(BaseModel):
    rowIndex: int
    identificacion: str
    esAusencia: bool
    turno: Optional[str] = None
    tipoAusencia: Optional[str] = None
    tipoBono: Optional[str] = None
    proyecto: Optional[str] = None
    horaExtra: Optional[float] = None
    notas: Optional[str] = None
    parseErrors: List[str] = []
