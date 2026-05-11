from pydantic import BaseModel, Field
from typing import List, Optional


class RdpExportRow(BaseModel):
    fecha: str
    unidadOrganizativa: str
    identificacion: str
    nombre: str
    tipo: str
    turnoAusencia: str
    tipoBono: str = ""
    centroCosto: str = ""
    actividad: str = ""
    horarioEntrada: str = ""
    horarioSalida: str = ""
    totalHoras: str = ""
    heDiurna: float = 0.0
    heDiurnaFestiva: float = 0.0
    heNocturna: float = 0.0
    heNocturnaFestiva: float = 0.0
    heTotal: float = 0.0
    hrd: float = 0.0
    rnf: float = 0.0
    rotacion: str = ""
    impactoVacaciones: float = 0.0
    impactoCompensatorios: float = 0.0
    estadoReporte: str
    biostarEntrada: str = ""
    biostarSalida: str = ""


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
    shifts:      List[RdpCatalogEntry] = []
    absences:    List[RdpCatalogEntry] = []
    bonuses:     List[RdpCatalogEntry] = []
    workCenters: List[RdpCatalogEntry] = []
    employees:   List[RdpEmployeeEntry] = []


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
