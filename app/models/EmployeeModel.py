from pydantic import BaseModel
from typing import List, Dict


class EmployeeExportColumn(BaseModel):
    key: str
    label: str
    width: int = 18


class EmployeeExportRequest(BaseModel):
    columns: List[EmployeeExportColumn]
    rows: List[Dict[str, str]]
    title: str = "LISTADO DE EMPLEADOS"
