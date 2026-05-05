from pydantic import BaseModel
from typing import Optional, List


class HvEmployeePayload(BaseModel):
    # Identificación
    documentType: Optional[str] = None
    identification: str
    idExpeditionPlace: Optional[str] = None
    idExpeditionDate: Optional[str] = None

    # Datos personales
    name: str
    lastName: str
    genre: Optional[str] = None
    birthDate: Optional[str] = None
    placeOfBirth: Optional[str] = None
    nationality: Optional[str] = None
    maritalStatus: Optional[str] = None
    rh: Optional[str] = None

    # Contacto
    phone: Optional[str] = None
    countryCode: Optional[str] = None
    email: str
    address: Optional[str] = None
    department: Optional[str] = None

    # Formación
    educationLevel: Optional[str] = None
    profession: Optional[str] = None

    # Cargo corporativo
    position: Optional[str] = None
    serviceLineName: Optional[str] = None
    reportGroup: Optional[str] = None


class HvGenerateRequest(BaseModel):
    employee: HvEmployeePayload


class HvBatchRequest(BaseModel):
    employees: List[HvEmployeePayload]
