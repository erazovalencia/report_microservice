from pydantic import BaseModel
from typing import List, Optional


class AssetExportRow(BaseModel):
    assetTag: str
    sapEquipmentCode: Optional[str] = None
    parentEquipment: Optional[str] = None
    assetLine: str
    category: Optional[str] = None
    description: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    serialNumber: Optional[str] = None
    sapInventoryNumber: Optional[str] = None
    sapLocation: Optional[str] = None
    sapTechnicalLocation: Optional[str] = None
    emplacementCode: Optional[str] = None
    emplacementCenter: Optional[str] = None
    sapCompany: Optional[str] = None
    sapCostCenter: Optional[str] = None
    sapPepElement: Optional[str] = None
    status: str
    sapSystemStatus: Optional[str] = None
    assignedTo: Optional[str] = None
    createdAt: Optional[str] = None


class AssetExportRequest(BaseModel):
    rows: List[AssetExportRow]


class AssetHistoryExportAsset(BaseModel):
    """El activo como está hoy (cabecera del Excel del historial)."""
    assetTag: str
    sapEquipmentCode: Optional[str] = None
    description: Optional[str] = None
    assetLine: str
    status: str
    brand: Optional[str] = None
    model: Optional[str] = None
    serialNumber: Optional[str] = None
    currentAssignment: Optional[str] = None


class AssetHistoryExportRow(BaseModel):
    """Un movimiento con los textos ya resueltos en español por valera."""
    type: str
    date: str
    movement: str
    detail: str
    statusBefore: Optional[str] = None
    statusAfter: Optional[str] = None
    assignee: Optional[str] = None
    note: Optional[str] = None
    actor: Optional[str] = None
    source: str


class AssetHistoryExportRequest(BaseModel):
    asset: AssetHistoryExportAsset
    rows: List[AssetHistoryExportRow]
    generatedAt: str
    generatedBy: Optional[str] = None


class AssetAssignmentNoticeAsset(BaseModel):
    assetTag: str
    sapEquipmentCode: Optional[str] = None
    description: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    serialNumber: Optional[str] = None


class AssetAssignmentNoticeRequest(BaseModel):
    """Formato de asignación de activos a un empleado (acta de entrega)."""
    lineLabel: str
    employeeName: str
    employeeDocumentId: str
    position: Optional[str] = None
    department: Optional[str] = None
    company: Optional[str] = None
    authorizedBy: Optional[str] = None
    deliveredBy: Optional[str] = None
    deliveredByDocumentId: Optional[str] = None
    assignedAt: str
    condition: Optional[str] = None
    assets: List[AssetAssignmentNoticeAsset]
