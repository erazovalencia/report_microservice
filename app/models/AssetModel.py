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
