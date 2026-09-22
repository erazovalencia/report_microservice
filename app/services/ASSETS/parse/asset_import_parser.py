import io
import re
import unicodedata
from typing import List, Dict, Any, Optional
import openpyxl

# Import masivo de activos desde un export tipo IH08 de SAP (ver
# valera/ASSETS_INVENTORY_DESIGN.md §2/§6). Detección por encabezado, no por
# nombre/posición de hoja (mismo criterio que balance_import_parser.py) — el
# export real de IH08 trae varias columnas con acentos/símbolos que openpyxl
# entrega como texto Unicode normal; normalizamos para matchear sin
# depender de la codificación exacta del archivo.

def _norm(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _norm_header(val) -> str:
    """Minúsculas, sin tildes ni símbolos, sin espacios: solo letras/dígitos.

    Los indicadores ordinales (º, ª) se eliminan ANTES de NFKD: NFKD descompone
    "º" en "o", y "Nº Pieza fabric." nunca coincidiría con "n pieza fabric".
    Colapsar los espacios también hace inmune la detección a variantes de
    puntuación ("Fabr. Nº-serie" / "Fabr Nº serie").
    """
    s = _norm(val).lower().replace("º", "").replace("ª", "").replace("°", "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s)


def _headers(*labels: str) -> set:
    return {_norm_header(label) for label in labels}


EQUIPMENT_HEADERS = _headers("Equipo")
PARENT_EQUIPMENT_HEADERS = _headers("Equipo superior")
FIXED_ASSET_HEADERS = _headers("Activo fijo")
MANUFACTURER_HEADERS = _headers("Fabricante")
# MAESTRO AF toma SERIAL de "Fabr. Nº-serie" (serial del fabricante) y deja
# "Número de serie" (identificador propio de SAP) como campo separado.
MANUFACTURER_SERIAL_HEADERS = _headers("Fabr. Nº-serie")
PLAIN_SERIAL_HEADERS = _headers("Número de serie")
COMPANY_HEADERS = _headers("Sociedad")
COST_CENTER_HEADERS = _headers("Centro coste", "Centro de costo")
PEP_HEADERS = _headers("Elemento PEP")
# "Denominación" aparece 2 veces en IH08: la del Equipo (justo después de
# "Fabr. Nº-serie") y la del Material (justo después de la columna "Material").
# _find_col toma la primera coincidencia = la del equipo.
DESCRIPTION_HEADERS = _headers("Denominación")
PART_NUMBER_HEADERS = _headers("NºPieza fabric.", "Nº Pieza fabric.")
INVENTORY_NUMBER_HEADERS = _headers("Nº inventario")
LOCATION_HEADERS = _headers("Local")
TECHNICAL_LOCATION_HEADERS = _headers("Ubicac.técnica")
SUB_NUMBER_HEADERS = _headers("Subnúmero")
EMPLACEMENT_HEADERS = _headers("Emplazamiento")
MATERIAL_HEADERS = _headers("Material")
MODIFIED_AT_HEADERS = _headers("Modificado el")
MODIFIED_BY_HEADERS = _headers("Modificado por")


def _norm_code(val) -> str:
    """SAP entrega 'Equipo'/'Activo fijo' como número — sin '.0' final."""
    s = _norm(val)
    if not s:
        return ""
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def _norm_date(val) -> Optional[str]:
    """openpyxl (data_only=True) entrega celdas de fecha como datetime — ISO
    date (sin hora, IH08 no la trae con significado) para que Node la parsee
    con `new Date()` sin depender de la zona horaria de quien corre el import."""
    if val is None or val == "":
        return None
    if hasattr(val, "date"):
        return val.date().isoformat()
    if hasattr(val, "isoformat"):
        return val.isoformat()
    return _norm(val) or None


def _find_col(headers: List[str], candidates: set) -> Optional[int]:
    for i, h in enumerate(headers):
        if h in candidates:
            return i
    return None


def _find_asset_sheet(wb):
    for ws in wb.worksheets:
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not first_row:
            continue
        headers = [_norm_header(c) for c in first_row]
        equipment_idx = _find_col(headers, EQUIPMENT_HEADERS)
        if equipment_idx is None:
            continue

        has_data = any(
            len(row) > equipment_idx and _norm(row[equipment_idx]) != ""
            for row in ws.iter_rows(min_row=2, values_only=True)
        )
        if not has_data:
            continue

        manufacturer_serial_idx = _find_col(headers, MANUFACTURER_SERIAL_HEADERS)
        plain_serial_idx = _find_col(headers, PLAIN_SERIAL_HEADERS)
        material_idx = _find_col(headers, MATERIAL_HEADERS)
        # La "Denominación" del Material es la columna siguiente a "Material"
        # en el layout real de IH08 — mismo texto de encabezado que la del
        # equipo, se distingue solo por posición.
        material_description_idx = (
            material_idx + 1
            if material_idx is not None
            and material_idx + 1 < len(headers)
            and headers[material_idx + 1] in DESCRIPTION_HEADERS
            else None
        )

        return ws, {
            "equipment": equipment_idx,
            "parentEquipment": _find_col(headers, PARENT_EQUIPMENT_HEADERS),
            "fixedAsset": _find_col(headers, FIXED_ASSET_HEADERS),
            "manufacturer": _find_col(headers, MANUFACTURER_HEADERS),
            "serial": manufacturer_serial_idx if manufacturer_serial_idx is not None else plain_serial_idx,
            "company": _find_col(headers, COMPANY_HEADERS),
            "costCenter": _find_col(headers, COST_CENTER_HEADERS),
            "pep": _find_col(headers, PEP_HEADERS),
            "description": _find_col(headers, DESCRIPTION_HEADERS),
            "partNumber": _find_col(headers, PART_NUMBER_HEADERS),
            "plainSerial": plain_serial_idx,
            "inventoryNumber": _find_col(headers, INVENTORY_NUMBER_HEADERS),
            "location": _find_col(headers, LOCATION_HEADERS),
            "technicalLocation": _find_col(headers, TECHNICAL_LOCATION_HEADERS),
            "subNumber": _find_col(headers, SUB_NUMBER_HEADERS),
            "emplacement": _find_col(headers, EMPLACEMENT_HEADERS),
            "material": material_idx,
            "materialDescription": material_description_idx,
            "modifiedAt": _find_col(headers, MODIFIED_AT_HEADERS),
            "modifiedBy": _find_col(headers, MODIFIED_BY_HEADERS),
        }
    return None


def _cell(row: tuple, idx: Optional[int]):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def parse_asset_import_file(file_bytes: bytes) -> List[Dict[str, Any]]:
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)

    found = _find_asset_sheet(wb)
    if found is None:
        raise ValueError(
            "No se encontró una hoja con columna 'Equipo' (export de equipos, ej. IH08)"
        )
    ws, cols = found

    rows_out: List[Dict[str, Any]] = []
    for raw_row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if all(v is None or _norm(v) == "" for v in row):
            continue

        row_number = raw_row_idx + 2
        errors: List[str] = []

        sap_equipment_code = _norm_code(_cell(row, cols["equipment"]))
        if not sap_equipment_code:
            errors.append("Equipo (código SAP) vacío")

        parent_sap_equipment_code = _norm_code(_cell(row, cols["parentEquipment"])) or None
        sap_fixed_asset_number = _norm_code(_cell(row, cols["fixedAsset"])) or None
        brand = _norm(_cell(row, cols["manufacturer"])) or None
        serial_number = _norm(_cell(row, cols["serial"])) or None
        sap_company = _norm(_cell(row, cols["company"])) or None
        sap_cost_center = _norm(_cell(row, cols["costCenter"])) or None
        sap_pep_element = _norm(_cell(row, cols["pep"])) or None

        description = _norm(_cell(row, cols["description"])) or None
        part_number = _norm(_cell(row, cols["partNumber"])) or None
        internal_serial_number = _norm(_cell(row, cols["plainSerial"])) or None
        sap_inventory_number = _norm(_cell(row, cols["inventoryNumber"])) or None
        sap_location = _norm(_cell(row, cols["location"])) or None
        sap_technical_location = _norm(_cell(row, cols["technicalLocation"])) or None
        sub_number = _norm(_cell(row, cols["subNumber"])) or None
        emplacement_code = _norm(_cell(row, cols["emplacement"])) or None
        material_code = _norm_code(_cell(row, cols["material"])) or None
        material_description = _norm(_cell(row, cols["materialDescription"])) or None
        sap_modified_at = _norm_date(_cell(row, cols["modifiedAt"]))
        sap_modified_by = _norm(_cell(row, cols["modifiedBy"])) or None

        rows_out.append({
            "rowIndex": row_number,
            "sapEquipmentCode": sap_equipment_code,
            "parentSapEquipmentCode": parent_sap_equipment_code,
            "sapFixedAssetNumber": sap_fixed_asset_number,
            "brand": brand,
            "serialNumber": serial_number,
            "sapCostCenter": sap_cost_center,
            "sapPepElement": sap_pep_element,
            "description": description,
            "partNumber": part_number,
            "internalSerialNumber": internal_serial_number,
            "sapInventoryNumber": sap_inventory_number,
            "sapCompany": sap_company,
            "sapLocation": sap_location,
            "sapTechnicalLocation": sap_technical_location,
            "subNumber": sub_number,
            "emplacementCode": emplacement_code,
            "materialCode": material_code,
            "materialDescription": material_description,
            "sapModifiedAt": sap_modified_at,
            "sapModifiedBy": sap_modified_by,
            "parseErrors": errors,
        })

    return rows_out
