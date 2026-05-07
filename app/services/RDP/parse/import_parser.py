import io
import openpyxl
from typing import List, Dict, Any, Optional

# Columnas esperadas en la plantilla (índice 0-based)
COL_IDENTIFICACION = 0
COL_ES_AUSENCIA    = 1
COL_TURNO          = 2
COL_TIPO_AUSENCIA  = 3
COL_TIPO_BONO      = 4
COL_PROYECTO       = 5
COL_HORA_EXTRA     = 6
COL_NOTAS          = 7

# Filas de encabezado a saltar cuando NO hay empleados pre-rellenos (título + instrucción + header + hint + ejemplo)
# Cuando hay empleados, los datos empiezan en fila 5 (sin fila de ejemplo), así que skip = 4
SKIP_ROWS_DEFAULT    = 5   # sin empleados: fila ejemplo en row 5, datos desde row 6
SKIP_ROWS_WITH_EMPS  = 4   # con empleados: datos desde row 5


def _str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _bool_field(val) -> bool:
    s = _str(val).lower()
    return s in ("si", "sí", "yes", "1", "true", "s")


def _float_field(val) -> Optional[float]:
    if val is None or _str(val) == "":
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _detect_skip_rows(ws) -> int:
    """
    Detecta cuántas filas de encabezado saltar.
    Si la fila 5 col A tiene un valor numérico (cédula pre-rellena), los datos empiezan en row 5.
    Si la fila 5 es la fila ejemplo (valor de ejemplo), los datos empiezan en row 6.
    """
    cell_a5 = ws.cell(row=5, column=1).value
    if cell_a5 is not None and _str(cell_a5) not in ("", "1069742877"):
        return SKIP_ROWS_WITH_EMPS   # datos desde row 5
    return SKIP_ROWS_DEFAULT         # datos desde row 6


def parse_import_file(file_bytes: bytes) -> List[Dict[str, Any]]:
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)

    # Use first sheet regardless of name
    ws = wb.worksheets[0]
    skip = _detect_skip_rows(ws)
    rows_out = []

    for raw_row_idx, row in enumerate(ws.iter_rows(min_row=skip + 1, values_only=True)):
        # Skip fully empty rows
        if all(v is None or _str(v) == "" for v in row):
            continue

        row_number = raw_row_idx + skip + 1
        errors: List[str] = []

        identificacion = _str(row[COL_IDENTIFICACION] if len(row) > COL_IDENTIFICACION else None)
        es_ausencia    = _bool_field(row[COL_ES_AUSENCIA]    if len(row) > COL_ES_AUSENCIA    else None)
        turno          = _str(row[COL_TURNO]          if len(row) > COL_TURNO          else None) or None
        tipo_ausencia  = _str(row[COL_TIPO_AUSENCIA]  if len(row) > COL_TIPO_AUSENCIA  else None) or None
        tipo_bono      = _str(row[COL_TIPO_BONO]      if len(row) > COL_TIPO_BONO      else None) or None
        proyecto       = _str(row[COL_PROYECTO]        if len(row) > COL_PROYECTO        else None) or None
        hora_extra     = _float_field(row[COL_HORA_EXTRA] if len(row) > COL_HORA_EXTRA else None)
        notas          = _str(row[COL_NOTAS]           if len(row) > COL_NOTAS           else None) or None

        if not identificacion:
            errors.append("Identificación es obligatoria")

        if not es_ausencia and not turno:
            errors.append("Turno es obligatorio cuando Es Ausencia = No")

        if es_ausencia and not tipo_ausencia:
            errors.append("Tipo Ausencia es obligatorio cuando Es Ausencia = Si")

        rows_out.append({
            "rowIndex":     row_number,
            "identificacion": identificacion,
            "esAusencia":   es_ausencia,
            "turno":        turno,
            "tipoAusencia": tipo_ausencia,
            "tipoBono":     tipo_bono,
            "proyecto":     proyecto,
            "horaExtra":    hora_extra,
            "notas":        notas,
            "parseErrors":  errors,
        })

    return rows_out
