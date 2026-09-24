import re
import unicodedata
from typing import List, Optional


def norm(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def norm_header(val) -> str:
    """Minúsculas, sin tildes ni símbolos, sin espacios: solo letras/dígitos.

    Los indicadores ordinales (º, ª) se eliminan ANTES de NFKD: NFKD descompone
    "º" en "o", y "Nº Pieza fabric." nunca coincidiría con "n pieza fabric".
    Colapsar los espacios también hace inmune la detección a variantes de
    puntuación ("Fabr. Nº-serie" / "Fabr Nº serie").
    """
    s = norm(val).lower().replace("º", "").replace("ª", "").replace("°", "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s)


def headers(*labels: str) -> set:
    return {norm_header(label) for label in labels}


def find_col(row_headers: List[str], candidates: set) -> Optional[int]:
    for i, h in enumerate(row_headers):
        if h in candidates:
            return i
    return None
