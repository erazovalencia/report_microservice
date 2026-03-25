import os
from typing import Any, Dict, List, Iterable, Tuple, Mapping, Union
import requests


def _get_base_url() -> str:
    base = os.getenv("VALERA_API", "http://10.0.0.45:3000/api/")
    return base.rstrip("/")


def _get_headers() -> Dict[str, str]:
    token = os.getenv("VALERA_API_TOKEN", "").strip()
    if not token:
        return {}

    return {"x-valera-api-token": token}


def get_reports() -> List[Dict[str, Any]]:
    """Obtiene todos los reportes LORA desde la API VALERA"""
    url = f"{_get_base_url()}/lora-report"
    resp = requests.get(url, headers=_get_headers(), timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_report_by_id(report_id: int) -> Dict[str, Any]:
    """Obtiene un reporte LORA por ID desde la API VALERA"""
    url = f"{_get_base_url()}/lora-report/{report_id}"
    resp = requests.get(url, headers=_get_headers(), timeout=30)
    resp.raise_for_status()
    return resp.json()

def get_report_by_userId(user_id: int) -> Dict[str, Any]:
    """Obtiene los reportes LORA filtrados por usuario desde la API VALERA"""
    url = f"{_get_base_url()}/lora-report/getReportFilter"
    resp = requests.get(url, params={"userId": user_id}, headers=_get_headers(), timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_reports_by_filters(filters: Union[Mapping[str, Any], Iterable[Tuple[str, Any]]]) -> Dict[str, Any]:
    """Obtiene reportes LORA filtrados desde la API VALERA usando los query params recibidos.

    Acepta un mapeo clave-valor o una lista de tuplas para permitir claves repetidas.
    """
    url = f"{_get_base_url()}/lora-report/getReportFilter"
    # Permite timeouts más generosos para consultas con múltiples filtros
    resp = requests.get(url, params=filters, headers=_get_headers(), timeout=60)
    print(resp)
    resp.raise_for_status()
    return resp.json()
