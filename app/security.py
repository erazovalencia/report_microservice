import os
from fastapi import Header, HTTPException


def require_api_key(x_api_key: str = Header(..., alias="x-api-key")) -> None:
    expected = os.environ.get("EXPORT_SERVICE_API_KEY", "")
    if not expected:
        raise HTTPException(status_code=500, detail="EXPORT_SERVICE_API_KEY not configured")
    if x_api_key != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")
