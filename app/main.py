from dotenv import load_dotenv
load_dotenv()

import os
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn
from contextlib import asynccontextmanager

from .api.routes import router as export_router
from .api.hv_routes import router as hv_router
from .api.rdp_routes import router as rdp_router
from .api.vacations_routes import router as vacations_router

@asynccontextmanager
async def lifespan(app: FastAPI):
    if not os.environ.get("EXPORT_SERVICE_API_KEY"):
        raise RuntimeError("EXPORT_SERVICE_API_KEY no configurada — el servicio no puede arrancar")
    print("Iniciando microservicio de creacion de reportes...")
    yield
    print("Cerrando microservicio de creacion de reportes...")


# Crear aplicación FastAPI
app = FastAPI(
    title="REPORT MICROSERVICE",
    description="Microservicio para exportar datos en diferentes formatos (PDF, DOCX, XLSX)",
    version="1.0.0",
    lifespan=lifespan
)

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # En producción, especificar orígenes específicos
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Incluir rutas
app.include_router(export_router, prefix="/api/v1", tags=["export"])
app.include_router(hv_router, prefix="/hv", tags=["hv"])
app.include_router(rdp_router, prefix="/rdp", tags=["rdp"])
app.include_router(vacations_router, prefix="/vacations", tags=["vacations"])


@app.get("/")
async def root():
    """Endpoint raíz"""
    return {
        "message": "REPORT generate microservice",
        "version": "1.0.0",
        "status": "running",
        "docs": "/docs",
        "health": "/api/v1/health"
    }


@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc: HTTPException):
    """Manejo personalizado de excepciones HTTP"""
    return JSONResponse(
        status_code=exc.status_code,
        content={
            "error": True,
            "message": exc.detail,
            "status_code": exc.status_code
        }
    )


@app.exception_handler(Exception)
async def general_exception_handler(request, exc: Exception):
    """Manejo personalizado de excepciones generales"""
    return JSONResponse(
        status_code=500,
        content={
            "error": True,
            "message": "Error interno del servidor",
            "status_code": 500
        }
    )


if __name__ == "__main__":
    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )
