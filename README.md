# Export Files Microservice

> ⚠️ **README OBSOLETO (2026-06-24).** Este documento describe un esqueleto genérico (`pdf_service.py`/`doc_service.py`/`xls_service.py`, endpoints `/api/export/*`, scripts `install-*.bat`) que **ya no existe**. El servicio real está organizado por dominio (LORA / HV / RDP / VACATIONS) con auth `x-api-key`. **Fuente de verdad: [CONTEXT.md](./CONTEXT.md) y [ESTADO.md](./ESTADO.md).** Pendiente reescribir o eliminar este README.

Microservicio para exportar datos en diferentes formatos de archivo (PDF, DOCX, XLSX).

## Características

- ✅ Exportación a PDF usando fpdf2
- ✅ Exportación a DOCX usando python-docx
- ✅ Exportación a XLSX usando openpyxl
- ✅ API REST con FastAPI
- ✅ Documentación automática con Swagger
- ✅ Contenedorización con Docker
- ✅ Validación de datos con Pydantic

## Instalación

### Opción 1: Instalación Automática (Windows)

```bash
# 1. Instalar dependencias compatibles
install-deps.bat

# 2. Ejecutar el microservicio
start.bat
```

### Opción 2: Instalación de Emergencia (Si hay problemas de compilación)

```bash
# Para problemas con Rust/compilación de Pydantic
install-emergency.bat
```

### Opción 3: Instalación Manual

```bash
# Actualizar pip
python -m pip install --upgrade pip

# Instalar versiones compatibles (sin Rust)
pip install "fastapi>=0.100.0,<0.105.0" "uvicorn>=0.20.0,<0.25.0" "pydantic>=1.10.0,<2.0.0" fpdf2 python-docx openpyxl requests

# Ejecutar la aplicación
python -m uvicorn app.main:app --host 0.0.0.0 --port 7000 --reload
```

### Opción 4: Usando Docker (Recomendado para Producción)

```bash
# Construir la imagen
docker build -t export-files-microservice .

# Ejecutar el contenedor
docker run -p 8000:8000 export-files-microservice
```

### Solución de Problemas

Si tienes problemas con la instalación de dependencias en Windows:

1. **Actualiza Python**: Asegúrate de tener Python 3.8+ instalado
2. **Problema con Rust/Pydantic**: Usa `install-emergency.bat` para instalar Pydantic v1
3. **Usa wheels precompilados**: `pip install --only-binary=all <paquete>`
4. **Instala dependencias una por una** usando `install-deps.bat`
5. **Para errores de compilación**: Usa versiones compatibles sin Rust

## Uso

### Endpoints Principales

- `GET /` - Información del servicio
- `GET /api/v1/health` - Estado del servicio
- `POST /api/v1/export` - Exportar archivo en cualquier formato
- `POST /api/v1/export/pdf` - Exportar específicamente a PDF
- `POST /api/v1/export/docx` - Exportar específicamente a DOCX
- `POST /api/v1/export/xlsx` - Exportar específicamente a XLSX
- `GET /api/v1/formats` - Obtener formatos soportados
- `GET /api/v1/examples` - Obtener ejemplos de datos válidos
- `POST /api/v1/test/{format}` - Generar archivo de prueba (pdf/docx/xlsx)

### Ejemplo de Uso

#### Exportar una tabla a PDF (Endpoint general)

```json
POST /api/v1/export
{
  "file_format": "pdf",
  "filename": "mi_reporte",
  "data": {
    "title": "Reporte de Ventas",
    "headers": ["Producto", "Cantidad", "Precio"],
    "rows": [
      ["Laptop", 10, 1200.00],
      ["Mouse", 50, 25.00],
      ["Teclado", 30, 45.00]
    ]
  }
}
```

#### Exportar documento con contenido a DOCX (Endpoint específico)

```json
POST /api/v1/export/docx
{
  "title": "Mi Documento",
  "content": [
    "Este es el primer párrafo del documento.",
    "Aquí va el segundo párrafo con más información.",
    "Y este es el tercer párrafo de cierre."
  ],
  "tables": [
    {
      "title": "Datos de Ejemplo",
      "headers": ["Columna 1", "Columna 2"],
      "rows": [["Dato 1", "Dato 2"], ["Dato 3", "Dato 4"]]
    }
  ]
}
```

#### Exportar múltiples tablas a Excel

```json
POST /api/v1/export/xlsx
{
  "tables": [
    {
      "title": "Ventas Q1",
      "headers": ["Mes", "Ventas"],
      "rows": [["Enero", 1000], ["Febrero", 1200], ["Marzo", 1100]]
    },
    {
      "title": "Ventas Q2", 
      "headers": ["Mes", "Ventas"],
      "rows": [["Abril", 1300], ["Mayo", 1400], ["Junio", 1250]]
    }
  ]
}
```

## Estructura del Proyecto

```
app/
├── __init__.py
├── main.py                 # Aplicación principal FastAPI
├── api/
│   ├── __init__.py
│   └── routes.py          # Definición de rutas y endpoints
├── models/
│   ├── __init__.py
│   └── ExportModel.py     # Modelos Pydantic para validación
├── services/
│   ├── __init__.py
│   ├── base.py           # Clase base abstracta para servicios
│   ├── pdf_service.py    # Servicio para generar PDFs
│   ├── doc_service.py    # Servicio para generar DOCX
│   └── xls_service.py    # Servicio para generar XLSX
└── utils/
    ├── __init__.py
    └── file_utils.py     # Utilidades para manejo de archivos
```

## Documentación de la API

Una vez ejecutado el servicio, la documentación interactiva estará disponible en:

- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Configuración

El servicio se ejecuta por defecto en el puerto 8000. Puedes cambiar esto modificando el archivo `main.py` o usando variables de entorno.

## Contribuir

1. Fork del repositorio
2. Crear una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crear un Pull Request

## Licencia

Este proyecto está bajo la licencia MIT.
