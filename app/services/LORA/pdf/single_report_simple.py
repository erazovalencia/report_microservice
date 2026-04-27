from fpdf import FPDF
import io
from typing import Any, Dict, List
from datetime import datetime
from pathlib import Path
from ...base import BaseExportService


class ExportSinglePDFReportSimple(BaseExportService):
    """Genera PDF con diseno institucional serio y limpio"""

    FIELDS = [
        ("id", "ID de reporte"),
        ("userId", "ID de usuario"),
        ("user.documentId", "Documento de usuario"),
        ("user.userInformation.name", "Nombre del usuario"),
        ("user.userInformation.lastName", "Apellido del usuario"),
        ("externalNameUser", "Nombre externo del usuario"),
        ("externalOrganization", "Organizacion externa"),
        ("reportTitle", "Titulo del reporte"),
        ("conversation", "Conversacion"),
        ("base", "Base"),
        ("createdAt", "Fecha de creacion"),
        ("updatedAt", "Fecha de actualizacion"),
        ("unity", "Unidad"),
        ("rig", "Equipo (rig)"),
        ("project", "Proyecto"),
        ("field", "Campo"),
        ("reportType", "Tipo de reporte"),
        ("hazardClassification", "Clasificacion del peligro"),
        ("hazardType", "Tipo de peligro"),
        ("detailedDescription", "Descripcion detallada"),
        ("findingCause", "Causa del hallazgo"),
        ("reportEvidence", "Evidencias del reporte"),
        ("actions", "Acciones"),
        ("reportStatus", "Estado del reporte"),
        ("loraReportCode", "Codigo de reporte LORA"),
    ]

    def __init__(self):
        self.pdf = FPDF(orientation="P", unit="mm", format="A4")
        # Registrar fuentes Unicode (DejaVu) para acentos/caracteres especiales
        BASE_DIR = Path(__file__).resolve().parents[3]
        FONT_DIR = BASE_DIR / "fonts"
        normal = FONT_DIR / "DejaVuSerif.ttf"
        bold = FONT_DIR / "DejaVuSerif-Bold.ttf"
        self.pdf.add_font("DejaVu", "", str(normal), uni=True)
        if bold.exists():
            self.pdf.add_font("DejaVu", "B", str(bold), uni=True)
        else:
            self.pdf.add_font("DejaVu", "B", str(normal), uni=True)
        # Register italic (fallback to normal)
        self.pdf.add_font("DejaVu", "I", str(normal), uni=True)

    def generate_file(self, data: Any, options: Dict = None) -> io.BytesIO:
        if not self.validate_data(data):
            raise ValueError("Datos no validos")

        self.pdf.add_page()
        self.pdf.set_auto_page_break(auto=True, margin=15)
        self.pdf.set_font("DejaVu", "", 11)
        self.pdf.set_text_color(0, 0, 0)

        self._render_header(data)
        self._render_section("Detalles del reporte", self._format_fields(data))
        self._render_section("Acciones", self._format_actions(data.get("actions", [])))
        self._render_footer(data)

        buffer = io.BytesIO()
        buffer.write(self.pdf.output())

        buffer.seek(0)
        return buffer

    def _render_header(self, data: Dict):
        self.pdf.set_font("DejaVu", "B", 14)
        self.pdf.cell(0, 8, f"Reporte LORA - {data.get('reportTitle', 'N/A')}", ln=True, align="C")
        self.pdf.set_font("DejaVu", "B", 10)
        self.pdf.cell(0, 8, f"ID de reporte: {data.get('id', 'N/A')}", ln=True, align="L")
        self.pdf.set_font("DejaVu", "", 10)
        self.pdf.cell(0, 6, f"Codigo: {data.get('loraReportCode', 'N/A')}", ln=True)
        self.pdf.cell(0, 6, f"Estado: {data.get('reportStatus', 'N/A').upper()}", ln=True)
        self.pdf.cell(0, 6, f"Fecha de creacion: {data.get('createdAt', 'N/A')}", ln=True)
        self._draw_separator()

    def _render_section(self, title: str, content: str):
        self.pdf.set_font("DejaVu", "B", 12)
        self.pdf.cell(0, 7, title.upper(), ln=True)
        self.pdf.set_font("DejaVu", "", 10)
        self.pdf.multi_cell(0, 6, str(content).strip() or "N/A")
        self._draw_separator()

    def _render_footer(self, data: Dict):
        self.pdf.set_y(-30)
        self._draw_separator()
        self.pdf.set_font("DejaVu", "I", 8)
        self.pdf.cell(0, 5, f"ID de reporte: {data.get('id', 'N/A')}", ln=True, align="R")
        self.pdf.cell(0, 5, f"Exportado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="R")
        self.pdf.cell(0, 5, "Documento generado automaticamente por VALERA ECOSYSTEM", ln=True, align="C")

    def _draw_separator(self):
        self.pdf.ln(2)
        self.pdf.cell(0, 0, "-" * 120, ln=True)
        self.pdf.ln(3)

    def _format_fields(self, data: Dict) -> str:
        lines = []
        for key, label in self.FIELDS:
            if key == "actions":
                value = self._format_actions(data.get("actions", []))
            else:
                value = self._format_value(self._get_nested_value(data, key))
            cleaned = value if value not in (None, "") else "N/A"
            lines.append(f"{label}: {cleaned}")
        return "\n\n".join(lines)

    def _format_actions(self, actions: List[Dict]) -> str:
        if not actions:
            return "Sin acciones registradas."
        result = []
        for i, act in enumerate(actions, 1):
            desc = act.get("description", "N/A")
            resp = act.get("responsible", "N/A")
            due = act.get("dueDate", "N/A")
            status = act.get("status", "N/A").upper()
            result.append(f"{i}. {desc}\n   Responsable: {resp}\n   Fecha limite: {due}\n   Estado: {status}\n")
        return "\n".join(result)

    def _format_value(self, value: Any) -> Any:
        if value is None:
            return "N/A"
        if isinstance(value, list):
            formatted_items = []
            for item in value:
                if isinstance(item, dict):
                    parts = [f"{k}: {v}" for k, v in item.items() if v not in (None, "")]
                    formatted_items.append(", ".join(parts) if parts else str(item))
                else:
                    formatted_items.append(str(item))
            return "\n".join(formatted_items) if formatted_items else "N/A"
        if isinstance(value, dict):
            parts = [f"{k}: {v}" for k, v in value.items() if v not in (None, "")]
            return ", ".join(parts) if parts else "N/A"
        return value

    def _get_nested_value(self, obj: Dict, key: str) -> Any:
        parts = key.split('.')
        value = obj
        for part in parts:
            if isinstance(value, dict):
                value = value.get(part, None)
            else:
                value = None
            if value is None:
                break
        if value is None and key == "reportEvidence":
            evidences = obj.get("evidence", [])
            return evidences if evidences else None
        if value is None and key == "actions":
            return obj.get("actions", [])
        return value

    def get_content_type(self) -> str:
        return "application/pdf"

    def get_file_extension(self) -> str:
        return ".pdf"
