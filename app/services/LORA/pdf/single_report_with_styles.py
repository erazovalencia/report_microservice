from fpdf import FPDF
import io
import os
from typing import Any, Dict, List
from datetime import datetime
from ...base import BaseExportService


class ExportSinglePDFReportWithStyle(BaseExportService):
    """Genera PDF de un reporte LORA con estilo visual."""

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
        ("reportStatus", "Estado del reporte"),
        ("loraReportCode", "Codigo de reporte LORA"),
    ]

    def __init__(self):
        self.pdf = None
        self.logo_path = None
        self.colors = {
            'background': (247, 243, 233),
            'border': (224, 207, 163),
            'text_dark': (45, 36, 24),
            'text_brown': (122, 92, 27),
            'open_stamp_bg': (224, 112, 112),
            'close_stamp_bg': (127, 186, 0),
            'open_stamp_border': (224, 26, 26),
            'close_stamp_border': (170, 240, 18),
            'block_bg': (255, 251, 231),
            'summary_bg': (244, 236, 217),
            'action_open': (224, 112, 112),
            'action_close': (198, 223, 144),
        }

    def generate_file(self, data: Any, options: Dict = None) -> io.BytesIO:
        if not isinstance(data, dict):
            raise ValueError("Se requiere un diccionario con los datos del reporte")

        report_data = data
        # Permitir pasar el logo por options o por los datos, y usar un default
        if isinstance(options, dict):
            self.logo_path = options.get("logo_path") or options.get("logo")
        if not self.logo_path and isinstance(report_data, dict):
            self.logo_path = report_data.get("logo_path") or report_data.get("logo")
        if not self.logo_path:
            # Ruta por defecto indicada: common/logo.png
            self.logo_path = "common/logo.png"
        self.pdf = FPDF(orientation="P", unit="mm", format="A4")
        self.pdf.set_auto_page_break(auto=True, margin=15)
        self.pdf.add_page()
        self.pdf.set_fill_color(*self.colors['background'])
        self.pdf.rect(0, 0, 210, 297, 'F')
        self.pdf.set_font("Courier", "", 11)
        self.pdf.set_text_color(*self.colors['text_dark'])

        self._draw_stamp(report_data)
        self._draw_header(report_data)
        self._draw_summary(report_data)
        self._draw_description(report_data)
        self._draw_conversation(report_data)
        self._draw_evidences(report_data)
        self._draw_actions(report_data)
        self._draw_footer(report_data)

        buffer = io.BytesIO()
        pdf_output = self.pdf.output(dest='S')
        buffer.write(pdf_output if isinstance(pdf_output, (bytes, bytearray)) else pdf_output.encode('latin-1'))
        buffer.seek(0)
        return buffer

    # Visual components
    def _draw_stamp(self, data: Dict):
        status = str(data.get("reportStatus", "open")).lower()
        stamp_text = "ABIERTO" if status == "open" else "CERRADO"
        if status == "open":
            bg = self.colors['open_stamp_bg']
            border = self.colors['open_stamp_border']
        else:
            bg = self.colors['close_stamp_bg']
            border = self.colors['close_stamp_border']
        self.pdf.set_xy(150, 15)
        self.pdf.set_fill_color(*bg)
        self.pdf.set_draw_color(*border)
        self.pdf.set_text_color(0, 0, 0)
        self.pdf.set_font("Courier", "B", 14)
        self.pdf.cell(45, 10, stamp_text, border=1, align='C', fill=True)

    def _draw_header(self, data: Dict):
        code = data.get("loraReportCode", "Sin codigo")
        title = data.get("reportTitle", "Sin titulo")
        created = data.get("createdAt", datetime.now().strftime("%Y-%m-%d"))
        # Dibuja el logo centrado encima del código si está disponible
        placed_y = self._draw_logo()
        if placed_y:
            self.pdf.set_y(max(self.pdf.get_y(), placed_y))
        else:
            self.pdf.ln(20)
        self.pdf.set_font("Courier", "B", 11)
        self.pdf.set_text_color(*self.colors['text_brown'])
        self.pdf.cell(0, 8, f"CODIGO: {code}", ln=True)
        self.pdf.set_font("Courier", "B", 18)
        self.pdf.set_text_color(*self.colors['text_dark'])
        self.pdf.cell(0, 10, title.upper(), ln=True)
        self.pdf.set_font("Courier", "", 9)
        self.pdf.set_text_color(*self.colors['text_brown'])
        self.pdf.cell(0, 8, f"Creado: {created}", ln=True)
        self._draw_line()

    def _draw_summary(self, data: Dict):
        field_items = [
            (label, self._format_value(self._get_nested_value(data, key)))
            for key, label in self.FIELDS
        ]
        self._draw_block("Detalles del reporte", self._render_summary_grid, field_items)

    def _draw_description(self, data: Dict):
        if data.get("detailedDescription"):
            self._draw_block("Descripcion", self._render_paragraph, data["detailedDescription"])
        if data.get("findingCause"):
            self._draw_block("Causa", self._render_paragraph, data["findingCause"])

    def _draw_conversation(self, data: Dict):
        if data.get("conversation"):
            self._draw_block("Conversacion", self._render_paragraph, data["conversation"])

    def _draw_evidences(self, data: Dict):
        evidences = data.get("evidence", [])
        if evidences:
            self._draw_block("Evidencias", self._render_evidences, evidences)
        else:
            self._draw_block("Evidencias", self._render_paragraph, "Sin evidencias adjuntas")

    def _draw_actions(self, data: Dict):
        actions = data.get("actions", [])
        if not actions:
            return
        self._draw_block("Acciones", self._render_actions, actions)

    def _draw_footer(self, data: Dict):
        self.pdf.ln(10)
        self._draw_line()
        self.pdf.set_font("Courier", "I", 8)
        self.pdf.set_text_color(*self.colors['text_brown'])
        self.pdf.cell(0, 6, f"ID: {data.get('id', 'N/A')}", ln=True, align="R")
        self.pdf.cell(0, 5, "Documento generado automaticamente por VALERA ECOSYSTEM", ln=True, align="C")

    # Blocks / sections helpers
    def _draw_block(self, title: str, render_fn, content):
        y = self.pdf.get_y() + 4
        self.pdf.set_fill_color(*self.colors['block_bg'])
        self.pdf.set_draw_color(*self.colors['border'])
        self.pdf.rect(10, y, 190, 10, 'F')
        self.pdf.set_y(y + 2)
        self.pdf.set_x(15)
        self.pdf.set_font("Courier", "B", 12)
        self.pdf.set_text_color(*self.colors['text_brown'])
        self.pdf.cell(0, 8, title.upper(), ln=True)
        self._draw_line()
        render_fn(content)
        self.pdf.ln(3)

    def _render_summary_grid(self, items: List):
        """Renderiza pares clave/valor en dos columnas con salto de línea controlado para evitar desbordes."""
        self.pdf.set_font("Courier", "", 10)
        self.pdf.set_fill_color(*self.colors['summary_bg'])
        self.pdf.set_text_color(*self.colors['text_dark'])
        col_width = 88  # ancho por columna: x=15+88=103 izq, x=110+88=198 der (< margen 200mm)
        line_height = 6

        idx = 0
        while idx < len(items):
            y_start = self.pdf.get_y()

            # Columna izquierda
            left_label, left_val = items[idx]
            left_text = f"{left_label}: {left_val}"
            self.pdf.set_xy(15, y_start)
            self.pdf.multi_cell(col_width, line_height, left_text, fill=True)
            left_height = self.pdf.get_y() - y_start

            # Columna derecha (si existe)
            right_height = 0
            if idx + 1 < len(items):
                right_label, right_val = items[idx + 1]
                right_text = f"{right_label}: {right_val}"
                self.pdf.set_xy(110, y_start)
                self.pdf.multi_cell(col_width, line_height, right_text, fill=True)
                right_height = self.pdf.get_y() - y_start

            row_height = max(left_height, right_height, line_height)
            self.pdf.set_y(y_start + row_height + 2)
            idx += 2

        self.pdf.ln(2)

    def _render_paragraph(self, text: str):
        clean = str(text).replace("\n", " ").strip()
        self.pdf.set_font("Courier", "", 10)
        self.pdf.multi_cell(0, 6, clean)
        self.pdf.ln(2)

    def _render_evidences(self, evidences: List):
        for e in evidences:
            self.pdf.set_font("Courier", "", 10)
            self.pdf.cell(0, 6, f"- {e}", ln=True)
        self.pdf.ln(2)

    def _render_actions(self, actions: List[Dict]):
        for act in actions:
            status = str(act.get("status", "open")).lower()
            bg = self.colors['action_close'] if status == "close" else self.colors['action_open']
            self.pdf.set_fill_color(*bg)
            self.pdf.set_x(15)
            self.pdf.set_font("Courier", "B", 10)
            self.pdf.multi_cell(0, 8, f"Accion: {act.get('description', 'N/A')}", border=0, fill=True)
            self.pdf.set_font("Courier", "", 9)
            resp = act.get("responsible", "No asignado")
            due = act.get("dueDate", "N/A")
            self.pdf.multi_cell(0, 6, f"Responsable: {resp}", ln=True)
            self.pdf.multi_cell(0, 6, f"Fecha limite: {due}", ln=True)
            self.pdf.ln(3)

    # Utils
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

    def _draw_line(self):
        self.pdf.set_draw_color(*self.colors['border'])
        y = self.pdf.get_y()
        self.pdf.line(10, y, 200, y)
        self.pdf.ln(4)

    def _draw_logo(self) -> int:
        """Coloca un logo centrado debajo del sello y por encima del encabezado.
        Devuelve la coordenada Y para continuar debajo del logo; 0 si no se dibuja.
        """
        try:
            if not self.logo_path or not isinstance(self.logo_path, str):
                return 0
            # Solo dibuja si la ruta existe para evitar errores de FPDF
            if os.path.isfile(self.logo_path):
                logo_w = 40  # ancho del logo en mm
                page_w = 210  # A4 portrait width
                x = (page_w - logo_w) / 2
                y = 30  # debajo del sello (que está en y~15)
                self.pdf.image(self.logo_path, x=x, y=y, w=logo_w)
                return int(y + 22)
            return 0
        except Exception:
            return 0

    def get_content_type(self) -> str:
        return "application/pdf"

    def get_file_extension(self) -> str:
        return ".pdf"
