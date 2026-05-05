from fpdf import FPDF
import io
from pathlib import Path
from typing import Any, Dict, Optional
from datetime import datetime, date


class HvPdfService:

    # Paleta corporativa Erazo Valencia
    _NAV_BG = (0, 48, 87)          # Azul marino — header
    _SEC_BG = (215, 228, 242)       # Azul claro — cabecera de sección
    _ROW_ALT = (245, 248, 252)      # Fila alternada
    _ROW_WH = (255, 255, 255)       # Fila blanca
    _TXT_DARK = (30, 30, 30)        # Texto principal
    _TXT_LABEL = (80, 80, 80)       # Etiqueta de campo
    _TXT_SEC = (0, 48, 87)          # Texto de sección
    _TXT_WHITE = (255, 255, 255)
    _TXT_HEADER_SUB = (180, 205, 225)  # Subtítulo en header

    _PAGE_W = 180   # Ancho útil A4 con márgenes 15mm
    _LABEL_W = 68
    _VALUE_W = 112  # _PAGE_W - _LABEL_W
    _ROW_H = 7
    _SEC_H = 8

    def __init__(self):
        self.pdf = FPDF(orientation="P", unit="mm", format="A4")
        self.pdf.set_margins(15, 15, 15)
        self.pdf.set_auto_page_break(auto=True, margin=20)

        BASE_DIR = Path(__file__).resolve().parents[3]
        FONT_DIR = BASE_DIR / "fonts"
        normal = str(FONT_DIR / "DejaVuSerif.ttf")
        bold_path = FONT_DIR / "DejaVuSerif-Bold.ttf"
        bold = str(bold_path) if bold_path.exists() else normal

        self.pdf.add_font("DejaVu", "", normal, uni=True)
        self.pdf.add_font("DejaVu", "B", bold, uni=True)
        self.pdf.add_font("DejaVu", "I", normal, uni=True)

    def generate_pdf(self, employee: Dict[str, Any]) -> bytes:
        self.pdf.add_page()
        self._render_header(employee)
        self._render_section("IDENTIFICACIÓN", [
            ("Tipo de documento", employee.get("documentType")),
            ("Número de identificación", employee.get("identification")),
            ("Lugar de expedición", employee.get("idExpeditionPlace")),
            ("Fecha de expedición", self._fmt_date(employee.get("idExpeditionDate"))),
        ])
        self._render_section("DATOS PERSONALES", [
            ("Género", employee.get("genre")),
            ("Fecha de nacimiento", self._fmt_date(employee.get("birthDate"))),
            ("Lugar de nacimiento", employee.get("placeOfBirth")),
            ("Nacionalidad", employee.get("nationality")),
            ("Estado civil", employee.get("maritalStatus")),
            ("Grupo sanguíneo (RH)", employee.get("rh")),
        ])
        self._render_section("INFORMACIÓN DE CONTACTO", [
            ("Teléfono", self._fmt_phone(employee)),
            ("Correo electrónico", employee.get("email")),
            ("Dirección", employee.get("address")),
            ("Departamento / Ciudad", employee.get("department")),
        ])
        self._render_section("FORMACIÓN ACADÉMICA", [
            ("Nivel de educación", employee.get("educationLevel")),
            ("Profesión / Título", employee.get("profession")),
        ])
        self._render_section("CARGO EN ERAZO VALENCIA", [
            ("Cargo", employee.get("position")),
            ("Línea de servicio", employee.get("serviceLineName")),
            ("Grupo empresarial", employee.get("reportGroup")),
        ])
        self._render_footer()
        return bytes(self.pdf.output())

    # ─── Secciones ────────────────────────────────────────────────────────────

    def _render_header(self, employee: Dict[str, Any]) -> None:
        pdf = self.pdf
        lm = pdf.l_margin
        y0 = pdf.get_y()
        h = 50

        # Fondo navy
        pdf.set_fill_color(*self._NAV_BG)
        pdf.rect(lm, y0, self._PAGE_W, h, "F")

        # "ERAZO VALENCIA" — top right
        pdf.set_xy(lm, y0 + 5)
        pdf.set_font("DejaVu", "B", 9)
        pdf.set_text_color(*self._TXT_HEADER_SUB)
        pdf.cell(self._PAGE_W, 5, "ERAZO VALENCIA", align="R")

        # Nombre completo — centrado, grande
        full_name = (
            f"{employee.get('name', '')} {employee.get('lastName', '')}".strip().upper()
        )
        pdf.set_xy(lm, y0 + 13)
        pdf.set_font("DejaVu", "B", 16)
        pdf.set_text_color(*self._TXT_WHITE)
        pdf.cell(self._PAGE_W, 10, full_name, align="C")

        # Cargo · Línea de servicio
        position = employee.get("position") or ""
        service_line = employee.get("serviceLineName") or ""
        if position and service_line:
            subtitle = f"{position}  ·  {service_line}"
        else:
            subtitle = position or service_line or ""
        if subtitle:
            pdf.set_xy(lm, y0 + 27)
            pdf.set_font("DejaVu", "", 10)
            pdf.set_text_color(*self._TXT_HEADER_SUB)
            pdf.cell(self._PAGE_W, 6, subtitle, align="C")

        # Tipo doc + número
        doc_type = employee.get("documentType") or "CC"
        identification = employee.get("identification") or ""
        pdf.set_xy(lm, y0 + 37)
        pdf.set_font("DejaVu", "", 9)
        pdf.set_text_color(150, 185, 215)
        pdf.cell(self._PAGE_W, 5, f"{doc_type}: {identification}", align="C")

        # Reposicionar cursor bajo el header
        pdf.set_xy(lm, y0 + h + 6)
        pdf.set_text_color(*self._TXT_DARK)

    def _render_section(self, title: str, fields: list) -> None:
        pdf = self.pdf

        # Cabecera de sección
        pdf.set_fill_color(*self._SEC_BG)
        pdf.set_text_color(*self._TXT_SEC)
        pdf.set_font("DejaVu", "B", 10)
        pdf.cell(self._PAGE_W, self._SEC_H, f"  {title}", ln=True, fill=True)

        # Filas de datos
        for i, (label, raw_value) in enumerate(fields):
            value = str(raw_value) if raw_value not in (None, "", "None") else "—"
            fill_color = self._ROW_ALT if i % 2 == 0 else self._ROW_WH
            pdf.set_fill_color(*fill_color)

            pdf.set_font("DejaVu", "B", 9)
            pdf.set_text_color(*self._TXT_LABEL)
            pdf.cell(self._LABEL_W, self._ROW_H, f"  {label}", fill=True)

            pdf.set_font("DejaVu", "", 9)
            pdf.set_text_color(*self._TXT_DARK)
            pdf.cell(self._VALUE_W, self._ROW_H, f"  {value}", ln=True, fill=True)

        pdf.ln(5)

    def _render_footer(self) -> None:
        pdf = self.pdf
        pdf.set_y(-18)
        # Línea separadora
        lm = pdf.l_margin
        y = pdf.get_y()
        pdf.set_draw_color(180, 200, 220)
        pdf.line(lm, y, lm + self._PAGE_W, y)
        pdf.ln(3)
        generated = datetime.now().strftime("%Y-%m-%d %H:%M")
        pdf.set_font("DejaVu", "I", 8)
        pdf.set_text_color(*self._TXT_LABEL)
        pdf.cell(self._PAGE_W / 2, 5, "DOCUMENTO CONFIDENCIAL — Uso exclusivo Erazo Valencia", align="L")
        pdf.cell(self._PAGE_W / 2, 5, f"Generado por VALERA  ·  {generated}", ln=True, align="R")

    # ─── Helpers ──────────────────────────────────────────────────────────────

    def _fmt_date(self, value: Optional[str]) -> Optional[str]:
        if not value:
            return None
        try:
            return date.fromisoformat(str(value)[:10]).strftime("%d/%m/%Y")
        except Exception:
            return value

    def _fmt_phone(self, employee: Dict[str, Any]) -> Optional[str]:
        phone = employee.get("phone")
        if not phone:
            return None
        country = employee.get("countryCode") or ""
        return f"{country} {phone}".strip() if country else phone
