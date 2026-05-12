from fpdf import FPDF
import io
import unicodedata
from ....models.RdpModel import RdpConsolidatedRequest


def _t(text) -> str:
    """Sanitiza texto para Helvetica (Latin-1/Windows-1252).
    Mantiene todos los caracteres del español (á é í ó ú ñ ü Á É Í Ó Ú Ñ Ü).
    Los que queden fuera del rango Latin-1 se descomponen al equivalente ASCII más cercano.
    """
    s = unicodedata.normalize("NFC", str(text) if text is not None else "")
    result = []
    for ch in s:
        try:
            ch.encode("latin-1")
            result.append(ch)
        except UnicodeEncodeError:
            decomposed = unicodedata.normalize("NFKD", ch).encode("ascii", "ignore").decode("ascii")
            result.append(decomposed if decomposed else "")
    return "".join(result)

# A4 landscape: 297 x 210 mm
PAGE_W = 297
PAGE_H = 210
MARGIN = 8

NAVY = (31, 56, 100)
WHITE = (255, 255, 255)
ALT  = (235, 243, 251)
YELLOW = (255, 242, 204)
HEADER_BG = (189, 215, 238)

# Columnas: (header, width_mm, align)
COLS = [
    ("FECHA",         16, "C"),
    ("TURNO",          9, "C"),
    ("DESCRIPCION",   40, "L"),
    ("H. INGRESO",    12, "C"),
    ("H. SALIDA",     12, "C"),
    ("TOTAL\nHORAS",  10, "C"),
    ("HED",            9, "C"),
    ("HEN",            9, "C"),
    ("HEDF",           9, "C"),
    ("HENF",           9, "C"),
    ("C. COSTOS",     22, "L"),
    ("ACTIVIDAD",     37, "L"),
]

ROW_H     = 6.0   # mm por fila de datos
HEADER_H  = 8.0   # mm para encabezado de tabla


class ConsolidadoPDF(FPDF):
    def __init__(self, title: str):
        super().__init__(orientation="L", unit="mm", format="A4")
        self._report_title = title

    def header(self):
        pass  # lo manejamos manualmente

    def footer(self):
        self.set_y(-8)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(150, 150, 150)
        self.cell(0, 4, f"Pagina {self.page_no()}", align="R")


def _rgb(pdf: FPDF, color):
    pdf.set_fill_color(*color)
    pdf.set_text_color(*color)


def _draw_cell(pdf: FPDF, w, h, text, bold=False, size=8,
               fill=None, text_color=WHITE, align="C", border=1, ln=0):
    if fill:
        pdf.set_fill_color(*fill)
    pdf.set_text_color(*text_color)
    pdf.set_font("Helvetica", "B" if bold else "", size)
    pdf.cell(w, h, _t(str(text))[:60], border=border, ln=ln, align=align, fill=bool(fill))


def generate_consolidated_pdf(data: RdpConsolidatedRequest) -> io.BytesIO:
    per = data.periodo
    emp = data.empleado

    title = (
        f"REPORTE CONSOLIDADO DE TURNOS\n"
        f"PERIODO: {per.get('from','')} al {per.get('to','')}"
    )

    pdf = ConsolidadoPDF(title)
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_margins(MARGIN, MARGIN, MARGIN)

    usable_w = PAGE_W - 2 * MARGIN  # 281 mm

    # ── Título ────────────────────────────────────────────────────────────────
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(*NAVY)
    pdf.set_xy(MARGIN, MARGIN)
    pdf.multi_cell(usable_w, 5, title, align="C")
    pdf.ln(3)

    # ── Datos empleado ─────────────────────────────────────────────────────────
    label_w = 28
    val_w   = 50
    gap     = 4

    def emp_pair(label, value, x, y):
        pdf.set_xy(x, y)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(*NAVY)
        pdf.cell(label_w, 5, _t(label), border=0)
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(val_w, 5, _t(str(value))[:35], border="B")

    y0 = pdf.get_y()
    emp_pair("Identificacion:", emp.identificacion,                    MARGIN,             y0)
    emp_pair("Nombre:",         emp.nombre[:35],                       MARGIN + label_w + val_w + gap, y0)
    emp_pair("Edad:",           emp.edad,                              MARGIN + 2*(label_w+val_w+gap), y0)
    y0 += 6
    emp_pair("Cargo:",          emp.cargo[:35],                        MARGIN,             y0)
    emp_pair("Unidad Org.:",    emp.unidadOrganizativa[:30],           MARGIN + label_w + val_w + gap, y0)
    emp_pair("Sexo:",           emp.sexo,                              MARGIN + 2*(label_w+val_w+gap), y0)
    y0 += 6
    emp_pair("Tipo Nomina:",    emp.tipoNomina,                        MARGIN,             y0)
    emp_pair("Empresa:",        emp.empresa[:35],                      MARGIN + label_w + val_w + gap, y0)
    pdf.ln(10)

    # ── Encabezado de tabla ────────────────────────────────────────────────────
    y_table = pdf.get_y()
    pdf.set_xy(MARGIN, y_table)
    for label, w, _ in COLS:
        # wrap doble línea manualmente
        lines = label.split("\n")
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*WHITE)
        pdf.set_font("Helvetica", "B", 7)
        if len(lines) == 2:
            x_save = pdf.get_x()
            y_save = pdf.get_y()
            pdf.cell(w, HEADER_H / 2, lines[0], border="TLR", align="C", fill=True)
            pdf.set_xy(x_save, y_save + HEADER_H / 2)
            pdf.cell(w, HEADER_H / 2, lines[1], border="BLR", align="C", fill=True)
            pdf.set_xy(x_save + w, y_save)
        else:
            pdf.cell(w, HEADER_H, lines[0], border=1, align="C", fill=True)
    pdf.ln(HEADER_H)

    # ── Filas de datos ─────────────────────────────────────────────────────────
    for i, fila in enumerate(data.filas):
        bg = ALT if i % 2 == 0 else WHITE
        pdf.set_fill_color(*bg)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 7.5)

        if pdf.get_y() + ROW_H > PAGE_H - 15:
            pdf.add_page()
            # Re-render header
            for label, w, _ in COLS:
                pdf.set_fill_color(*NAVY)
                pdf.set_text_color(*WHITE)
                pdf.set_font("Helvetica", "B", 7)
                lines = label.split("\n")
                if len(lines) == 2:
                    x_s = pdf.get_x(); y_s = pdf.get_y()
                    pdf.cell(w, HEADER_H/2, lines[0], border="TLR", align="C", fill=True)
                    pdf.set_xy(x_s, y_s + HEADER_H/2)
                    pdf.cell(w, HEADER_H/2, lines[1], border="BLR", align="C", fill=True)
                    pdf.set_xy(x_s + w, y_s)
                else:
                    pdf.cell(w, HEADER_H, lines[0], border=1, align="C", fill=True)
            pdf.ln(HEADER_H)

        vals = [
            _t(fila.fecha), _t(fila.turno), _t(fila.descripcion)[:38],
            _t(fila.horaIngreso), _t(fila.horaSalida),
            f"{fila.totalHoras:.1f}" if fila.totalHoras else "",
            f"{fila.hed:.1f}"  if fila.hed  else "",
            f"{fila.hen:.1f}"  if fila.hen  else "",
            f"{fila.hedf:.1f}" if fila.hedf else "",
            f"{fila.henf:.1f}" if fila.henf else "",
            _t(fila.centroCostos)[:20], _t(fila.actividad)[:35],
        ]
        for (_, w, align), val in zip(COLS, vals):
            pdf.cell(w, ROW_H, _t(val), border=1, align=align, fill=True)
        pdf.ln(ROW_H)

    # ── Pie ────────────────────────────────────────────────────────────────────
    pdf.ln(4)
    y_foot = pdf.get_y()
    pdf.set_xy(MARGIN, y_foot)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    pdf.set_fill_color(*WHITE)
    pdf.cell(60, 7, "FIRMA__________________________", border="B", align="L")

    # Total horas — lado derecho
    total_x = MARGIN + usable_w - 70
    pdf.set_xy(total_x, y_foot)
    pdf.set_fill_color(*HEADER_BG)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(45, 7, "TOTAL HORAS ADICIONALES", border=1, align="C", fill=True)
    pdf.set_fill_color(*YELLOW)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(25, 7, str(data.totalHorasAdicionales), border=1, align="C", fill=True)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf
