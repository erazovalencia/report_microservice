import io
import os
import unicodedata

from fpdf import FPDF
from fpdf.fonts import FontFace

from ....models.AssetModel import AssetAssignmentNoticeRequest

# Formato de asignación de activos a un empleado (acta de entrega), basado en
# el formato "ASIGNACIÓN EQUIPOS" de IT: información del empleado, equipos
# asignados con su condición de entrega y firmas de entrega/recibido.

LOGO_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "..", "common", "logo.png")

MARGIN = 15
CONTENT_W = 210 - 2 * MARGIN

BAR_GRAY = (89, 89, 89)
LABEL_GRAY = (110, 110, 110)
LINE_GRAY = (153, 153, 153)
TEXT_DARK = (30, 30, 30)
HEAD_FILL = (235, 235, 235)


def _t(text) -> str:
    """Texto seguro para Helvetica (Latin-1): conserva tildes y ñ, degrada el resto a ASCII."""
    s = unicodedata.normalize("NFC", str(text) if text is not None else "")
    out = []
    for ch in s:
        try:
            ch.encode("latin-1")
            out.append(ch)
        except UnicodeEncodeError:
            out.append(unicodedata.normalize("NFKD", ch).encode("ascii", "ignore").decode("ascii"))
    return "".join(out)


class AssignmentPDF(FPDF):
    def __init__(self, footer_text: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self._footer_text = footer_text
        self.set_margins(MARGIN, MARGIN, MARGIN)
        self.set_auto_page_break(auto=True, margin=18)

    def footer(self):
        self.set_y(-12)
        self.set_draw_color(*LINE_GRAY)
        self.line(MARGIN, self.get_y(), 210 - MARGIN, self.get_y())
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*LABEL_GRAY)
        self.cell(CONTENT_W / 2, 6, _t(self._footer_text), align="L")
        self.cell(CONTENT_W / 2, 6, f"Página {self.page_no()}", align="R")


def _section_bar(pdf: FPDF, title: str):
    pdf.ln(4)
    pdf.set_fill_color(*BAR_GRAY)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(CONTENT_W, 7, _t(f"  {title}"), fill=True, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)


def _field(pdf: FPDF, x: float, width: float, label: str, value):
    """Etiqueta pequeña arriba y valor subrayado, como las casillas del formato."""
    y = pdf.get_y()
    pdf.set_xy(x, y)
    pdf.set_font("Helvetica", "", 7)
    pdf.set_text_color(*LABEL_GRAY)
    pdf.cell(width, 4, _t(label))
    pdf.set_xy(x, y + 4)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(*TEXT_DARK)
    pdf.cell(width - 3, 6, _t(value or "No registrado"))
    pdf.set_draw_color(*LINE_GRAY)
    pdf.line(x, y + 10.5, x + width - 3, y + 10.5)


def _field_row(pdf: FPDF, fields):
    """fields: lista de (label, value, fracción del ancho)."""
    y = pdf.get_y()
    x = MARGIN
    for label, value, frac in fields:
        pdf.set_y(y)
        width = CONTENT_W * frac
        _field(pdf, x, width, label, value)
        x += width
    pdf.set_xy(MARGIN, y + 13)


def _signature(pdf: FPDF, x: float, width: float, title: str, name, detail=None):
    y = pdf.get_y()
    pdf.set_draw_color(*TEXT_DARK)
    pdf.line(x + 4, y + 18, x + width - 4, y + 18)
    pdf.set_xy(x, y + 19)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(*TEXT_DARK)
    pdf.cell(width, 5, _t(title), align="C")
    pdf.set_xy(x, y + 24)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.cell(width, 5, _t(name or ""), align="C")
    if detail:
        pdf.set_xy(x, y + 29)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(*LABEL_GRAY)
        pdf.cell(width, 4, _t(detail), align="C")


def generate_asset_assignment_pdf(data: AssetAssignmentNoticeRequest) -> io.BytesIO:
    pdf = AssignmentPDF(f"ERAZO VALENCIA · CONTROL DE ACTIVOS {data.lineLabel}")
    pdf.add_page()

    # Encabezado: logo + título
    top = pdf.get_y()
    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=MARGIN, y=top, h=16)
    pdf.set_xy(MARGIN + 60, top + 2)
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(*TEXT_DARK)
    pdf.cell(CONTENT_W - 60, 8, _t(f"ASIGNACIÓN DE EQUIPOS {data.lineLabel}"), align="R")
    pdf.set_xy(MARGIN + 60, top + 10)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(*LABEL_GRAY)
    pdf.cell(CONTENT_W - 60, 5, _t(f"Fecha de entrega: {data.assignedAt}"), align="R")
    pdf.set_xy(MARGIN, top + 20)

    _section_bar(pdf, "INFORMACIÓN DEL EMPLEADO")
    _field_row(pdf, [("Nombre del empleado", data.employeeName, 0.5), ("Cédula", data.employeeDocumentId, 0.5)])
    _field_row(pdf, [("Puesto", data.position, 0.5), ("Departamento", data.department, 0.5)])
    _field_row(pdf, [("Superior que autoriza", data.authorizedBy, 0.5), ("Empresa", data.company, 0.5)])

    _section_bar(pdf, "EQUIPOS ASIGNADOS")
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*TEXT_DARK)
    pdf.set_draw_color(*LINE_GRAY)
    # La barra de sección deja el relleno en gris oscuro; las filas van sin fondo.
    pdf.set_fill_color(255, 255, 255)
    headings = FontFace(emphasis="BOLD", fill_color=HEAD_FILL)
    with pdf.table(
        col_widths=(26, 20, 52, 24, 30, 28),
        text_align=("LEFT", "LEFT", "LEFT", "LEFT", "LEFT", "LEFT"),
        headings_style=headings,
        line_height=5,
        width=CONTENT_W,
    ) as table:
        head = table.row()
        for label in ("Nº inventario", "Equipo SAP", "Descripción", "Marca", "Modelo", "Serie"):
            head.cell(_t(label))
        for asset in data.assets:
            row = table.row()
            for value in (
                asset.assetTag,
                asset.sapEquipmentCode,
                asset.description,
                asset.brand,
                asset.model,
                asset.serialNumber,
            ):
                row.cell(_t(value or "-"))

    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(CONTENT_W, 5, _t("Condición de entrega"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 8.5)
    pdf.multi_cell(CONTENT_W, 4.5, _t(data.condition or "Sin observaciones registradas al momento de la entrega."), align="L")

    pdf.ln(3)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(*LABEL_GRAY)
    pdf.multi_cell(
        CONTENT_W,
        4.2,
        _t(
            "Con la firma de este documento, el empleado confirma haber recibido los equipos descritos "
            "en las condiciones indicadas."
        ),
        align="L",
    )

    _section_bar(pdf, "FIRMAS")
    if pdf.get_y() > 250:
        pdf.add_page()
    pdf.ln(4)
    half = CONTENT_W / 2
    y = pdf.get_y()
    delivered_detail = f"C.C. {data.deliveredByDocumentId}" if data.deliveredByDocumentId else None
    _signature(pdf, MARGIN, half, "Entrega", data.deliveredBy, delivered_detail)
    pdf.set_y(y)
    _signature(pdf, MARGIN + half, half, "Recibe", data.employeeName, f"C.C. {data.employeeDocumentId}")

    return io.BytesIO(bytes(pdf.output()))
