from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from typing import List

from ...base import BaseExportService
from ....models.RdpModel import RdpCatalogEntry

thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

HEADER_BG   = "1F3864"
HEADER_FG   = "FFFFFF"
EXAMPLE_BG  = "EAF4EA"
REQUIRED_BG = "FFF3CD"
OPTIONAL_BG = "F0F0F0"

COLUMNS = [
    ("Identificación",   "Cédula del empleado",                     "1069742877",  True,  16),
    ("Es Ausencia",      "Si = ausencia  /  No = turno normal",     "No",          True,  14),
    ("Turno",            "Código de turno (ver hoja Turnos)",        "TUR1",        False, 14),
    ("Tipo Ausencia",    "Código de ausencia (ver hoja Ausencias)", "300",         False, 16),
    ("Tipo de Bono",     "Código de bono (ver hoja Bonos)",          "BONO_CAMPO",  False, 16),
    ("Proyecto",         "Código de proyecto (ver hoja Proyectos)", "",            False, 16),
    ("Hora Extra Total", "Horas extras totales (número decimal)",   "2.5",         False, 18),
    ("Notas",            "Observaciones del registro",              "",            False, 28),
]


class RdpImportTemplateService(BaseExportService):

    def generate_file(
        self,
        data,
        options=None,
    ) -> io.BytesIO:
        shifts   = data.get("shifts",   [])
        absences = data.get("absences", [])
        bonuses  = data.get("bonuses",  [])
        projects = data.get("projects", [])

        wb = Workbook()
        self._build_main_sheet(wb)
        self._build_catalog_sheet(wb, "Turnos",    shifts,   ["Código", "Nombre"])
        self._build_catalog_sheet(wb, "Ausencias", absences, ["Código", "Nombre"])
        self._build_catalog_sheet(wb, "Bonos",     bonuses,  ["Código", "Nombre"])
        self._build_catalog_sheet(wb, "Proyectos", projects, ["Código", "Nombre"])

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def _build_main_sheet(self, wb: Workbook):
        ws = wb.active
        ws.title = "Reporte"
        ws.freeze_panes = "A4"

        # Title
        ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
        t = ws["A1"]
        t.value = "PLANTILLA IMPORTACIÓN MASIVA RDP — Erazo Valencia"
        t.font = Font(bold=True, size=12, color=HEADER_FG)
        t.fill = PatternFill(start_color=HEADER_BG, fill_type="solid")
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # Subtitle / instructions
        ws.merge_cells(f"A2:{get_column_letter(len(COLUMNS))}2")
        s = ws["A2"]
        s.value = (
            "Instrucciones: complete una fila por empleado. "
            "Columnas con * son obligatorias. "
            "Use los códigos de las hojas de catálogo. "
            "No modifique los encabezados."
        )
        s.font = Font(italic=True, size=9, color="555555")
        s.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 18

        # Column headers (row 3)
        for col, (label, hint, _, required, width) in enumerate(COLUMNS, start=1):
            header_label = f"{label} *" if required else label
            cell = ws.cell(row=3, column=col, value=header_label)
            cell.font = Font(bold=True, size=10, color=HEADER_FG)
            cell.fill = PatternFill(start_color=HEADER_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.row_dimensions[3].height = 24

        # Hint row (row 4) — light color hint per column
        for col, (_, hint, _, required, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=4, column=col, value=hint)
            cell.font = Font(italic=True, size=8, color="666666")
            cell.fill = PatternFill(
                start_color=REQUIRED_BG if required else OPTIONAL_BG,
                fill_type="solid",
            )
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
        ws.row_dimensions[4].height = 20

        # Example row (row 5)
        for col, (_, _, example, _, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=5, column=col, value=example)
            cell.font = Font(italic=True, size=9, color="1E6822")
            cell.fill = PatternFill(start_color=EXAMPLE_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
        ws.row_dimensions[5].height = 16

        # Leave rows 6+ blank for data entry
        for r in range(6, 106):
            for col in range(1, len(COLUMNS) + 1):
                cell = ws.cell(row=r, column=col, value="")
                cell.border = BORDER
                cell.fill = PatternFill(
                    start_color="FAFAFA" if r % 2 == 0 else "FFFFFF",
                    fill_type="solid",
                )
            ws.row_dimensions[r].height = 16

    def _build_catalog_sheet(
        self,
        wb: Workbook,
        title: str,
        entries: list,
        col_headers: list,
    ):
        ws = wb.create_sheet(title)

        for col, header in enumerate(col_headers, start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color=HEADER_FG)
            cell.fill = PatternFill(start_color=HEADER_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER

        if not entries:
            ws.cell(row=2, column=1, value="(sin registros)")
            ws.column_dimensions["A"].width = 20
            ws.column_dimensions["B"].width = 40
            return

        max_name = 10
        for i, entry in enumerate(entries, start=2):
            code = entry.get("code", "") if isinstance(entry, dict) else getattr(entry, "code", "")
            name = entry.get("name", "") if isinstance(entry, dict) else getattr(entry, "name", "")
            bg = "F5F5F5" if i % 2 == 0 else "FFFFFF"
            fill = PatternFill(start_color=bg, fill_type="solid")

            c1 = ws.cell(row=i, column=1, value=code)
            c1.fill = fill
            c1.border = BORDER
            c1.font = Font(bold=True, size=9)

            c2 = ws.cell(row=i, column=2, value=name)
            c2.fill = fill
            c2.border = BORDER
            c2.font = Font(size=9)

            max_name = max(max_name, len(name))

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = min(max_name + 4, 50)

    def get_content_type(self) -> str:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def get_file_extension(self) -> str:
        return ".xlsx"
