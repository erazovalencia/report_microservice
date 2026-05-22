from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
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
    ("Identificación",   "Cédula del empleado",                         "1069742877",  True,  16),
    ("Turno",            "Código de turno (ver hoja Turnos)",            "ADMI",        False, 14),
    ("Tipo Ausencia",    "Nombre de ausencia (ver hoja Ausencias)",     "Licencia No Remunerada", False, 22),
    ("Tipo de Bono",     "Código de bono (ver hoja Bonos)",              "BONO_CAMPO",  False, 16),
    ("Centro de Costo",  "Código centro de costo (ver hoja Costos)",    "",            False, 18),
    ("Pozo / Ubicacion", "Descripción de actividad del turno",           "",            False, 22),
    ("Hora Ingreso",     "Hora entrada real HH:MM (ej. 06:00)",         "06:00",       False, 14),
    ("Hora Salida",      "Hora salida real HH:MM  (ej. 18:00)",         "18:00",       False, 14),
    ("Notas",            "Observaciones del registro",                  "",            False, 28),
]


class RdpImportTemplateService(BaseExportService):

    def generate_file(
        self,
        data,
        options=None,
    ) -> io.BytesIO:
        shifts      = data.get("shifts",      [])
        absences    = data.get("absences",    [])
        bonuses     = data.get("bonuses",     [])
        workCenters = data.get("workCenters", [])
        employees   = data.get("employees",   [])

        wb = Workbook()
        self._build_main_sheet(wb, employees)
        self._build_catalog_sheet(wb, "Turnos",    shifts,      ["Código", "Nombre"],  display="code")
        self._build_catalog_sheet(wb, "Ausencias", absences,    ["Nombre", "Código"],  display="name")
        self._build_catalog_sheet(wb, "Bonos",     bonuses,     ["Código", "Nombre"],  display="code")
        self._build_catalog_sheet(wb, "Costos",    workCenters, ["Código", "Nombre"],  display="name")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def _build_main_sheet(self, wb: Workbook, employees: list):
        ws = wb.active
        ws.title = "Reporte"
        ws.freeze_panes = "A4"

        # Columna extra al final: Nombre (referencia, no importada)
        total_cols = len(COLUMNS) + 1
        ref_col    = len(COLUMNS) + 1

        # Title
        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
        t = ws["A1"]
        t.value = "PLANTILLA IMPORTACIÓN MASIVA RDP — Erazo Valencia"
        t.font = Font(bold=True, size=12, color=HEADER_FG)
        t.fill = PatternFill(start_color=HEADER_BG, fill_type="solid")
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # Subtitle / instructions
        ws.merge_cells(f"A2:{get_column_letter(total_cols)}2")
        s = ws["A2"]
        s.value = (
            "Instrucciones: complete una fila por empleado. "
            "Columnas con * son obligatorias. "
            "Ingrese Turno O Tipo Ausencia (no ambos). "
            "Use los códigos de las hojas de catálogo. "
            "No modifique las columnas A ni J."
        )
        s.font = Font(italic=True, size=9, color="555555")
        s.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 18

        # Column headers (row 3) — data columns
        for col, (label, hint, _, required, width) in enumerate(COLUMNS, start=1):
            header_label = f"{label} *" if required else label
            cell = ws.cell(row=3, column=col, value=header_label)
            cell.font = Font(bold=True, size=10, color=HEADER_FG)
            cell.fill = PatternFill(start_color=HEADER_BG, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.row_dimensions[3].height = 24

        # Reference column header (col I)
        ref_header = ws.cell(row=3, column=ref_col, value="Nombre (referencia)")
        ref_header.font = Font(bold=True, size=9, color="777777")
        ref_header.fill = PatternFill(start_color="E8E8E8", fill_type="solid")
        ref_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ref_header.border = BORDER
        ws.column_dimensions[get_column_letter(ref_col)].width = 30

        # Hint row (row 4)
        for col, (_, hint, _, required, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=4, column=col, value=hint)
            cell.font = Font(italic=True, size=8, color="666666")
            cell.fill = PatternFill(
                start_color=REQUIRED_BG if required else OPTIONAL_BG,
                fill_type="solid",
            )
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
        ws.cell(row=4, column=ref_col, value="Solo referencia — no se importa").font = Font(italic=True, size=8, color="999999")
        ws.row_dimensions[4].height = 20

        # Example row (row 5) — only shown if no employees
        if not employees:
            for col, (_, _, example, _, _) in enumerate(COLUMNS, start=1):
                cell = ws.cell(row=5, column=col, value=example)
                cell.font = Font(italic=True, size=9, color="1E6822")
                cell.fill = PatternFill(start_color=EXAMPLE_BG, fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BORDER
            ws.cell(row=5, column=ref_col, value="APELLIDO NOMBRE (ejemplo)").font = Font(italic=True, size=8, color="1E6822")
            ws.row_dimensions[5].height = 16
            start_data_row = 6
        else:
            start_data_row = 5

        # Pre-fill employee rows or blank rows
        rows_to_fill = employees if employees else []
        num_rows     = max(len(rows_to_fill), 50)

        for i in range(num_rows):
            r  = start_data_row + i
            bg = "F0F8FF" if i % 2 == 0 else "FFFFFF"

            emp    = rows_to_fill[i] if i < len(rows_to_fill) else None
            cedula = emp.get("cedula", "") if emp else ""
            nombre = emp.get("nombre", "") if emp else ""

            for col in range(1, total_cols + 1):
                cell = ws.cell(row=r, column=col)
                cell.border = BORDER

                if col == 1:
                    # Cédula pre-rellena — fondo destacado si tiene datos
                    cell.value = cedula
                    cell.font = Font(size=9, bold=bool(cedula))
                    cell.fill = PatternFill(
                        start_color="D9EAD3" if cedula else bg,
                        fill_type="solid",
                    )
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col == ref_col:
                    # Nombre referencia
                    cell.value = nombre
                    cell.font = Font(size=9, color="666666", italic=True)
                    cell.fill = PatternFill(start_color="F5F5F5", fill_type="solid")
                    cell.alignment = Alignment(vertical="center")
                else:
                    cell.value = ""
                    cell.fill = PatternFill(start_color=bg, fill_type="solid")

            ws.row_dimensions[r].height = 16

        # Data validation — dropdowns por columna
        self._add_validations(ws, start_data_row, start_data_row + num_rows - 1)

    def _add_validations(self, ws, first_row: int, last_row: int):
        # B — Turno
        dv_turno = DataValidation(
            type="list",
            formula1="Turnos!$A$2:$A$200",
            showDropDown=False,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Código inválido",
            error="Seleccione un turno de la hoja 'Turnos'",
        )
        dv_turno.sqref = f"B{first_row}:B{last_row}"
        ws.add_data_validation(dv_turno)

        # C — Tipo Ausencia
        dv_ausencia = DataValidation(
            type="list",
            formula1="Ausencias!$A$2:$A$200",
            showDropDown=False,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Código inválido",
            error="Seleccione un código de la hoja 'Ausencias'",
        )
        dv_ausencia.sqref = f"C{first_row}:C{last_row}"
        ws.add_data_validation(dv_ausencia)

        # D — Tipo de Bono
        dv_bono = DataValidation(
            type="list",
            formula1="Bonos!$A$2:$A$200",
            showDropDown=False,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Código inválido",
            error="Seleccione un bono de la hoja 'Bonos'",
        )
        dv_bono.sqref = f"D{first_row}:D{last_row}"
        ws.add_data_validation(dv_bono)

        # E — Centro de Costo
        dv_costo = DataValidation(
            type="list",
            formula1="Costos!$A$2:$A$200",
            showDropDown=False,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Código inválido",
            error="Seleccione un centro de costo de la hoja 'Costos'",
        )
        dv_costo.sqref = f"E{first_row}:E{last_row}"
        ws.add_data_validation(dv_costo)

    def _build_catalog_sheet(
        self,
        wb: Workbook,
        title: str,
        entries: list,
        col_headers: list,
        display: str = "code",  # "code" | "code_name" | "name"
    ):
        ws = wb.create_sheet(title)

        # Proteger la hoja para que el reportador no la modifique accidentalmente
        ws.protection.sheet = True
        ws.protection.selectLockedCells = False    # puede seleccionar para ver
        ws.protection.selectUnlockedCells = False

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

        max_a = 10
        max_b = 10
        for i, entry in enumerate(entries, start=2):
            code = entry.get("code", "") if isinstance(entry, dict) else getattr(entry, "code", "")
            name = entry.get("name", "") if isinstance(entry, dict) else getattr(entry, "name", "")
            bg   = "F5F5F5" if i % 2 == 0 else "FFFFFF"
            fill = PatternFill(start_color=bg, fill_type="solid")

            if display == "code_name":
                # Col A = "código — nombre"  (el que va al dropdown)
                # Col B = código limpio       (referencia visual)
                display_val = f"{code} — {name}" if name else code
                c1 = ws.cell(row=i, column=1, value=display_val)
                c2 = ws.cell(row=i, column=2, value=code)
                max_a = max(max_a, len(display_val))
                max_b = max(max_b, len(code))
            elif display == "name":
                # Col A = nombre (al dropdown)
                # Col B = código (referencia)
                c1 = ws.cell(row=i, column=1, value=name)
                c2 = ws.cell(row=i, column=2, value=code)
                max_a = max(max_a, len(name))
                max_b = max(max_b, len(code))
            else:
                # default: col A = código, col B = nombre
                c1 = ws.cell(row=i, column=1, value=code)
                c2 = ws.cell(row=i, column=2, value=name)
                max_a = max(max_a, len(code))
                max_b = max(max_b, len(name))

            c1.fill = fill; c1.border = BORDER; c1.font = Font(bold=(display == "code"), size=9)
            c2.fill = fill; c2.border = BORDER; c2.font = Font(size=9, color="888888")

        ws.column_dimensions["A"].width = min(max_a + 4, 55)
        ws.column_dimensions["B"].width = min(max_b + 4, 20)

    def get_content_type(self) -> str:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def get_file_extension(self) -> str:
        return ".xlsx"
