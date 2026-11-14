import io

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from ...base import BaseExportService
from ....models.TreasuryModel import TreasuryBalancesExportRequest
from ._cells import AMOUNT_FORMAT, DATE_FORMAT, cell, header_row, iso_date, text, title_rows

# History of sent bank balances: one row per account and cut, as it was sent,
# with the peso equivalent at the rate of each cut. Written in write-only mode
# like the transactions: a long range of two daily cuts over sixty accounts
# adds up to tens of thousands of rows.

COLUMNS = [
    ("Fecha", 12),
    ("Franja", 8),
    ("Sección", 30),
    ("Empresa", 30),
    ("Banco", 24),
    ("Cuenta", 18),
    ("Moneda", 9),
    ("Saldo en banco", 20),
    ("Saldo en libros", 20),
    ("En pesos", 20),
]
HEADER_ROW = 4


class TreasuryBalancesExportService(BaseExportService):

    def generate_file(self, data: TreasuryBalancesExportRequest, options=None) -> io.BytesIO:
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Saldos")
        for index, (_, width) in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = f"A{HEADER_ROW + 1}"

        by = f" por {data.generatedBy}" if data.generatedBy else ""
        title_rows(ws, "SALDOS DE BANCOS", data.filtersSummary, f"Generado el {data.generatedAt}{by} · {len(data.rows)} saldos", len(COLUMNS))
        header_row(ws, [label for label, _ in COLUMNS])

        for row in data.rows:
            ws.append([
                cell(ws, iso_date(row.cutDate), number_format=DATE_FORMAT),
                row.slot,
                text(ws, row.balanceGroupName),
                text(ws, row.companyName),
                text(ws, row.bankName),
                text(ws, row.accountNumber),
                text(ws, row.currencyLabel),
                cell(ws, row.bankBalance, number_format=AMOUNT_FORMAT),
                cell(ws, row.bookBalance, number_format=AMOUNT_FORMAT),
                cell(ws, row.copEquivalent, number_format=AMOUNT_FORMAT),
            ])

        ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(COLUMNS))}{HEADER_ROW + len(data.rows)}"
        ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
        ws.page_setup.orientation = "landscape"
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def get_content_type(self) -> str:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def get_file_extension(self) -> str:
        return ".xlsx"
