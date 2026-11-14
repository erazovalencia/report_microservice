import io

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from ...base import BaseExportService
from ....models.TreasuryModel import TreasuryTransactionsExportRequest
from ._cells import AMOUNT_FORMAT, BOLD, DATE_TIME_FORMAT, NOTE_FONT, cell, header_row, iso_datetime, text, title_rows

# One row per treasury payment, as the transactions table shows it, plus a
# second sheet with the totals per currency. A wide date range brings the whole
# AuraQuantic history (tens of thousands of rows), so the workbook is written in
# write-only mode: rows stream to the file instead of living in memory, and only
# the cells that need a style or a type are built as cells.

COLUMNS = [
    ("Referencia", 15),
    ("Estado", 13),
    ("Registrado", 17),
    ("Decidido", 17),
    ("Empresa que paga", 30),
    ("Banco", 22),
    ("Cuenta", 18),
    ("Mecanismo", 16),
    ("Lote", 12),
    ("QR", 6),
    ("NIT tercero", 16),
    ("Tercero", 34),
    ("Descripción", 48),
    ("Valor", 18),
    ("Moneda", 9),
    ("Registró", 26),
]
HEADER_ROW = 4


class TreasuryTransactionsExportService(BaseExportService):

    def generate_file(self, data: TreasuryTransactionsExportRequest, options=None) -> io.BytesIO:
        wb = Workbook(write_only=True)
        self._payments_sheet(wb, data)
        self._totals_sheet(wb, data)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def _payments_sheet(self, wb: Workbook, data: TreasuryTransactionsExportRequest):
        ws = wb.create_sheet("Pagos")
        # Write-only sheets emit columns and panes with the first row: both go before any append.
        for index, (_, width) in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = f"A{HEADER_ROW + 1}"

        by = f" por {data.generatedBy}" if data.generatedBy else ""
        title_rows(ws, "PAGOS DE TESORERÍA", data.filtersSummary, f"Generado el {data.generatedAt}{by} · {len(data.rows)} pagos", len(COLUMNS))
        header_row(ws, [label for label, _ in COLUMNS])

        for row in data.rows:
            ws.append([
                row.reference,
                row.status,
                cell(ws, iso_datetime(row.registeredAt), number_format=DATE_TIME_FORMAT),
                cell(ws, iso_datetime(row.decidedAt), number_format=DATE_TIME_FORMAT),
                text(ws, row.companyName),
                text(ws, row.bankName),
                text(ws, row.accountNumber),
                row.paymentMethod,
                text(ws, row.batchReference),
                "Sí" if row.isQr else "",
                text(ws, row.payeeTaxId),
                text(ws, row.payeeName),
                text(ws, row.description),
                cell(ws, row.amount, number_format=AMOUNT_FORMAT),
                text(ws, row.currencyLabel),
                text(ws, row.treasurerName),
            ])

        last_row = HEADER_ROW + len(data.rows)
        ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(COLUMNS))}{last_row}"
        ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
        ws.page_setup.orientation = "landscape"

    def _totals_sheet(self, wb: Workbook, data: TreasuryTransactionsExportRequest):
        ws = wb.create_sheet("Totales")
        for index, width in enumerate((12, 10, 22), start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        header_row(ws, ["Moneda", "Pagos", "Total"])
        for total in data.totals:
            ws.append([text(ws, total.currencyLabel), total.count, cell(ws, total.total, number_format=AMOUNT_FORMAT, font=BOLD)])
        ws.append([])
        ws.append([cell(ws, "Cada total suma en la moneda de sus cuentas; no se convierte con tasas de cambio.", font=NOTE_FONT)])

    def get_content_type(self) -> str:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def get_file_extension(self) -> str:
        return ".xlsx"
