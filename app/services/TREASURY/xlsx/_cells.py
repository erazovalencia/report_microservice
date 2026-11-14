from datetime import date, datetime
from typing import Optional

from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill

# Cells shared by the treasury workbooks, all written in write-only mode. Text
# typed by users never becomes a formula and never carries control characters.

NAVY = "1F3864"
WHITE = "FFFFFF"
GRAY = "5B6779"

TITLE_FONT = Font(bold=True, size=14, color=WHITE)
HEADER_FONT = Font(bold=True, size=10, color=WHITE)
NOTE_FONT = Font(size=9, italic=True, color=GRAY)
BOLD = Font(bold=True)
NAVY_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

DATE_FORMAT = "yyyy-mm-dd"
DATE_TIME_FORMAT = "yyyy-mm-dd hh:mm"
AMOUNT_FORMAT = "#,##0.00"


def is_formula_like(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def clean(value):
    """
    Control characters pasted from other systems are not valid in the sheet XML
    and openpyxl refuses the whole workbook over a single one: they are dropped.
    """
    return ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value


def cell(ws, value, font=None, fill=None, number_format=None, alignment=None) -> WriteOnlyCell:
    """A styled cell. These exports never write formulas, so text is always stored as text."""
    value = clean(value)
    result = WriteOnlyCell(ws, value=value)
    if is_formula_like(value):
        result.data_type = "s"
    if font:
        result.font = font
    if fill:
        result.fill = fill
    if number_format:
        result.number_format = number_format
    if alignment:
        result.alignment = alignment
    return result


def text(ws, value):
    """
    openpyxl stores any string that starts with "=" as a formula, and names and
    descriptions are typed by users: such a value goes as a text cell so opening
    the file never evaluates it. Any other value stays a bare value, which
    write-only mode serializes without building a cell object.
    """
    value = clean(value)
    return cell(ws, value) if is_formula_like(value) else value


def iso_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def iso_date(value: Optional[str]) -> Optional[date]:
    parsed = iso_datetime(value)
    return parsed.date() if parsed else None


def title_rows(ws, title: str, summary: str, generated: str, width: int) -> None:
    """Title band, filters line and generation line, the three rows above every header."""
    ws.append([cell(ws, title, font=TITLE_FONT, fill=NAVY_FILL)] + [cell(ws, None, fill=NAVY_FILL) for _ in range(width - 1)])
    ws.append([cell(ws, summary or "Sin filtros", font=NOTE_FONT)])
    ws.append([cell(ws, generated, font=NOTE_FONT)])


def header_row(ws, labels) -> None:
    ws.append([cell(ws, label, font=HEADER_FONT, fill=NAVY_FILL, alignment=CENTER) for label in labels])
