"""Post-processing helpers for the generated Excel workbooks."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional, Sequence, Set

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def _is_row_empty(values: Iterable[object]) -> bool:
    return all((cell is None or (isinstance(cell, str) and not cell.strip())) for cell in values)

def tidy_workbook(
    path: Path,
    sheet_name: str = "EOD",
    *,
    headers: Optional[Sequence[str]] = None,
    column_colours: Optional[Mapping[str, str]] = None,
    variance_headers: Optional[Set[str]] = None,
    decimal_headers: Optional[Set[str]] = None,
) -> None:
    """Remove trailing empty rows and apply consistent formatting.

    ``load_workbook`` is tolerant of missing files, so we guard for that
    explicitly and simply return if the workbook is not present yet.  The
    function mutates the workbook in-place and re-saves it, ensuring the
    heading row is bold, columns are auto-sized, and key categories are
    colour coded.
    """

    workbook_path = Path(path)
    if not workbook_path.exists():
        return

    wb = load_workbook(workbook_path)
    if sheet_name not in wb.sheetnames:
        wb.save(workbook_path)
        return

    ws = wb[sheet_name]
    last_row: Optional[int] = ws.max_row

    while last_row and last_row > 1:
        if not _is_row_empty(cell.value for cell in ws[last_row]):
            break
        ws.delete_rows(last_row)
        last_row -= 1

    header_values = list(headers or [])
    if header_values:
        for idx, header in enumerate(header_values, start=1):
            ws.cell(row=1, column=idx, value=header)

    ws.freeze_panes = "A2"

    header_font = Font(bold=True)
    variance_headers = set(variance_headers or set())
    decimal_headers = set(decimal_headers or set())
    column_colours = dict(column_colours or {})

    header_fill_cache: Dict[str, PatternFill] = {}

    for idx, cell in enumerate(ws[1], start=1):
        header_text = cell.value if cell.value is not None else header_values[idx - 1] if idx <= len(header_values) else None
        if isinstance(header_text, str):
            header_text = header_text.strip()
        fill_colour = column_colours.get(header_text) if header_text else None
        if fill_colour:
            if fill_colour not in header_fill_cache:
                header_fill_cache[fill_colour] = PatternFill(start_color=fill_colour, end_color=fill_colour, fill_type="solid")
            cell.fill = header_fill_cache[fill_colour]
        cell.font = header_font

    neg_font = Font(color="FFC00000")
    pos_font = Font(color="FF006400")
    neutral_font = Font(color="FF000000")

    max_lengths: Dict[int, int] = {}
    max_column = ws.max_column
    for row in ws.iter_rows(min_row=2, max_col=max_column):
        for cell in row:
            header_cell = ws.cell(row=1, column=cell.column)
            header_text = header_cell.value if isinstance(header_cell.value, str) else None
            if header_text and header_text in column_colours:
                fill_colour = column_colours[header_text]
                cell.fill = header_fill_cache.setdefault(
                    fill_colour,
                    PatternFill(start_color=fill_colour, end_color=fill_colour, fill_type="solid"),
                )

            if header_text in decimal_headers and cell.value not in (None, ""):
                try:
                    cell.value = float(cell.value)
                except (TypeError, ValueError):
                    pass
                else:
                    cell.number_format = "#,##0.00"

            if header_text in variance_headers and isinstance(cell.value, (int, float)):
                if cell.value < 0:
                    cell.font = neg_font
                elif cell.value > 0:
                    cell.font = pos_font
                else:
                    cell.font = neutral_font

            length = len(str(cell.value)) if cell.value not in (None, "") else 0
            max_lengths[cell.column] = max(max_lengths.get(cell.column, 0), length)

    for column_idx in range(1, max_column + 1):
        column_letter = get_column_letter(column_idx)
        width = max_lengths.get(column_idx, len(str(ws.cell(row=1, column=column_idx).value or ""))) + 2
        ws.column_dimensions[column_letter].width = max(10, min(width, 40))

    wb.save(workbook_path)


__all__ = ["tidy_workbook"]
