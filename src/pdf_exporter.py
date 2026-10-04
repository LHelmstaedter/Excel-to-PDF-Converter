from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

import pythoncom
import win32com.client  # type: ignore

from .utils import sanitize_filename, build_output_path


def _col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", letter):
        raise ValueError(f"Ungültige Spalte: {letter}")
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _autodetect_filter_sheet_and_range(wb) -> Tuple[str, str]:
    """Sheet with a print area (else the active sheet) and its used range."""
    candidates = []
    for ws in wb.Worksheets:
        try:
            pa = ws.PageSetup.PrintArea
        except Exception:
            pa = ""
        if pa:
            candidates.append(ws)
    ws = candidates[0] if candidates else wb.ActiveSheet
    used = ws.UsedRange
    end_row = used.Row + used.Rows.Count - 1
    end_col = used.Column + used.Columns.Count - 1
    return ws.Name, f"{_col_letter(used.Column)}{used.Row}:{_col_letter(end_col)}{end_row}"


def _read_unique_groups(ws, group_col_letter: str, header_row: int = 1) -> List[str]:
    col_idx = _col_letter_to_index(group_col_letter)
    last_row = ws.Cells(ws.Rows.Count, col_idx).End(-4162).Row  # xlUp
    first_data_row = header_row + 1
    if last_row < first_data_row:
        return []
    values = ws.Range(ws.Cells(first_data_row, col_idx), ws.Cells(last_row, col_idx)).Value
    if values is None:
        return []
    if isinstance(values, (str, int, float)):
        values = ((values,),)
    groups = [str(row[0]).strip() for row in values if row[0] is not None]
    return list(dict.fromkeys(g for g in groups if g))  # unique, keeps order


def export_group_pdfs(
    excel_path: str,
    out_dir: str,
    sheet_name: Optional[str] = None,
    filter_range: Optional[str] = None,
    group_col_letter: str = "N",
) -> None:
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AutomationSecurity = 3  # disable macros
        except Exception:
            pass

        wb = excel.Workbooks.Open(excel_path, ReadOnly=True)
        if sheet_name is None or filter_range is None:
            auto_sheet, auto_range = _autodetect_filter_sheet_and_range(wb)
            sheet_name = sheet_name or auto_sheet
            filter_range = filter_range or auto_range

        ws = wb.Worksheets(sheet_name)
        excel.CalculateFull()

        groups = _read_unique_groups(ws, group_col_letter)
        if not groups:
            raise RuntimeError(f"Keine Gruppen gefunden in Spalte {group_col_letter}.")

        Path(out_dir).mkdir(parents=True, exist_ok=True)

        rng = ws.Range(filter_range)
        field = _col_letter_to_index(group_col_letter) - rng.Column + 1
        if field < 1 or field > rng.Columns.Count:
            raise RuntimeError(f"Spalte {group_col_letter} nicht in AutoFilter-Range {filter_range}.")

        for g in groups:
            rng.AutoFilter(Field=field, Criteria1=g)
            excel.CalculateFull()

            pdf_file = build_output_path(out_dir, g, sanitize_filename(g) + ".pdf")
            Path(pdf_file).parent.mkdir(parents=True, exist_ok=True)

            ws.ExportAsFixedFormat(
                Type=0,
                Filename=str(pdf_file),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
