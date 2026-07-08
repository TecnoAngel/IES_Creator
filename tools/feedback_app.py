from __future__ import annotations

import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

import streamlit as st
import win32com.client  # type: ignore
import xlwings as xw


def safe_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    return name or "Alumno"


def _read_source_bytes(source: Any) -> bytes:
    if isinstance(source, (str, Path)):
        path = Path(source).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"No existe el archivo: {path}")
        return path.read_bytes()
    if hasattr(source, "getvalue"):
        return source.getvalue()
    raise TypeError("El origen debe ser una ruta o un archivo subido")


def run_feedback(source: Any, output_dir: Path | str) -> None:
    output_dir = Path(output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    payload = _read_source_bytes(source)
    with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(payload)
        temp_path = Path(tmp.name)

    app = xw.App(visible=False, add_book=False)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        wb = app.books.open(str(temp_path))
        try:
            sh = wb.sheets.active
            data = sh.used_range.value
            if not data or len(data) < 2 or len(data[0]) < 3:
                raise ValueError("La tabla debe tener encabezados y al menos 3 columnas.")

            headers = data[0]
            activity_label = str(headers[0]).strip() if headers[0] is not None else ""
            item_headers = [("" if h is None else str(h).strip()) for h in headers[2:]]

            tmp_wb = app.books.add()
            tmp_sh = tmp_wb.sheets[0]
            tmp_sh.range("A:A").column_width = 70
            tmp_sh.range("B:B").column_width = 12
            tmp_sh.range("A:B").api.VerticalAlignment = -4160

            for r in data[1:]:
                if not r or r[0] is None or str(r[0]).strip() == "":
                    continue

                student = str(r[0]).strip()
                grade = r[1]
                values = (r[2:] + [None] * len(item_headers))[:len(item_headers)]

                tmp_sh.clear()
                tmp_sh.range("A1").value = activity_label
                tmp_sh.range("B1").value = grade

                for i, (h, v) in enumerate(zip(item_headers, values), start=2):
                    tmp_sh.range(f"A{i}").value = h
                    tmp_sh.range(f"B{i}").value = v

                last_row = 1 + len(item_headers)
                rng = tmp_sh.range(f"A1:B{last_row}")
                rng.api.Borders.Weight = 2
                tmp_sh.range("A1:B1").api.Font.Bold = True

                for i in range(2, last_row + 1):
                    v = tmp_sh.range(f"B{i}").value
                    cell = tmp_sh.range(f"B{i}").api
                    if isinstance(v, (int, float)) and float(v) == 0.0:
                        cell.Interior.Color = 0xC0C0FF
                    else:
                        cell.Interior.Color = 0xC0FFC0

                doc = word.Documents.Add()
                p = doc.Paragraphs.Add()
                p.Range.Text = student
                p.Range.Font.Bold = True
                p.Range.InsertParagraphAfter()
                rng.api.Copy()
                doc.Paragraphs.Add().Range.PasteExcelTable(False, False, True)

                docx_path = output_dir / f"{safe_filename(student)}.docx"
                doc.SaveAs2(str(docx_path))
                doc.Close(False)

            tmp_wb.close()
            st.success(f"Se han generado los documentos en {output_dir}")
        finally:
            wb.close()
    finally:
        app.quit()
        word.Quit()
        temp_path.unlink(missing_ok=True)
