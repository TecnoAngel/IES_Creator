from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from openpyxl import load_workbook


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


def _is_zero(value: Any) -> bool:
    try:
        return float(value) == 0.0
    except (TypeError, ValueError):
        return False


def _set_cell_background(cell, color_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _read_rows(payload: bytes) -> list[tuple]:
    wb = load_workbook(BytesIO(payload), data_only=True)
    sh = wb.active
    return list(sh.iter_rows(values_only=True))


def build_student_docx(student: str, activity_label: str, grade: Any, item_headers: list[str], values: list[Any]) -> bytes:
    doc = Document()

    p = doc.add_paragraph()
    run = p.add_run(student)
    run.bold = True

    last_row = 1 + len(item_headers)
    table = doc.add_table(rows=last_row, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    header_row = table.rows[0]
    header_row.cells[0].text = activity_label
    header_row.cells[1].text = "" if grade is None else str(grade)
    for cell in header_row.cells:
        for run in cell.paragraphs[0].runs:
            run.bold = True

    for i, (label, value) in enumerate(zip(item_headers, values), start=1):
        row = table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = "" if value is None else str(value)
        color = "FFC0C0" if _is_zero(value) else "C0FFC0"
        _set_cell_background(row.cells[1], color)

    for row in table.rows:
        row.cells[0].width = Cm(12)
        row.cells[1].width = Cm(2.5)

    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def run_feedback(source: Any) -> bytes:
    """Genera un ZIP en memoria con un .docx de feedback por alumno."""
    import zipfile

    payload = _read_source_bytes(source)
    data = _read_rows(payload)

    if not data or len(data) < 2 or len(data[0]) < 3:
        raise ValueError("La tabla debe tener encabezados y al menos 3 columnas.")

    headers = data[0]
    activity_label = str(headers[0]).strip() if headers[0] is not None else ""
    item_headers = [("" if h is None else str(h).strip()) for h in headers[2:]]

    zip_buffer = BytesIO()
    generated = 0
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in data[1:]:
            if not r or r[0] is None or str(r[0]).strip() == "":
                continue

            student = str(r[0]).strip()
            grade = r[1]
            values = (list(r[2:]) + [None] * len(item_headers))[: len(item_headers)]

            docx_bytes = build_student_docx(student, activity_label, grade, item_headers, values)
            zf.writestr(f"{safe_filename(student)}.docx", docx_bytes)
            generated += 1

    if generated == 0:
        raise ValueError("No se ha encontrado ningún alumno con datos en el Excel.")

    return zip_buffer.getvalue()
