from __future__ import annotations

import random
import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

import qrcode
import streamlit as st
import xlwings as xw
from docx import Document
from docx.shared import Cm
from docxtpl import DocxTemplate


class NextValue:
    def __init__(self, values: list):
        self._it = iter(values)

    def __str__(self) -> str:
        try:
            v = next(self._it)
        except StopIteration:
            v = ""
        return "" if v is None else str(v)

    def __html__(self) -> str:
        return str(self)


def safe_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name if name else "alumno"


def _read_source_bytes(source: Any) -> bytes:
    if isinstance(source, (str, Path)):
        path = Path(source).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"No existe el archivo: {path}")
        return path.read_bytes()
    if hasattr(source, "getvalue"):
        return source.getvalue()
    raise TypeError("El origen debe ser una ruta o un archivo subido")


def ensure_template(template_path: Path | None) -> Path:
    if template_path and template_path.exists():
        return template_path

    default_template = Path(__file__).resolve().parent.parent / "templates" / "test_template.docx"
    default_template.parent.mkdir(parents=True, exist_ok=True)

    if not default_template.exists():
        doc = Document()
        doc.add_heading("Plantilla de prueba", level=1)
        doc.add_paragraph("Pregunta 1")
        doc.add_paragraph("A. Opción A")
        doc.add_paragraph("B. Opción B")
        doc.add_paragraph("C. Opción C")
        doc.add_paragraph("D. Opción D")
        doc.save(str(default_template))

    return default_template


def read_usedrange(excel_path: Path) -> list[dict]:
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))
        try:
            sh = wb.sheets[0]
            data = sh.used_range.value
            if not data or len(data) < 2:
                return []
            headers = [str(h).strip() for h in data[0]]
            rows: list[dict] = []
            for row in data[1:]:
                if row is None or all(c in (None, "") for c in row):
                    continue
                rows.append({headers[i]: row[i] if i < len(row) else None for i in range(len(headers))})
            return rows
        finally:
            wb.close()
    finally:
        app.quit()


def _norm(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def build_shuffled_version(rows: list[dict], rng: random.Random) -> tuple[list[dict], list[str]]:
    rows_copy = [dict(r) for r in rows]
    rng.shuffle(rows_copy)
    out: list[dict] = []
    correct_letters: list[str] = []
    letters = ["A", "B", "C", "D"]

    for r in rows_copy:
        answers = [r.get("ans1", ""), r.get("ans2", ""), r.get("ans3", ""), r.get("ans4", "")]
        correct_value = r.get("ans_correct", "")
        rng.shuffle(answers)
        idx = None
        cv = _norm(correct_value)
        for i, a in enumerate(answers):
            if _norm(a) == cv:
                idx = i
                break
        correct_letter = letters[idx] if idx is not None else "?"
        new_r = dict(r)
        new_r["ans1"], new_r["ans2"], new_r["ans3"], new_r["ans4"] = answers
        out.append(new_r)
        correct_letters.append(correct_letter)
    return out, correct_letters


def encode_answers_for_qr(correct_letters: list[str]) -> str:
    mapping = {"A": "*", "B": "?", "C": "0", "D": "!"}
    encoded = "".join(mapping.get(ch, "?") for ch in correct_letters)
    return "ANS" + encoded


def create_qr_png(payload: str, png_path: Path) -> None:
    img = qrcode.make(payload)
    img.save(str(png_path))


def insert_qr_at_end(docx_path: Path, qr_png_path: Path, width_cm: float = 1.5) -> None:
    doc = Document(str(docx_path))
    doc.add_paragraph("")
    doc.add_picture(str(qr_png_path), width=Cm(width_cm))
    doc.save(str(docx_path))


def generate_doc(template_path: Path, output_path: Path, rows_shuffled: list[dict]) -> None:
    questions = [r.get("question", "") for r in rows_shuffled]
    n = len(rows_shuffled)
    ctx = {
        "test": questions,
        "numpreg": list(range(1, n + 1)),
        "ans1": NextValue([r.get("ans1", "") for r in rows_shuffled]),
        "ans2": NextValue([r.get("ans2", "") for r in rows_shuffled]),
        "ans3": NextValue([r.get("ans3", "") for r in rows_shuffled]),
        "ans4": NextValue([r.get("ans4", "") for r in rows_shuffled]),
    }
    doc = DocxTemplate(str(template_path))
    doc.render(ctx)
    doc.save(str(output_path))


def run_test_generation(source: Any, output_dir: Path | str, template_source: Path | str | None = None) -> None:
    output_dir = Path(output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    payload = _read_source_bytes(source)
    with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(payload)
        temp_path = Path(tmp.name)

    template_path = ensure_template(Path(template_source).expanduser().resolve() if template_source else None)

    rows = read_usedrange(temp_path)
    if not rows:
        raise ValueError("No se han leído filas del Excel.")

    required = {"question", "ans1", "ans2", "ans3", "ans4", "ans_correct"}
    missing = required - set(rows[0].keys())
    if missing:
        raise ValueError(f"Faltan columnas: {sorted(missing)}")

    alumnos = ["alumno1", "alumno2", "alumno3", "alumno4", "alumno5"]
    for student in alumnos:
        seed = (hash(student) & 0xFFFFFFFF)
        rng = random.Random(seed)
        rows_shuffled, correct_letters = build_shuffled_version(rows, rng)
        qr_payload = encode_answers_for_qr(correct_letters)
        filename = safe_filename(student) + ".docx"
        output_path = output_dir / filename
        generate_doc(template_path, output_path, rows_shuffled)
        qr_png_path = output_dir / (safe_filename(student) + "_QR.png")
        create_qr_png(qr_payload, qr_png_path)
        insert_qr_at_end(output_path, qr_png_path, width_cm=1.5)

    st.success(f"Se han generado los test en {output_dir}")
    temp_path.unlink(missing_ok=True)
