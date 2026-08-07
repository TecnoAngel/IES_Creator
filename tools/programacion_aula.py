from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any

import jinja2
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docxtpl import DocxTemplate
from openpyxl import load_workbook

# Un índice/TOC guardado en la plantilla puede dejar en caché texto de una versión
# anterior de un encabezado (p.ej. "{{ situacion.titulo }}") fuera del bucle real.
# Con el Undefined por defecto de Jinja eso rompe el render ('situacion' is undefined);
# con ChainableUndefined esas referencias sueltas se imprimen en blanco sin fallar.
_JINJA_ENV = jinja2.Environment(undefined=jinja2.ChainableUndefined)

DATOS_GENERALES_SHEET = "P_Aula_Datos_Generales"
SITUACIONES_SHEET = "P_Aula_SA"
INDICADORES_SHEET = "IndicadoresLogro"
ACTIVIDADES_SHEET = "Actividades"
CRITERIOS_SHEET_PREFIX = "CriteriosEvaluaci"  # evita líos de acentos con el nombre exacto

INDICADORES_COLUMNAS = ["CE", "IL", "CT", "IE", "CC", "AE"]
INDICADORES_CABECERAS = INDICADORES_COLUMNAS

TABLE_HEADER_FILL = "000000"
TABLE_HEADER_FONT_SIZE_PT = 9


def _read_source_bytes(source: Any) -> bytes:
    if isinstance(source, (str, Path)):
        path = Path(source).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"No existe el archivo: {path}")
        return path.read_bytes()
    if hasattr(source, "getvalue"):
        return source.getvalue()
    raise TypeError("El origen debe ser una ruta o un archivo subido")


def ensure_template(template_source: Any) -> bytes:
    if template_source is not None:
        return _read_source_bytes(template_source)

    default_template = Path(__file__).resolve().parent.parent / "templates" / "programacion_aula_template.docx"
    if not default_template.exists():
        raise FileNotFoundError(
            f"No se ha subido plantilla y no existe la plantilla por defecto: {default_template}"
        )
    return default_template.read_bytes()


def _find_sheet(wb, prefix: str):
    for name in wb.sheetnames:
        if name.startswith(prefix):
            return wb[name]
    return None


def _sheet_rows_as_dicts(ws, header_marker: str = "CE") -> list[dict]:
    """Lee una hoja localizando la fila de cabecera real: la primera que contiene
    literalmente `header_marker` como valor de celda. Algunas hojas del Excel traen
    bloques de resumen (p.ej. 'SUMA de Pesos') por encima de la tabla de datos, así
    que no basta con coger la primera fila no vacía."""
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    for i, row in enumerate(rows):
        if row and any(str(c).strip() == header_marker for c in row if c is not None):
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(h).strip() if h is not None else "" for h in rows[header_idx]]
    result: list[dict] = []
    for row in rows[header_idx + 1 :]:
        if row is None or all(c in (None, "") for c in row):
            continue
        result.append({headers[i]: row[i] for i in range(len(headers)) if i < len(row)})
    return result


def _to_number(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _to_int(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def read_datos_generales(payload: bytes) -> dict:
    wb = load_workbook(BytesIO(payload), data_only=True)
    if DATOS_GENERALES_SHEET not in wb.sheetnames:
        raise ValueError(f"El Excel no tiene una hoja llamada '{DATOS_GENERALES_SHEET}'.")
    ws = wb[DATOS_GENERALES_SHEET]
    datos: dict = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
        key = str(row[0]).strip()
        value = row[1] if len(row) > 1 else None
        datos[key] = "" if value is None else value
    return datos


def read_situaciones(payload: bytes) -> list[dict]:
    """Lee la hoja transpuesta: columna A = nombre de campo, cada columna siguiente = una
    situación de aprendizaje (una columna con cabecera no vacía = una situación)."""
    wb = load_workbook(BytesIO(payload), data_only=True)
    if SITUACIONES_SHEET not in wb.sheetnames:
        raise ValueError(f"El Excel no tiene una hoja llamada '{SITUACIONES_SHEET}'.")
    ws = wb[SITUACIONES_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header_row = rows[0]
    data_rows = rows[1:]
    situacion_columns = [i for i in range(1, len(header_row)) if header_row[i] not in (None, "")]

    situaciones: list[dict] = []
    for numero, col_idx in enumerate(situacion_columns, start=1):
        situacion: dict = {"numero": numero}
        for row in data_rows:
            if not row or row[0] is None or str(row[0]).strip() == "":
                continue
            campo = str(row[0]).strip()
            valor = row[col_idx] if col_idx < len(row) else None
            situacion[campo] = "" if valor is None else valor

        if any(v not in (None, "") for k, v in situacion.items() if k != "numero"):
            situaciones.append(situacion)

    return situaciones


def read_indicadores_por_sa(payload: bytes) -> dict[int, list[dict]]:
    """Filtra IndicadoresLogro por SA, quedándose solo con CE, IL, CT, IE, CC, AE."""
    wb = load_workbook(BytesIO(payload), data_only=True)
    if INDICADORES_SHEET not in wb.sheetnames:
        return {}
    filas = _sheet_rows_as_dicts(wb[INDICADORES_SHEET])

    por_sa: dict[int, list[dict]] = {}
    for r in filas:
        sa = _to_int(r.get("SA"))
        if sa is None:
            continue
        fila = {col: r.get(col) for col in INDICADORES_COLUMNAS}
        por_sa.setdefault(sa, []).append(fila)
    return por_sa


def compute_valores_actividades_por_sa(payload: bytes) -> dict[int, list[dict]]:
    """Replica en Python las medidas DAX de Power Pivot 'Valor_actividad_sobre_programación'
    y 'Valor_actividad_sobre_SA', para poder calcularlas para todas las situaciones
    (el Excel solo trae en caché la última que se vio con el segmentador de datos)."""
    wb = load_workbook(BytesIO(payload), data_only=True)
    if ACTIVIDADES_SHEET not in wb.sheetnames or INDICADORES_SHEET not in wb.sheetnames:
        return {}
    ws_criterios = _find_sheet(wb, CRITERIOS_SHEET_PREFIX)
    if ws_criterios is None:
        return {}

    actividades = _sheet_rows_as_dicts(wb[ACTIVIDADES_SHEET], header_marker="IL")
    indicadores = _sheet_rows_as_dicts(wb[INDICADORES_SHEET], header_marker="CE")
    criterios = _sheet_rows_as_dicts(ws_criterios, header_marker="CE")

    # IL -> {CE, PIL}
    info_por_il: dict[str, dict] = {}
    for r in indicadores:
        il = r.get("IL")
        if il is None:
            continue
        info_por_il[str(il)] = {"CE": r.get("CE"), "PIL": _to_number(r.get("PIL"))}

    # CE -> peso (P)
    peso_por_ce: dict[str, float] = {}
    for r in criterios:
        ce = r.get("CE")
        if ce is None:
            continue
        peso_por_ce[str(ce)] = _to_number(r.get("P"))
    total_suma_pesos = sum(peso_por_ce.values())

    # CE -> suma de PIL de todos sus IL
    total_pil_por_ce: dict[str, float] = {}
    for info in info_por_il.values():
        ce = info["CE"]
        if ce is None:
            continue
        total_pil_por_ce[str(ce)] = total_pil_por_ce.get(str(ce), 0.0) + info["PIL"]

    # IL -> suma de PA de todas sus actividades
    total_pa_por_il: dict[str, float] = {}
    for a in actividades:
        il = a.get("IL")
        if il is None:
            continue
        total_pa_por_il[str(il)] = total_pa_por_il.get(str(il), 0.0) + _to_number(a.get("PA"))

    resultados: list[dict] = []
    for a in actividades:
        il_raw = a.get("IL")
        il = str(il_raw) if il_raw is not None else None
        pa = _to_number(a.get("PA"))
        total_pa_il = total_pa_por_il.get(il, 0.0)
        porcentaje_a_en_il = (pa / total_pa_il) if total_pa_il else 0.0

        info_il = info_por_il.get(il, {})
        ce_raw = info_il.get("CE")
        ce = str(ce_raw) if ce_raw is not None else None
        pil_del_il = info_il.get("PIL", 0.0)
        total_pil_ce = total_pil_por_ce.get(ce, 0.0)
        porcentaje_il_en_ce = (pil_del_il / total_pil_ce) if total_pil_ce else 0.0

        peso_ce = peso_por_ce.get(ce, 0.0)
        porcentaje_ce_relativo = (peso_ce / total_suma_pesos) if total_suma_pesos else 0.0

        valor_programacion = porcentaje_a_en_il * porcentaje_il_en_ce * porcentaje_ce_relativo

        resultados.append(
            {
                "sa": _to_int(a.get("SA")),
                "A": a.get("A"),
                "DA": a.get("DA"),
                "valor_programacion": valor_programacion,
            }
        )

    total_valor_por_sa: dict[int, float] = {}
    for r in resultados:
        if r["sa"] is None:
            continue
        total_valor_por_sa[r["sa"]] = total_valor_por_sa.get(r["sa"], 0.0) + r["valor_programacion"]

    por_sa: dict[int, list[dict]] = {}
    for r in resultados:
        if r["sa"] is None:
            continue
        total_sa = total_valor_por_sa.get(r["sa"], 0.0)
        valor_sa = (r["valor_programacion"] / total_sa) if total_sa else 0.0
        por_sa.setdefault(r["sa"], []).append(
            {
                "A": r["A"],
                "DA": r["DA"],
                "valor_programacion": r["valor_programacion"],
                "valor_sa": valor_sa,
            }
        )
    return por_sa


def _fmt_pct(value: float) -> str:
    return f"{value * 100:.2f}".replace(".", ",") + " %"


def _set_cell_background(cell, color_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _set_table_borders(table) -> None:
    """Bordes finos vía XML directo: no depende de que exista el estilo con
    nombre 'Table Grid' en la plantilla (puede faltar si Word nunca lo llegó
    a aplicar y lo optimizó fuera de styles.xml al guardar)."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "999999")
        borders.append(el)
    tblPr.append(borders)


def _build_table_subdoc(tpl: DocxTemplate, headers: list[str], rows: list[list[str]], empty_msg: str):
    subdoc = tpl.new_subdoc()

    if not rows:
        p = subdoc.add_paragraph()
        run = p.add_run(empty_msg)
        run.italic = True
        return subdoc

    table = subdoc.add_table(rows=1, cols=len(headers))
    _set_table_borders(table)
    table.autofit = True

    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        header_cells[i].text = header
        header_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        paragraph = header_cells[i].paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = Pt(TABLE_HEADER_FONT_SIZE_PT)
        _set_cell_background(header_cells[i], TABLE_HEADER_FILL)

    for row_values in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row_values):
            cells[i].text = "" if value is None else str(value)
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    return subdoc


def build_tabla_indicadores(tpl: DocxTemplate, indicadores: list[dict]):
    rows = [[fila.get(col) for col in INDICADORES_COLUMNAS] for fila in indicadores]
    return _build_table_subdoc(
        tpl,
        INDICADORES_CABECERAS,
        rows,
        empty_msg="No hay indicadores de logro asociados a esta situación de aprendizaje.",
    )


def build_tabla_actividades(tpl: DocxTemplate, actividades: list[dict]):
    headers = ["Actividad", "Descripción", "% sobre la programación", "% sobre la situación"]
    rows = [
        [a.get("A"), a.get("DA"), _fmt_pct(a.get("valor_programacion", 0.0)), _fmt_pct(a.get("valor_sa", 0.0))]
        for a in actividades
    ]
    return _build_table_subdoc(
        tpl,
        headers,
        rows,
        empty_msg="No hay actividades asociadas a esta situación de aprendizaje.",
    )


def run_programacion_aula(excel_source: Any, template_source: Any = None) -> bytes:
    """Genera el .docx de la programación de aula a partir del Excel de datos."""
    payload = _read_source_bytes(excel_source)
    template_bytes = ensure_template(template_source)

    ctx = read_datos_generales(payload)
    situaciones = read_situaciones(payload)

    if not situaciones:
        raise ValueError(
            f"No se han encontrado situaciones de aprendizaje en la hoja '{SITUACIONES_SHEET}'."
        )

    tpl = DocxTemplate(BytesIO(template_bytes))

    indicadores_por_sa = read_indicadores_por_sa(payload)
    actividades_por_sa = compute_valores_actividades_por_sa(payload)

    for situacion in situaciones:
        numero = situacion["numero"]
        situacion["tabla_indicadores"] = build_tabla_indicadores(tpl, indicadores_por_sa.get(numero, []))
        situacion["tabla_actividades"] = build_tabla_actividades(tpl, actividades_por_sa.get(numero, []))

    ctx["situaciones"] = situaciones

    tpl.render(ctx, jinja_env=_JINJA_ENV)
    buffer = BytesIO()
    tpl.save(buffer)
    return buffer.getvalue()
