# IES Creator — contexto del proyecto

> Archivo de control para Claude Code (se carga solo en cada conversación).
> **Mantenerlo al día**: al terminar un cambio relevante, actualizar las secciones
> "Estado actual" y "Registro de cambios". Así no hace falta releer todo el código.

## Qué es
App **Streamlit** (Python 3.10+, Windows) con utilidades para un profesor de IES.
Todo se procesa **en memoria**: se suben archivos con `st.file_uploader`
y se descargan resultados (ZIP) con `st.download_button`. UI y comentarios en **español**.

- Arranque: `run_app.bat` (activa `.venv` y lanza `streamlit run app.py --server.port 8501`).
- Dependencias: [requirements.txt](requirements.txt) (streamlit, openpyxl, python-docx, docxtpl,
  qrcode, Pillow, opencv-python-headless, numpy).
- Python del venv: `.venv/Scripts/python.exe`.
- No hay tests automatizados ni linter configurado.

## Estructura
| Ruta | Contenido |
|---|---|
| [app.py](app.py) | Toda la UI. CSS + cabecera, navegación con `st.segmented_control` (key `active_page`; no se usa `st.tabs` porque pierde la pestaña en cada rerun). Páginas: Inicio, Generador de tests, Generador de QR (una rama `if/elif` por página). Los módulos de `tools/` se importan **dentro** de cada rama (import perezoso). |
| [tools/test_app.py](tools/test_app.py) | Página **Generador de tests** + lectura de QR + lectura en voz alta |
| [tools/qr_app.py](tools/qr_app.py) | Página **Generador de QR** |
| `templates/` | Solo se versionan `test_template_INI.docx` y `Qr_Template.xlsx` (ver `.gitignore`) |
| `salida/` | Ejemplos/pruebas locales, ignorado por git |
| `src/` | Vacío, sin uso |

## Convenciones del código
- Cada módulo de `tools/` expone una función `run_*(source, ...) -> bytes` que la UI llama.
- `_read_source_bytes(source)` (acepta UploadedFile / ruta) y `safe_filename` están duplicados en cada módulo, a propósito: módulos independientes.
- Los resultados se guardan en `st.session_state` (`test_zip`, `qr_zip`, `qr_letters`...) y un botón "Limpiar" hace `pop` + `st.rerun()`.
- Errores: `try/except Exception` en la UI → `st.error(...)`. Los módulos lanzan `ValueError` con mensaje en español.
- Componentes HTML/JS propios: plantilla `r"""..."""` en el módulo + `build_*_html()` que la rellena con JSON,
  renderizado con `components.html(...)`.
- `from __future__ import annotations` y type hints en todos los módulos.

## Detalle por herramienta

### Generador de tests (`test_app.py`)
- Excel (hoja activa) con columnas `question, ans1, ans2, ans3, ans4, ans_correct`.
- `run_test_generation(excel, plantilla, num_students)` → ZIP con un .docx por alumno:
  preguntas y respuestas barajadas (`build_shuffled_version`), render con `docxtpl`
  (contexto: `test`, `numpreg`, `ans1..ans4` como `NextValue`, un iterador que avanza en cada `{{ }}`),
  y QR al final (`insert_qr_at_end`).
- Codificación QR: `"ANS"` + símbolos `A→*  B→?  C→0  D→!` (`encode_answers_for_qr` / `decode_answers_from_qr_payload`).
- Lectura QR: foto con `st.camera_input` o subida → `decode_qr_image` (OpenCV) → letras en
  `session_state.qr_letters` → `build_tts_component_html` (Web Speech API, ritmo ajustable).
- Keys de session_state de esta página que empiezan por `qr_` (`qr_scanning`, `qr_letters`, `qr_show_result`...) son del **lector**, no del Generador de QR.

### Generador de QR (`qr_app.py`)
- Plantilla: [templates/Qr_Template.xlsx](templates/Qr_Template.xlsx) (hoja activa, fila 1 = cabeceras `ID`, `CODE`;
  se localizan por nombre, sin importar mayúsculas ni orden). La UI ofrece descargarla.
- `run_qr_generation(excel)` → ZIP con un PNG por fila: contenido = `CODE`, nombre de archivo = `ID`
  (si falta el ID se usa el CODE; IDs repetidos → `_2`, `_3`...). Filas sin CODE se saltan;
  floats enteros (12.0) se escriben como "12".
- **Diseño** (`QRDesign`): usa `StyledPilImage` de qrcode 8.x. `PRESETS` (Clásico → Textura de imagen) rellenan
  forma de módulos (`MODULE_STYLES`), esquinas/ojos (`EYE_STYLES`), relleno (`FILL_STYLES`: sólido, 4 degradados
  o `ImageColorMask` con textura subida) y colores. En la UI, las keys de los widgets de "Personalizar" llevan el
  nombre del preset para que se reinicien al cambiarlo.
- Imagen central opcional: se centra en un lienzo cuadrado del color de fondo (`_prepare_logo`) y, si hay logo,
  la corrección de errores pasa a H. Tamaño 10–30 % del ancho.
- Vista previa con el primer CODE del Excel + comprobación con `test_app.decode_qr_image` (OpenCV).
  Probado: todos los presets se leen; con logo, los de barras no los lee OpenCV (la UI avisa).

## Problemas conocidos / pendientes
- `test_app.ensure_template` busca `templates/test_template.docx`, pero la plantilla versionada es
  `test_template_INI.docx` → sin plantilla subida se genera una plantilla mínima de prueba.

## Estado actual
- 2 herramientas: Generador de tests (+QR y lectura en voz alta) y Generador de QR (Excel ID/CODE → PNGs, con estilos e imagen central).

## Registro de cambios
Añadir una línea por sesión de trabajo (fecha — qué se hizo).
- 2026-10-26 — Creado este CLAUDE.md.
- 2026-10-26 — Eliminadas Feedback, Sesiones por trimestre y Programación de aula (y sus plantillas
  y dependencias requests/bs4/jinja2/docxcompose). Añadido Generador de QR (`tools/qr_app.py`).
- 2026-10-26 — Generador de QR adaptado a `Qr_Template.xlsx` (columnas ID/CODE), solo PNGs (sin Word).
- 2026-10-26 — Generador de QR: estilos de diseño (presets + personalización), imagen central y vista previa.
