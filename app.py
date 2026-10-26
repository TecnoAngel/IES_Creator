from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

ROOT = Path(__file__).resolve().parent
sys.path.append(str(ROOT))


st.set_page_config(page_title="IES Creator", page_icon="🎓", layout="wide")

st.markdown(
    """
    <style>
    .topbar {
        padding: 1.1rem 1.2rem;
        border-radius: 16px;
        background: linear-gradient(135deg, #ff006f 0%, #0047ab 50%, #1a1a1a 100%);
        color: white;
        margin-bottom: 1rem;
        box-shadow: 0 6px 18px rgba(255,0,111,0.25);
    }
    .topbar-title { font-size: 1.8rem; font-weight: 700; line-height: 1.1; }
    .topbar-sub { font-size: 0.95rem; opacity: 0.95; margin-top: 0.35rem; }
    .section-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 1.5rem;
        font-weight: 700;
        margin-top: 0;
        margin-bottom: 0.5rem;
    }
    .panel-card {
        border-radius: 14px;
        padding: 1rem 1.1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .panel-card h3 {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-top: 0;
    }
    @media (prefers-color-scheme: light) {
        .panel-card {
            background: #fff;
            border: 1px solid #f1e3d4;
            color: #333;
        }
    }
    @media (prefers-color-scheme: dark) {
        .panel-card {
            background: #1a1a2e;
            border: 1px solid #333;
            color: #e0e0e0;
        }
    }
    </style>
    <div class="topbar">
        <div style="font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.18em; opacity: 0.9; margin-bottom: 0.25rem;">Panel de profesor</div>
        <div class="topbar-title">IES Creator</div>
        <div class="topbar-sub">Mi cajita de utilidades.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("")

# st.tabs no conserva la pestaña activa entre reruns (cada clic en un botón
# devuelve la vista a la primera pestaña), así que la navegación se hace con
# un widget normal atado a session_state para que sea persistente.
PAGES = ["Inicio", "Generador de tests", "Generador de QR"]
page = st.segmented_control(
    "Navegación",
    PAGES,
    default=PAGES[0],
    key="active_page",
    required=True,
    label_visibility="collapsed",
)

# PÁGINA: INICIO
if page == "Inicio":
    st.markdown('<div class="panel-card"><h3>Bienvenido</h3><p>Elige una sección arriba para empezar.</p></div>', unsafe_allow_html=True)

# PÁGINA: GENERADOR DE TESTS
elif page == "Generador de tests":
    st.markdown('<div class="panel-card"><h3>Generador de tests</h3><p>Se leerá un Excel con preguntas y se generarán test en Word con códigos QR.</p></div>', unsafe_allow_html=True)

    st.markdown("### Archivos")

    test_excel = st.file_uploader(
        "Archivo Excel con preguntas",
        type=["xlsx", "xlsm"],
        key="test_excel_uploader",
    )
    test_template = st.file_uploader(
        "Plantilla Word (opcional, se usará una por defecto si no se sube)",
        type=["docx"],
        key="test_template_uploader",
    )
    num_students = st.number_input("Número de alumnos", min_value=1, max_value=200, value=5, step=1)

    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Generar tests", use_container_width=True, type="primary", key="btn_gen_tests"):
            if test_excel is None:
                st.error("Necesitas subir un archivo Excel válido.")
            else:
                try:
                    from tools.test_app import run_test_generation

                    zip_bytes = run_test_generation(test_excel, test_template, num_students=int(num_students))
                    st.session_state.test_zip = zip_bytes
                    st.success("Proceso completado.")
                except Exception as exc:
                    st.error(f"Error al generar los tests: {exc}")

    with col2:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_tests"):
            st.session_state.pop("test_zip", None)
            st.rerun()

    if st.session_state.get("test_zip"):
        st.download_button(
            "Descargar tests (ZIP)",
            data=st.session_state.test_zip,
            file_name="tests.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.divider()
    st.markdown("### Leer QR y escuchar las respuestas")
    st.caption("Escanea el QR de un test corregido para guardar sus respuestas, y léelas después cuando quieras (con opción de escucharlas en voz alta a un ritmo ajustable).")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📷 Escanear código", use_container_width=True, key="btn_qr_scan"):
            st.session_state.qr_scanning = True
            st.session_state.qr_show_result = False
    with col2:
        if st.button(
            "🔊 Leer últimas respuestas",
            use_container_width=True,
            key="btn_qr_show",
            disabled=not st.session_state.get("qr_letters"),
        ):
            st.session_state.qr_show_result = True

    if st.session_state.get("qr_scanning"):
        qr_photo = st.camera_input("Encuadra el QR y haz la foto", key="qr_camera")
        qr_uploaded = st.file_uploader("...o sube una foto del QR", type=["png", "jpg", "jpeg"], key="qr_uploader")
        qr_source = qr_photo if qr_photo is not None else qr_uploaded

        if qr_source is not None:
            try:
                from tools.test_app import decode_answers_from_qr_payload, decode_qr_image

                payload = decode_qr_image(qr_source.getvalue())
                letters = decode_answers_from_qr_payload(payload)
                st.session_state.qr_letters = letters
                st.session_state.qr_scanning = False
                st.session_state.qr_show_result = True
                st.rerun()
            except Exception as exc:
                st.error(f"Error al leer el QR: {exc}")

    if st.session_state.get("qr_show_result") and st.session_state.get("qr_letters"):
        letters = st.session_state.qr_letters

        from tools.test_app import build_tts_component_html

        components.html(build_tts_component_html(letters), height=420, scrolling=True)

# PÁGINA: GENERADOR DE QR
elif page == "Generador de QR":
    st.markdown('<div class="panel-card"><h3>Generador de QR</h3><p>Se leerá un Excel con columnas ID y CODE y se generará un código QR en PNG por fila, con el ID como nombre de archivo. Puedes elegir el diseño y añadir una imagen en el centro.</p></div>', unsafe_allow_html=True)

    st.markdown("### Archivos")

    from tools.qr_app import DEFAULT_TEMPLATE

    if DEFAULT_TEMPLATE.exists():
        st.download_button(
            "Descargar plantilla Excel (Qr_Template.xlsx)",
            data=DEFAULT_TEMPLATE.read_bytes(),
            file_name=DEFAULT_TEMPLATE.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_qr_template",
        )

    qr_excel = st.file_uploader(
        "Archivo Excel (columnas ID y CODE)",
        type=["xlsx", "xlsm"],
        key="qr_excel_uploader",
    )

    st.divider()
    st.markdown("### Diseño")

    from tools.qr_app import EYE_STYLES, FILL_STYLES, MODULE_STYLES, PRESETS, QRDesign, create_qr_png_bytes

    qr_preset = st.selectbox("Estilo", list(PRESETS), key="qr_preset")
    base = QRDesign.from_preset(qr_preset)

    # Las keys llevan el nombre del estilo para que, al cambiarlo, los ajustes
    # vuelvan a sus valores en lugar de quedarse con los del estilo anterior.
    with st.expander("Personalizar diseño"):
        col1, col2, col3 = st.columns(3)
        with col1:
            qr_module = st.selectbox("Forma de los módulos", list(MODULE_STYLES), index=list(MODULE_STYLES).index(base.module), key=f"qr_module_{qr_preset}")
        with col2:
            qr_eye = st.selectbox("Forma de las esquinas", list(EYE_STYLES), index=list(EYE_STYLES).index(base.eye), key=f"qr_eye_{qr_preset}")
        with col3:
            qr_fill = st.selectbox("Relleno", FILL_STYLES, index=FILL_STYLES.index(base.fill), key=f"qr_fill_{qr_preset}")

        col1, col2, col3 = st.columns(3)
        with col1:
            qr_color1 = st.color_picker("Color principal", base.color1, key=f"qr_color1_{qr_preset}")
        with col2:
            qr_color2 = st.color_picker("Segundo color (degradados)", base.color2, key=f"qr_color2_{qr_preset}")
        with col3:
            qr_back = st.color_picker("Fondo", base.back, key=f"qr_back_{qr_preset}")
        st.caption("Para que se lea bien, los módulos deben ser bastante más oscuros que el fondo.")

    qr_texture = None
    if qr_fill == "Textura (imagen)":
        qr_texture_file = st.file_uploader(
            "Imagen de textura (sus colores rellenan los módulos; mejor si es oscura)",
            type=["png", "jpg", "jpeg"],
            key="qr_texture_uploader",
        )
        qr_texture = qr_texture_file.getvalue() if qr_texture_file else None

    qr_logo_file = st.file_uploader(
        "Imagen para el centro (opcional)",
        type=["png", "jpg", "jpeg"],
        key="qr_logo_uploader",
    )
    qr_logo_ratio = 0.25
    if qr_logo_file is not None:
        qr_logo_ratio = st.slider("Tamaño de la imagen central (% del ancho del QR)", 10, 30, 25, key="qr_logo_ratio") / 100

    qr_design = QRDesign(
        module=qr_module,
        eye=qr_eye,
        fill=qr_fill,
        color1=qr_color1,
        color2=qr_color2,
        back=qr_back,
        logo=qr_logo_file.getvalue() if qr_logo_file else None,
        logo_ratio=qr_logo_ratio,
        texture=qr_texture,
    )

    st.markdown("#### Vista previa")
    preview_content = "https://www.ejemplo.com"
    if qr_excel is not None:
        try:
            from tools.qr_app import read_qr_rows

            preview_content = read_qr_rows(qr_excel.getvalue())[0][1]
        except Exception:
            pass

    try:
        preview_png = create_qr_png_bytes(preview_content, qr_design)
        col1, col2 = st.columns([1, 2])
        with col1:
            st.image(preview_png, width=240)
        with col2:
            st.caption(f"Contenido de la vista previa: {preview_content}")
            from tools.test_app import decode_qr_image

            try:
                ok = decode_qr_image(preview_png) == preview_content
            except Exception:
                ok = False
            if ok:
                st.success("Comprobado: el QR se lee correctamente.")
            else:
                st.warning(
                    "El lector automático no ha podido leer este diseño. Puede funcionar igualmente con el móvil "
                    "(los diseños artísticos le cuestan más al lector), pero pruébalo antes de imprimir; "
                    "si no, sube el contraste, reduce la imagen central o usa un estilo más sencillo."
                )
    except Exception as exc:
        st.info(str(exc))

    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Generar QR", use_container_width=True, type="primary", key="btn_gen_qr"):
            if qr_excel is None:
                st.error("Necesitas subir un archivo Excel válido.")
            else:
                try:
                    from tools.qr_app import run_qr_generation

                    zip_bytes = run_qr_generation(qr_excel, qr_design)
                    st.session_state.qr_zip = zip_bytes
                    st.success("Proceso completado.")
                except Exception as exc:
                    st.error(f"Error al generar los QR: {exc}")

    with col2:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_qr"):
            st.session_state.pop("qr_zip", None)
            st.rerun()

    if st.session_state.get("qr_zip"):
        st.download_button(
            "Descargar QR (ZIP)",
            data=st.session_state.qr_zip,
            file_name="qr.zip",
            mime="application/zip",
            use_container_width=True,
        )
