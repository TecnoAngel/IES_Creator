from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

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
        <div class="topbar-sub">Herramientas para preparar materiales y feedback con un flujo más ordenado.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("")

tabs = st.tabs(["Inicio", "Feedback DOCX", "Generador de tests"])

# TAB 0: INICIO
with tabs[0]:
    st.markdown('<div class="panel-card"><h3>Bienvenido</h3><p>Selecciona <strong>Feedback DOCX</strong> para generar documentos de feedback personalizados para cada alumno a partir de un archivo Excel.</p><p>O selecciona <strong>Generador de tests</strong> para crear tests en Word con códigos QR a partir de un archivo Excel.</p></div>', unsafe_allow_html=True)

# TAB 1: FEEDBACK DOCX
with tabs[1]:
    st.markdown('<div class="panel-card"><h3>Feedback DOCX</h3><p>Se leerá un archivo Excel y se generarán documentos Word por alumno.</p></div>', unsafe_allow_html=True)

    st.markdown("### Archivos")

    fb_excel = st.file_uploader(
        "Archivo Excel",
        type=["xlsx", "xlsm"],
        key="fb_excel_uploader",
    )

    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Generar feedback", use_container_width=True, type="primary", key="btn_gen_feedback"):
            if fb_excel is None:
                st.error("Necesitas subir un archivo Excel válido.")
            else:
                try:
                    from tools.feedback_app import run_feedback

                    zip_bytes = run_feedback(fb_excel)
                    st.session_state.fb_zip = zip_bytes
                    st.success("Proceso completado.")
                except Exception as exc:
                    st.error(f"Error al generar el feedback: {exc}")

    with col2:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_fb"):
            st.session_state.pop("fb_zip", None)
            st.rerun()

    if st.session_state.get("fb_zip"):
        st.download_button(
            "Descargar documentos (ZIP)",
            data=st.session_state.fb_zip,
            file_name="feedback.zip",
            mime="application/zip",
            use_container_width=True,
        )

# TAB 2: GENERADOR DE TESTS
with tabs[2]:
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
