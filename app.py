from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.append(str(ROOT))


def pick_file(title: str, filetypes: list[tuple[str, str]]) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return Path(path).expanduser().resolve() if path else None


def pick_folder(title: str) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return Path(path).expanduser().resolve() if path else None


def open_path(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))
    else:
        st.info(f"Abre la carpeta manualmente: {path}")


if "input_path" not in st.session_state:
    st.session_state.input_path = ""
if "template_path" not in st.session_state:
    st.session_state.template_path = ""
if "output_dir" not in st.session_state:
    st.session_state.output_dir = str((ROOT / "salida").resolve())

st.set_page_config(page_title="IES Creator", page_icon="�", layout="wide")

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

    input_path = Path(st.session_state.input_path).resolve() if st.session_state.input_path else None
    output_dir = Path(st.session_state.output_dir).resolve() if st.session_state.output_dir else ROOT / "salida"
    
    st.markdown("### Archivos")
    
    # Archivo Excel
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("Elegir Excel", use_container_width=True, key="fb_excel"):
            selected = pick_file(
                "Selecciona el archivo Excel",
                [("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
            )
            if selected:
                st.session_state.input_path = str(selected)
                st.rerun()
    with col2:
        if input_path and input_path.exists():
            st.success(f"{input_path}")
        else:
            st.warning("Selecciona un archivo Excel")
    
    # Carpeta de salida
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("Elegir salida", use_container_width=True, key="fb_output"):
            selected = pick_folder("Selecciona la carpeta de resultados")
            if selected:
                st.session_state.output_dir = str(selected)
                st.rerun()
    with col2:
        st.info(f"{output_dir}")
    
    st.divider()
    
    # Botón de generar y abrir carpeta
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        if st.button("Generar feedback", use_container_width=True, type="primary", key="btn_gen_feedback"):
            if not input_path or not input_path.exists():
                st.error("Necesitas elegir un archivo Excel válido.")
            else:
                from tools.feedback_app import run_feedback
                run_feedback(input_path, output_dir)
                st.success("Proceso completado.")
    
    with col2:
        if st.button("Abrir salida", use_container_width=True, key="btn_open_fb"):
            open_path(output_dir)
    
    with col3:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_fb"):
            st.session_state.input_path = ""
            st.session_state.output_dir = str((ROOT / "salida").resolve())
            st.rerun()

# TAB 2: GENERADOR DE TESTS
with tabs[2]:
    st.markdown('<div class="panel-card"><h3>Generador de tests</h3><p>Se leerá un Excel con preguntas y se generarán test en Word con códigos QR.</p></div>', unsafe_allow_html=True)

    input_path = Path(st.session_state.input_path).resolve() if st.session_state.input_path else None
    template_path = Path(st.session_state.template_path).resolve() if st.session_state.template_path else None
    output_dir = Path(st.session_state.output_dir).resolve() if st.session_state.output_dir else ROOT / "salida"
    
    st.markdown("### Archivos")
    
    # Archivo Excel
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("Elegir Excel", use_container_width=True, key="test_excel"):
            selected = pick_file(
                "Selecciona el archivo Excel con preguntas",
                [("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
            )
            if selected:
                st.session_state.input_path = str(selected)
                st.rerun()
    with col2:
        if input_path and input_path.exists():
            st.success(f"{input_path}")
        else:
            st.warning("Selecciona un archivo Excel")
    
    # Plantilla Word
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("Elegir plantilla", use_container_width=True, key="test_word"):
            selected = pick_file("Selecciona la plantilla Word", [("Word", "*.docx"), ("Todos", "*.*")])
            if selected:
                st.session_state.template_path = str(selected)
                st.rerun()
    with col2:
        if template_path and template_path.exists():
            st.success(f"{template_path}")
        else:
            st.warning("Selecciona una plantilla Word")
    
    # Carpeta de salida
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("Elegir salida", use_container_width=True, key="test_output"):
            selected = pick_folder("Selecciona la carpeta de resultados")
            if selected:
                st.session_state.output_dir = str(selected)
                st.rerun()
    with col2:
        st.info(f"{output_dir}")
    
    st.divider()
    
    # Botón de generar y abrir carpeta
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        if st.button("Generar tests", use_container_width=True, type="primary", key="btn_gen_tests"):
            if not input_path or not input_path.exists():
                st.error("Necesitas elegir un archivo Excel válido.")
            elif not template_path or not template_path.exists():
                st.error("Necesitas elegir una plantilla Word válida.")
            else:
                from tools.test_app import run_test_generation
                run_test_generation(input_path, output_dir, template_path)
                st.success("Proceso completado.")
    
    with col2:
        if st.button("Abrir salida", use_container_width=True, key="btn_open_tests"):
            open_path(output_dir)
    
    with col3:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_tests"):
            st.session_state.input_path = ""
            st.session_state.template_path = ""
            st.session_state.output_dir = str((ROOT / "salida").resolve())
            st.rerun()
