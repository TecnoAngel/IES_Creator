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
PAGES = ["Inicio", "Feedback", "Generador de tests", "Sesiones por trimestre", "Programación de aula"]
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

# PÁGINA: FEEDBACK DOCX
elif page == "Feedback":
    st.markdown('<div class="panel-card"><h3>Feedback</h3><p>Se leerá un archivo Excel y se generará el feedback por alumno: como documentos Word para descargar, o como HTML para copiar y pegar directamente en una plataforma web.</p></div>', unsafe_allow_html=True)

    st.markdown("### Archivos")

    fb_excel = st.file_uploader(
        "Archivo Excel",
        type=["xlsx", "xlsm"],
        key="fb_excel_uploader",
    )

    st.divider()

    fb_modo = st.radio(
        "¿Qué quieres generar?",
        ["Documentos Word", "HTML para copiar y pegar"],
        horizontal=True,
        key="fb_modo",
    )

    if fb_modo == "Documentos Word":
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

    else:
        st.caption(
            "Escribe (opcionalmente) una introducción general con formato básico, y pulsa el botón "
            "de cada alumno para copiar al portapapeles la introducción + su rúbrica en HTML."
        )

        students = []
        if fb_excel is not None:
            try:
                from tools.feedback_app import read_students_html

                students = read_students_html(fb_excel)
                st.success(f"{len(students)} alumno(s) detectado(s).")
            except Exception as exc:
                st.error(f"Error al leer el Excel: {exc}")

        from tools.feedback_app import build_component_html

        components.html(build_component_html(students), height=700, scrolling=True)

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

# PÁGINA: SESIONES POR TRIMESTRE
elif page == "Sesiones por trimestre":
    st.markdown('<div class="panel-card"><h3>Sesiones por trimestre</h3><p>Estima cuántas sesiones lectivas tiene una materia en cada trimestre, a partir del calendario escolar oficial de Castilla y León.</p></div>', unsafe_allow_html=True)

    from tools.sessions_calculator import (
        DEFAULT_CALENDAR_URL,
        approx_sessions,
        default_trimester_ranges,
        fetch_school_calendar,
        free_weeks_in_range,
        trimester_index_for_date,
    )

    st.caption(
        "Los selectores de fecha muestran el calendario según el idioma configurado en tu navegador; "
        "con el navegador en español la semana ya empieza en lunes."
    )

    st.markdown("### Calendario escolar oficial")
    calendar_url = st.text_input("URL del calendario escolar (JCyL)", value=DEFAULT_CALENDAR_URL, key="calendar_url")

    if st.button("Cargar calendario oficial", key="btn_load_calendar"):
        try:
            calendar = fetch_school_calendar(calendar_url)
            st.session_state.school_calendar = calendar
            st.session_state.trimester_dates = dict(zip(("t1", "t2", "t3"), default_trimester_ranges(calendar)))
            st.success("Calendario cargado correctamente.")
        except Exception as exc:
            st.error(f"No se ha podido leer el calendario: {exc}")

    calendar = st.session_state.get("school_calendar")

    if calendar is None:
        st.info("Carga el calendario oficial para continuar.")
    else:
        st.markdown(
            f"**Curso ESO:** {calendar.course_start:%d/%m/%Y} — {calendar.course_end:%d/%m/%Y}  \n"
            f"**Vacaciones de Navidad:** {min(calendar.christmas_days):%d/%m/%Y} — {max(calendar.christmas_days):%d/%m/%Y} "
            f"({len(calendar.christmas_days)} días)  \n"
            f"**Vacaciones de Semana Santa:** {min(calendar.easter_days):%d/%m/%Y} — {max(calendar.easter_days):%d/%m/%Y} "
            f"({len(calendar.easter_days)} días)"
        )

        st.divider()
        st.markdown("### Trimestres")
        st.caption("Fechas propuestas automáticamente en torno a Navidad y Semana Santa. Puedes ajustarlas.")

        defaults = st.session_state.get("trimester_dates", {})
        cols = st.columns(3)
        trimester_ranges = []
        for col, key, label in zip(cols, ("t1", "t2", "t3"), ("1º trimestre", "2º trimestre", "3º trimestre")):
            with col:
                st.markdown(f"**{label}**")
                default_start, default_end = defaults.get(key, (calendar.course_start, calendar.course_end))
                start = st.date_input(
                    "Inicio", value=default_start,
                    min_value=calendar.course_start, max_value=calendar.course_end, key=f"{key}_start",
                )
                end = st.date_input(
                    "Fin", value=default_end,
                    min_value=calendar.course_start, max_value=calendar.course_end, key=f"{key}_end",
                )
                trimester_ranges.append((start, end))

        st.divider()
        st.markdown("### Otros festivos y días no lectivos")
        st.caption(
            "Además de Navidad y Semana Santa, marca qué festivos oficiales quieres restar de las semanas "
            "lectivas (si caen en mitad de la semana, esa semana cuenta como incompleta)."
        )

        extra_holiday_days: set = set()
        if not calendar.other_holidays:
            st.caption("No se han encontrado más festivos en el calendario cargado.")
        else:
            for i, group in enumerate(calendar.other_holidays):
                date_range = (
                    f"{group.start:%d/%m/%Y}"
                    if group.start == group.end
                    else f"{group.start:%d/%m/%Y} – {group.end:%d/%m/%Y}"
                )
                label = f"{group.name} ({date_range}, {len(group.days)} día{'s' if len(group.days) != 1 else ''})"
                checked = st.checkbox(label, value=True, key=f"holiday_{i}")
                if checked:
                    extra_holiday_days |= group.days

        st.divider()
        st.markdown("### Fiesta local")
        st.caption(
            "Cada localidad tiene su(s) propia(s) fiesta(s) local(es), que no aparecen en el calendario "
            "autonómico. Actívala e indica su fecha para restar ese día del trimestre correspondiente."
        )

        use_local_holiday = st.checkbox("Añadir fiesta local", value=False, key="use_local_holiday")
        local_holiday_day = None
        if use_local_holiday:
            local_holiday_day = st.date_input(
                "Fecha de la fiesta local",
                value=calendar.course_start,
                min_value=calendar.course_start,
                max_value=calendar.course_end,
                key="local_holiday_date",
            )
            idx = trimester_index_for_date(local_holiday_day, trimester_ranges)
            if idx is None:
                st.warning("Esa fecha no cae dentro de ninguno de los trimestres definidos arriba.")
            else:
                st.caption(f"Se restará del {idx + 1}º trimestre.")

        st.divider()
        st.markdown("### Materia")
        col1, col2 = st.columns(2)
        with col1:
            weekly_hours = st.selectbox(
                "Horas semanales de la materia",
                options=list(range(1, 11)),
                index=2,
                key="weekly_hours",
            )
        with col2:
            safety_pct = st.slider(
                "Factor de seguridad (% de sesiones que se prevén perder: excursiones, fiestas...)",
                min_value=0, max_value=50, value=15, step=1, key="safety_factor",
            )
        safety_factor = safety_pct / 100

        st.divider()
        st.markdown("### Resultado")

        total_sessions = 0.0
        labels = ("1º trimestre", "2º trimestre", "3º trimestre")
        for t_idx, (label, (start, end)) in enumerate(zip(labels, trimester_ranges)):
            if end < start:
                st.warning(f"{label}: la fecha de fin es anterior a la de inicio.")
                continue
            holiday_days = calendar.holiday_days | extra_holiday_days
            if use_local_holiday and local_holiday_day is not None and t_idx == trimester_index_for_date(local_holiday_day, trimester_ranges):
                holiday_days = holiday_days | {local_holiday_day}
            weeks = free_weeks_in_range(start, end, holiday_days)
            sessions = approx_sessions(weeks, weekly_hours, safety_factor)
            total_sessions += sessions
            st.markdown(f"**{label}:** {weeks:.1f} semanas netas → **{sessions:.0f} sesiones aprox.**")

        st.markdown(f"#### Total curso: **{total_sessions:.0f} sesiones aprox.**")

# PÁGINA: PROGRAMACIÓN DE AULA
elif page == "Programación de aula":
    st.markdown('<div class="panel-card"><h3>Programación de aula</h3><p>Se leerá un Excel con los datos del grupo y las situaciones de aprendizaje, y se generará el documento Word de la programación de aula.</p></div>', unsafe_allow_html=True)

    st.markdown("### Archivos")

    pa_excel = st.file_uploader(
        "Excel de datos (hojas 'Datos generales' y 'Situaciones de aprendizaje')",
        type=["xlsx", "xlsm"],
        key="pa_excel_uploader",
    )
    pa_template = st.file_uploader(
        "Plantilla Word (opcional, se usará la plantilla por defecto si no se sube)",
        type=["docx"],
        key="pa_template_uploader",
    )

    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Generar programación", use_container_width=True, type="primary", key="btn_gen_pa"):
            if pa_excel is None:
                st.error("Necesitas subir el Excel de datos.")
            else:
                try:
                    from tools.programacion_aula import run_programacion_aula

                    docx_bytes = run_programacion_aula(pa_excel, pa_template)
                    st.session_state.pa_docx = docx_bytes
                    st.success("Proceso completado.")
                except Exception as exc:
                    st.error(f"Error al generar la programación: {exc}")

    with col2:
        if st.button("Limpiar", use_container_width=True, key="btn_clear_pa"):
            st.session_state.pop("pa_docx", None)
            st.rerun()

    if st.session_state.get("pa_docx"):
        st.download_button(
            "Descargar programación (Word)",
            data=st.session_state.pa_docx,
            file_name="programacion_aula.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
