# IES Creator

Aplicación Streamlit para lanzar tareas de profesor desde un único panel.

## Requisitos

- Python 3.10+
- Windows para la parte de Word/Excel con `xlwings` y `win32com`

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

## Funcionalidades

- Feedback DOCX: genera documentos Word por alumno a partir de un Excel seleccionado.
- Generador de tests: genera documentos Word con preguntas aleatorias y QR.
- Se pueden elegir archivos de entrada y carpetas de salida desde la propia interfaz.
