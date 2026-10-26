# IES Creator

Aplicación Streamlit para lanzar tareas de profesor desde un único panel.

## Requisitos

- Python 3.10+

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app.py
```

(o doble clic en `run_app.bat`, que usa el entorno `.venv`).

## Funcionalidades

- Generador de tests: genera documentos Word con preguntas aleatorias y QR con las respuestas; permite escanear el QR y escuchar las respuestas.
- Generador de QR: a partir de un Excel con columnas `ID` y `CODE` (plantilla en `templates/Qr_Template.xlsx`), genera un PNG por fila con el contenido de `CODE`, llamado como su `ID`. Varios estilos de diseño (de clásico a degradados, puntos, barras o textura de imagen), personalizables, con imagen central opcional y vista previa con comprobación de lectura.
