# Reporte Canal 13 — Versión Local 📁

Script que lee archivos descargados manualmente desde S3, consulta YouTube,
y genera el Excel del reporte automáticamente.

---

## Instalación (solo la primera vez)

```bash
pip install openpyxl pandas google-api-python-client
```

---

## Estructura de carpetas

```
ReporteCanal13/
│
├── reporte_desde_carpeta.py   ← script principal
├── config_local.py            ← configuración (editar antes de usar)
├── README.md
│
└── datos/                     ← aquí van los Excel descargados de S3
    ├── 2026-04-28_programacion.xlsx
    ├── 2026-04-29_programacion.xlsx
    ├── 2026-04-30_programacion.xlsx
    └── ...
```

---

## Configuración (`config_local.py`)

Solo necesitas cambiar **2 cosas**:

```python
# 1. Ruta de tu carpeta de datos
CARPETA_DATOS = r"C:\ReporteCanal13\datos"   # Windows
# CARPETA_DATOS = "/Users/sergio/ReporteCanal13/datos"  # Mac

# 2. Tu API key de YouTube
YOUTUBE_API_KEY = "AIzaSy..."
```

Los Channel IDs de YouTube ya están configurados:
- `UCnjvCq03NHM0LxKecdUJSCg` → El 13 (entretenimiento)
- `UCsRnhjcUCR78Q3Ud6OXCTNg` → T13 (noticias)

---

## Flujo de trabajo semanal

### Paso 1 — Descargar archivos de S3
Cada día (o al final de la semana), entra a la consola de AWS S3:
```
datos-dps → 2026-04-28 → 2026-04-28_programacion.xlsx  ← descargar este
datos-dps → 2026-04-29 → 2026-04-29_programacion.xlsx  ← y este
...
```
Guárdalos en tu carpeta `datos/` sin cambiar el nombre.

### Paso 2 — Ejecutar el script
```bash
cd C:\ReporteCanal13
python reporte_desde_carpeta.py
```

### Paso 3 — Recoger el Excel
El reporte se genera automáticamente en la misma carpeta `datos/`:
```
Reporte_28-Apr_02-May_2026.xlsx
```

---

## Notas

- El script detecta **automáticamente** todos los días disponibles en la carpeta.
  No importa si tienes 3 días o 7, siempre genera el reporte con los que hay.
- Si un programa no aparece en los archivos de esa semana, su hoja no se crea.
- Si no hay API key de YouTube configurada, el reporte igual se genera
  con los datos de S3, y las columnas de YouTube quedan con `-`.
- El reporte se **sobreescribe** cada vez que ejecutas el script con los mismos archivos.

---

## Agregar un programa nuevo

En `config_local.py`, agrega una entrada en `PROGRAMAS_CONFIG`:

```python
"Nombre del Programa": {
    "patrones_s3": [r"NOMBRE DEL PROGRAMA"],   # como aparece en el xlsx (mayúsculas)
    "youtube": True,                            # False si no sube a YouTube
    "canal_youtube": "entretenimiento",         # o "noticias" o "ambos"
    "patrones_youtube": [r"nombre del programa"],
},
```
