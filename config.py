"""
config.py — Configuración del Reporte Semanal Canal 13

Editar las variables con tus credenciales y ajustar programas según necesidad.
"""

# ─── AWS S3 ──────────────────────────────────────────────────────────────────
AWS_BUCKET  = "datos-dps"          # Nombre del bucket S3
AWS_REGION  = "us-east-1"          # Región del bucket (ajustar)

# ─── YouTube ──────────────────────────────────────────────────────────────────
YOUTUBE_API_KEY = "AIzaSyDc75yPe--BM4npEhN5Yak3xdUnbKh_0Jc"  # ⚠️ Reemplaza con tu API KEY real

# Canal 13 tiene DOS canales de YouTube:
YOUTUBE_CANAL_ENTRETENIMIENTO = "UCnjvCq03NHM0LxKecdUJSCg"  # "El 13" → programas, shows
YOUTUBE_CANAL_NOTICIAS        = "UCsRnhjcUCR78Q3Ud6OXCTNg"  # "T13"   → noticias, Teletrece

# ─── Semana a reportar ────────────────────────────────────────────────────────
# Formato: "YYYY-MM-DD" — debe ser el LUNES de la semana deseada
SEMANA_INICIO = "2026-04-20"

# ─── Programas ────────────────────────────────────────────────────────────────
# Cada programa tiene:
#   patrones_s3:      lista de regex para matchear el título en programacion.xlsx (mayúsculas)
#   youtube:          True si el programa también aparece en YouTube
#   patrones_youtube: lista de regex para matchear el título del video en YouTube
#
# Los patrones usan regex. Ejemplos:
#   r"TU DIA"        → match exacto
#   r"TU.?D[IÍ]A"   → match flexible con o sin acento y cualquier char entre palabras
#   r"3\s*X\s*3"     → "3X3" o "3 X 3"

PROGRAMAS_CONFIG = {

    "3X3": {
        "patrones_s3": [r"3\s*X\s*3"],
        "youtube": False,
    },

    "Tu Día": {
        "patrones_s3": [r"TU\s+D[IÍ]A"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "patrones_youtube": [r"tu\s+d[ií]a", r"tu dia"],
    },

    "Teletrece": {
        "patrones_s3": [r"^TELETRECE$", r"TELETRECE\s+TARDE"],
        "youtube": True,
        "canal_youtube": "noticias",
        "patrones_youtube": [r"teletrece", r"t13 central"],
    },

    "El Tiempo": {
        "patrones_s3": [r"^EL TIEMPO$"],
        "youtube": False,
    },

    "La Tarde es Nuestra": {
        "patrones_s3": [r"LA TARDE ES NUESTRA", r"TARDE ES NUESTRA"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "patrones_youtube": [r"la tarde es nuestra", r"tarde es nuestra"],
    },

    "Hay que Decirlo": {
        "patrones_s3": [r"HAY QUE DECIRLO"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "patrones_youtube": [r"hay que decirlo"],
    },

    "Que Dice Chile": {
        "patrones_s3": [r"QUE DICE CHILE"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "patrones_youtube": [r"que dice chile"],
    },

    "Vecinos al Límite": {
        "patrones_s3": [r"VECINOS AL L[IÍ]MITE"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "patrones_youtube": [r"vecinos al l[ií]mite"],
    },

    "Teletrece Noche": {
        "patrones_s3": [r"TELETRECE NOCHE"],
        "youtube": False,
    },

    "T13 en Vivo": {
        "patrones_s3": [r"T13 EN VIVO"],
        "youtube": False,
    },

    "Mejor Tarde que Nunca": {
        "patrones_s3": [r"MEJOR TARDE QUE NUNCA"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "playlist_youtube": "PLnDONcPxnlq0gsDqssXR_4qLJw6nZxr0P",
    },

    "Las Milf": {
        "patrones_s3": [r"MILF"],
        "youtube": True,
        "canal_youtube": "entretenimiento",
        "playlist_youtube": "PLnDONcPxnlq23hL0qoAm45AILXZzzxVP-",
    },

}
