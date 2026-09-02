"""
config_local.py — Configuración del Reporte Local Canal 13
===========================================================
Edita este archivo antes de ejecutar reporte_desde_carpeta.py
"""

# ─── Carpeta donde guardas los archivos descargados de S3 ────────────────────
# Windows:  r"C:\Users\TuNombre\ReporteCanal13\datos"
# Mac/Linux: "/Users/tunombre/ReporteCanal13/datos"
CARPETA_DATOS = r"C:\procesos\ReporteCanal13\datos"

# ─── YouTube API ─────────────────────────────────────────────────────────────
YOUTUBE_API_KEY = "AIzaSyDc75yPe--BM4npEhN5Yak3xdUnbKh_0Jc"

# ─── Base de datos MySQL ──────────────────────────────────────────────────────
MYSQL_CONFIG = {
    "host":     "217.160.158.217",
    "port":     3306,
    "user":     "user_md_new",
    "password": "md_secuo_c13.$2025",
    "database": "MEDIOS_DIGITALES",
}
MYSQL_TABLA = "plays_diarios"

# ─── Programas ───────────────────────────────────────────────────────────────
# patrones_s3:      regex para matchear el título en programacion.xlsx
# playlist_youtube: ID de la playlist en YouTube (si aplica) para contar views totales
#
# Los patrones usan regex. Ejemplos:
#   r"TU DIA"        → match exacto
#   r"TU.?D[IÍ]A"   → match flexible con o sin acento

PROGRAMAS_CONFIG = {

    "3X3": {
        "patrones_s3": [r"^3\s*X\s*3$"],
        "playlist_youtube": None,
    },

    "Tu Día": {
        "patrones_s3": [r"^TU\s+DIA$"],
        "playlist_youtube": "PLnDONcPxnlq2Iy8sP5JFPjT5am2Cmmiln",
    },

    "Teletrece A.M.": {
        "patrones_s3": [r"TELETRECE\s+A\.?M\.?"],
        "playlist_youtube": None,
    },

    "Teletrece Tarde": {
        "patrones_s3": [r"^TELETRECE\s+TARDE$"],
        "playlist_youtube": None,  # Agregar si existe
    },

    "Teletrece": {
        "patrones_s3": [r"^TELETRECE$"],
        "playlist_youtube": None,  # Agregar si existe
    },

    "Teletrece Noche": {
        "patrones_s3": [r"^TELETRECE\s+NOCHE$"],
        "playlist_youtube": None,
    },

    "El Tiempo": {
        "patrones_s3": [r"^EL\s+TIEMPO$"],
        "playlist_youtube": None,
    },

    "El Tiempo (Tarde)": {
        "patrones_s3": [r"EL\s+TIEMPO\s*\(TARDE\)"],
        "playlist_youtube": None,
    },

    "La Tarde es Nuestra": {
        "patrones_s3": [r"^LA\s+TARDE\s+ES\s+NUESTRA$"],
        "playlist_youtube": "PLnDONcPxnlq3ijFAfv9O1nQP4ub46eUGc",
    },

    "Hay que Decirlo": {
        "patrones_s3": [r"^HAY\s+QUE\s+DECIRLO$"],
        "playlist_youtube": "PLnDONcPxnlq0yjE--VHMb8HxwUwoz1ILe",
    },

    "Hay que Decirlo (R)": {
        "patrones_s3": [r"HAY\s+QUE\s+DECIRLO\s*\(R\)"],
        "playlist_youtube": None,
    },

    "Que Dice Chile": {
        "patrones_s3": [r"^QUE\s+DICE\s+CHILE$"],
        "playlist_youtube": None,  # Agregar si existe
    },

    "Vecinos al Límite": {
        "patrones_s3": [r"^VECINOS\s+AL\s+LIMITE$",
                        r"VECINOS\s+AL\s+LIMITE\s*\(ESPECIAL\)",
                        r"VECINOS\s+AL\s+LIMITE\s*\(RESUMEN\)"],
        "playlist_youtube": None,
    },

    "Vecinos al Límite Extra": {
        "patrones_s3": [r"VECINOS\s+AL\s+LIMITE\s+EXTRA"],
        "playlist_youtube": None,
    },

    "Vecinos al Límite React": {
        "patrones_s3": None,
        "playlist_youtube": "PLnDONcPxnlq0Q4S9NJnWq9pcWC1MXXaYp",
    },

    # "El Clan React" y "Socios por Chile React" comparten la misma playlist de YouTube.
    # Se distinguen por filtro_titulo. El video se sube un día después de emitirse en TV,
    # por eso dia_fijo remapea la fecha al día real de emisión (0=Lunes ... 6=Domingo).
    "El Clan React": {
        "patrones_s3": None,
        "playlist_youtube": "PLnDONcPxnlq3aD54zvbw9fGmxQO63DjBc",
        "filtro_titulo": r"El Clan",
        "dia_fijo": 5,  # Sábado
    },

    "Socios por Chile React": {
        "patrones_s3": None,
        "playlist_youtube": "PLnDONcPxnlq3aD54zvbw9fGmxQO63DjBc",
        "filtro_titulo": r"Socios por Chile",
        "dia_fijo": 6,  # Domingo
    },

    "T13 en Vivo": {
        "patrones_s3": [r"^T13\s+EN\s+VIVO$"],
        "playlist_youtube": None,
    },

    "Mejor Tarde que Nunca": {
        "patrones_s3": [r"MEJOR\s+TARDE\s+QUE\s+NUNCA"],
        "playlist_youtube": "PLnDONcPxnlq0gsDqssXR_4qLJw6nZxr0P",
    },

    "Las Milf": {
        "patrones_s3": [r"MILF"],
        "playlist_youtube": "PLnDONcPxnlq23hL0qoAm45AILXZzzxVP-",
    },

    "Te Cuento Todo": {
        "patrones_s3": [r"^TE\s+CUENTO\s+TODO$"],
        "playlist_youtube": "PLMB2psPeunE4",
        "hora_corte": 6,  # se sube de madrugada (~03-04 AM Chile) al día siguiente de la emisión
    },

    "Jet Set": {
        "patrones_s3": [r"^JET\s+SET$"],
        "playlist_youtube": "PLRrtFqi6kkvI",
        "hora_corte": 6,  # a veces se sube de madrugada al día siguiente, a veces el mismo día
    },

}
