"""
informe_vertigo.py — Informe Vértigo Extremo (lunes) y Vértigo Estelar (jueves)
===============================================================================
Pensado para correr los VIERNES después de mediodía (el S3 del día anterior ya está
disponible). Descarga el archivo del jueves (ayer) y el del lunes, y regenera
Informe_Vertigo_S3_YouTube.xlsx con todos los capítulos.

Uso:
    python informe_vertigo.py                # hoy = referencia (viernes)
    python informe_vertigo.py 2026-11-20     # fecha de referencia específica (el viernes)

- 13Go/C13 En vivo: plays S3 del bloque en vivo (excluye "(LO MEJOR)").
- Youtube Total y duración: API de YouTube (video del capítulo N).
- Youtube En vivo/Repetición y Facebook: cifras manuales; se conservan del Excel
  existente (puedes editarlas a mano y no se pierden al regenerar).
"""
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from googleapiclient.discovery import build
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import descargar_s3 as s3
from actualizar_chileno import descargar_dia
from config_local import YOUTUBE_API_KEY, PROGRAMAS_CONFIG

SALIDA = Path(__file__).resolve().parent / "Informe_Vertigo_S3_YouTube.xlsx"

PROGRAMAS = [
    # hoja, título S3 exacto, 1ª emisión, día semana, prefijo fila, config YouTube, video cap.1 si no está en playlist
    dict(hoja="vertigo extremo", s3="VERTIGO EXTREMO", inicio=date(2026, 10, 19), dow=0,
         prefijo="React Vértigo Extremo", cfg="React Vértigo Extremo", extra={1: "TVrksRw8f_0"}),
    dict(hoja="vertigo estelar", s3="VERTIGO", inicio=date(2026, 10, 15), dow=3,
         prefijo="React Vértigo", cfg="React Vértigo", extra={}),
]


def en_vivo_s3(dia, titulo):
    """Plays del bloque en vivo de `titulo` en el archivo de `dia`, o None si no hay archivo/emisión."""
    ruta = s3.CARPETA_DESTINO / f"{dia}_programacion.xlsx"
    if not ruta.exists():
        return None
    df = pd.read_excel(ruta)
    df.columns = ["ini", "fin", "tit", "plays"] + list(df.columns[4:])
    m = df[df.tit.astype(str).str.upper().str.strip() == titulo]
    return int(m.plays.sum()) if len(m) else None


def videos_youtube(yt, cfg, extra):
    """{n° capítulo: (views, 'h:mm')} desde la playlist (filtrada por título) + videos extra."""
    pat = re.compile(cfg["filtro_titulo"], re.I)
    ids = dict(extra)
    tok = None
    while True:
        r = yt.playlistItems().list(part="snippet", playlistId=cfg["playlist_youtube"],
                                    maxResults=50, pageToken=tok).execute()
        for it in r["items"]:
            sn = it["snippet"]
            m = re.search(r"cap[ií]tulo\s*(\d+)", sn["title"], re.I)
            if m and pat.search(sn["title"]):
                ids[int(m.group(1))] = sn["resourceId"]["videoId"]
        tok = r.get("nextPageToken")
        if not tok:
            break
    out = {}
    inv = {v: k for k, v in ids.items()}
    lista = list(inv)
    for i in range(0, len(lista), 50):
        r = yt.videos().list(part="statistics,contentDetails", id=",".join(lista[i:i + 50])).execute()
        for v in r["items"]:
            h, mi = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?", v["contentDetails"]["duration"]).groups()
            out[inv[v["id"]]] = (int(v["statistics"]["viewCount"]), f"{int(h or 0)}:{int(mi or 0):02d}")
    return out


def manuales_previos():
    """{hoja: {programa: (dur, ye, yr, fe, fr, ft)}} leído del Excel existente."""
    prev = {}
    if not SALIDA.exists():
        return prev
    wb = load_workbook(SALIDA)
    for ws in wb:
        d = {}
        for r in range(5, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if a and str(a).startswith("React"):
                d[a] = tuple(ws.cell(r, c).value for c in (3, 4, 5, 7, 8, 9))
        prev[ws.title] = d
    return prev


def construir_filas(p, yt_vids, hasta, previos):
    filas = []
    n = 0
    while True:
        n += 1
        lunes_o_jueves = p["inicio"] + timedelta(days=7 * (n - 1))
        if lunes_o_jueves > hasta:
            break
        fecha, s3v = lunes_o_jueves, en_vivo_s3(lunes_o_jueves, p["s3"])
        if s3v is None and p["dow"] == 0:      # a veces el lunes se emite pasada la medianoche del martes
            alt = en_vivo_s3(lunes_o_jueves + timedelta(days=1), p["s3"])
            if alt is not None:
                fecha, s3v = lunes_o_jueves + timedelta(days=1), alt
        yv = yt_vids.get(n)
        if s3v is None and yv is None:
            continue
        nombre = f"{p['prefijo']} | Capítulo {n}"
        dur_p, ye, yr, fe, fr, ft = previos.get(nombre, (None,) * 6)
        dur = dur_p or (yv[1] if yv else None)
        filas.append((nombre, fecha, dur, ye, yr, yv[0] if yv else None, fe, fr, ft, s3v))
    return filas


AZUL = PatternFill("solid", fgColor="156082")
CELESTE = PatternFill("solid", fgColor="BDE3F2")
fino = Side(style="thin", color="000000")
BORDE = Border(top=fino, bottom=fino, left=fino, right=fino)
BLANCO = Font(bold=True, color="FFFFFF", size=12)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def hoja(wb, nombre, filas, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = nombre
    for r in range(1, 5):
        for c in range(1, 12):
            ws.cell(r, c).fill = AZUL
            ws.cell(r, c).font = BLANCO
            ws.cell(r, c).alignment = CENTRO
    for rng, txt in [("C1:F3", "Youtube"), ("G1:I3", "Facebook"), ("J1:J3", "13Go/C13"), ("K1:K4", "total")]:
        ws.merge_cells(rng)
        ws[rng.split(":")[0]] = txt
    heads = ["Programa", "Fecha", "Duración", "En vivo", "Repetición", "Total",
             "En vivo", "Repetición", "Total", "En vivo"]
    # Youtube abarca desde Duración (C) hasta Total (F), como en el modelo
    for i, h in enumerate(heads, 1):
        ws.cell(4, i, h)
        ws.cell(4, i).alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
    ws.row_dimensions[4].height = 21
    for r in range(1, 4):
        ws.row_dimensions[r].height = 22
    ini = 5
    for n, f in enumerate(filas):
        r = ini + n
        prog, fecha, dur, ye, yr, yt, fe, fr, ft, s3 = f
        vals = [prog, fecha, dur, ye, yr, yt, fe, fr, ft, s3]
        # Repetición = Total − En vivo (fórmula), mientras exista el En vivo
        if ye is not None and yt is not None:
            vals[4] = f"=F{r}-D{r}"
        elif ye is None:
            vals[4] = None
        if fe is not None and ft is not None:
            vals[7] = f"=I{r}-G{r}"
        elif fe is None:
            vals[7] = None
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDE
            cell.font = Font(size=12)
            if c >= 4:
                cell.number_format = "#,##0"
            if c in (2, 3):
                cell.alignment = Alignment(horizontal="right")
        ws.cell(r, 2).number_format = "DD-MMM"
        k = ws.cell(r, 11, f'=IF(COUNT(F{r},I{r},J{r})=0,"",SUM(F{r},I{r},J{r}))')
        k.number_format = "#,##0"; k.border = BORDE; k.font = Font(size=12)
    fin = ini + len(filas) - 1
    t = fin + 1
    for c in range(1, 11):
        cell = ws.cell(t, c)
        cell.fill = CELESTE; cell.border = BORDE; cell.font = Font(size=12)
        if c >= 4:
            L = cell.column_letter
            cell.value = f"=SUM({L}{ini}:{L}{fin})"
            cell.number_format = "#,##0"
    ws.cell(t, 1, "Total")
    k = ws.cell(t, 11, f"=SUM(K{ini}:K{fin})")
    k.number_format = "#,##0"; k.border = BORDE; k.font = Font(bold=True, size=12)
    notas = [
        "Fuentes: 13Go/C13 En vivo = plays S3 del bloque en vivo (excluye '(LO MEJOR)'). Youtube Total = vistas actuales del video del capítulo (API YouTube).",
        "Youtube En vivo/Repetición y Facebook = cifras ingresadas manualmente (la API no las entrega); pueden ser de una fecha anterior al Total de YouTube.",
        "Total = Youtube Total + Facebook Total + 13Go/C13 En vivo. Celdas vacías = dato aún no disponible.",
    ]
    for i, n in enumerate(notas):
        ws.cell(t + 2 + i, 1, n).font = Font(italic=True, color="555555", size=10)
    for col, w in zip("ABCDEFGHIJK", [38, 11, 11, 12, 13, 12, 12, 13, 12, 12, 13]):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    return ws



def main():
    ref = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else date.today()
    if ref.weekday() != 4:
        print(f"⚠️  {ref} no es viernes; se usa igual como fecha de referencia.")
    jueves, lunes = ref - timedelta(days=1), ref - timedelta(days=4)
    for d in (jueves, lunes):
        print(f"Descargando S3 de {d}...")
        descargar_dia(d)
    yt = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
    previos = manuales_previos()
    wb = Workbook()
    for i, p in enumerate(PROGRAMAS):
        vids = videos_youtube(yt, PROGRAMAS_CONFIG[p["cfg"]], p["extra"])
        filas = construir_filas(p, vids, ref - timedelta(days=1), previos.get(p["hoja"], {}))
        hoja(wb, p["hoja"], filas, primera=(i == 0))
        print(f"  {p['hoja']}: {len(filas)} capítulos")
    try:
        wb.save(SALIDA)
    except PermissionError:
        print(f"ERROR: cierra {SALIDA.name} (está abierto en Excel) y vuelve a ejecutar.")
        sys.exit(1)
    print(f"✅ {SALIDA.name} actualizado")


if __name__ == "__main__":
    main()
