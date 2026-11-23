"""Informe YouTube — React El Clan (temporada 2026): vistas por capítulo, promedio y retención.
Vistas/likes/comentarios/duración: API YouTube. Retención: manual (Studio) — se conserva al regenerar.
Uso: python informe_clan.py"""
import re
from pathlib import Path
from googleapiclient.discovery import build
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from config_local import YOUTUBE_API_KEY, PROGRAMAS_CONFIG

SALIDA = Path(__file__).resolve().parent / "Informe_React_El_Clan_YouTube.xlsx"
cfg = PROGRAMAS_CONFIG["El Clan React"]
pat = re.compile(r"React El Clan", re.I)

yt = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
ids, tok = {}, None
while True:
    r = yt.playlistItems().list(part="snippet", playlistId=cfg["playlist_youtube"], maxResults=50, pageToken=tok).execute()
    for it in r["items"]:
        t = it["snippet"]["title"]
        m = re.search(r"cap[ií]tulo\s*(\d+)", t, re.I)
        if m and pat.search(t):
            ids[int(m.group(1))] = it["snippet"]["resourceId"]["videoId"]
    tok = r.get("nextPageToken")
    if not tok:
        break
inv = {v: k for k, v in ids.items()}
datos = {}
for v in yt.videos().list(part="snippet,statistics,contentDetails", id=",".join(inv)).execute()["items"]:
    h, mi, s = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", v["contentDetails"]["duration"]).groups()
    seg = int(h or 0) * 3600 + int(mi or 0) * 60 + int(s or 0)
    st = v["statistics"]
    datos[inv[v["id"]]] = dict(
        fecha=v["snippet"]["publishedAt"][:10], titulo=v["snippet"]["title"], seg=seg,
        vistas=int(st.get("viewCount", 0)), likes=int(st.get("likeCount", 0)),
        coment=int(st.get("commentCount", 0)))

# retención manual previa: {cap: (min_vistos "mm:ss" texto/num, %)}
previo = {}
if SALIDA.exists():
    ws0 = load_workbook(SALIDA).active
    for r in range(5, ws0.max_row + 1):
        c = ws0.cell(r, 1).value
        if isinstance(c, int):
            previo[c] = (ws0.cell(r, 7).value, ws0.cell(r, 8).value)

wb = Workbook()
ws = wb.active
ws.title = "React El Clan"
AZUL = PatternFill("solid", fgColor="156082"); CEL = PatternFill("solid", fgColor="BDE3F2")
fino = Side(style="thin", color="000000"); B = Border(top=fino, bottom=fino, left=fino, right=fino)
ws["A1"] = "React El Clan — Temporada 2026 (YouTube)"; ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Vistas, likes y comentarios: API YouTube al día de generación. Retención: ingresar desde YouTube Studio (celdas en azul)."
ws["A2"].font = Font(italic=True, color="555555")
heads = ["Cap.", "Fecha", "Título", "Duración", "Vistas", "Likes", "Tiempo prom. visto (h:mm:ss)", "% retención prom.", "Comentarios"]
for i, h in enumerate(heads, 1):
    c = ws.cell(4, i, h); c.fill = AZUL; c.font = Font(bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = B
ws.row_dimensions[4].height = 34
r = 5
for n in sorted(datos):
    d = datos[n]
    tit = re.sub(r"^React El Clan\s*\|\s*Cap[ií]tulo\s*\d+\s*\|?\s*", "", d["titulo"]) or "(sin subtítulo)"
    mv, pr = previo.get(n, (None, None))
    vals = [n, d["fecha"], tit, d["seg"] / 86400, d["vistas"], d["likes"], mv, pr, d["coment"]]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v); cell.border = B
        if c in (1, 2): cell.alignment = Alignment(horizontal="center")
    ws.cell(r, 4).number_format = "h:mm:ss"
    for c in (5, 6, 9): ws.cell(r, c).number_format = "#,##0"
    ws.cell(r, 7).number_format = "h:mm:ss"; ws.cell(r, 8).number_format = "0.0%"
    for c in (7, 8): ws.cell(r, c).font = Font(color="0000FF")
    r += 1
ini, fin = 5, r - 1
for k, (lbl, fn) in enumerate([("Promedio", "AVERAGE"), ("Mediana", "MEDIAN"), ("Máximo", "MAX"), ("Mínimo", "MIN"), ("Total", "SUM")]):
    rr = fin + 1 + k
    for c in range(1, 10):
        ws.cell(rr, c).fill = CEL; ws.cell(rr, c).border = B; ws.cell(rr, c).font = Font(bold=True)
    ws.cell(rr, 3, lbl)
    cols = (4, 5, 6, 7, 8, 9) if fn != "SUM" else (5, 6, 9)
    for c in cols:
        L = ws.cell(1, c).column_letter
        ws.cell(rr, c, f'=IFERROR({fn}({L}{ini}:{L}{fin}),"")')
        ws.cell(rr, c).number_format = ws.cell(ini, c).number_format if c != 5 or True else "#,##0"
    ws.cell(rr, 4).number_format = "h:mm:ss"
for col, w in zip("ABCDEFGHI", [6, 12, 50, 11, 12, 10, 18, 14, 13]):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"
wb.save(SALIDA)
print("OK", len(datos), "capítulos:", sorted(datos), SALIDA)
