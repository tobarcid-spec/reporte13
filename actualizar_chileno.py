"""
actualizar_chileno.py — Informe "Siempre Hay un Chileno" (en vivo, S3)
=====================================================================
Pensado para correr los JUEVES después de mediodía (cuando el archivo de S3 del
miércoles ya está disponible). Descarga el archivo del día anterior (miércoles),
y regenera Siempre_hay_un_Chileno_en_vivo_S3.xlsx con todas las emisiones en vivo
de los miércoles desde el 2026-10-14 hasta ayer.

Uso:
    python actualizar_chileno.py                # ayer (debe ser miércoles)
    python actualizar_chileno.py 2026-11-04     # un miércoles específico
"""
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import descargar_s3 as s3

AQUI = Path(__file__).resolve().parent
SALIDA = AQUI / "Siempre_hay_un_Chileno_en_vivo_S3.xlsx"
PRIMER_EN_VIVO = date(2026, 10, 14)   # primera emisión en vivo (miércoles)
PROGRAMA = "SIEMPRE HAY UN CHILENO"


def descargar_dia(dia: date):
    """Descarga el archivo del día, reemplazando el existente (por si estaba incompleto)."""
    destino = s3.CARPETA_DESTINO / f"{dia}_programacion.xlsx"
    respaldo = destino.with_suffix(".xlsx.bak")
    if destino.exists():
        destino.replace(respaldo)
    try:
        s3.descargar_rango(dia, dia)
    finally:
        if respaldo.exists():
            if destino.exists():
                respaldo.unlink()          # descarga nueva OK
            else:
                respaldo.replace(destino)  # falló: se conserva el archivo anterior
                print(f"⚠️  No se pudo descargar {dia}; se conserva el archivo anterior.")


def emisiones_en_vivo(hasta: date):
    filas = []
    d = PRIMER_EN_VIVO
    while d <= hasta:
        ruta = s3.CARPETA_DESTINO / f"{d}_programacion.xlsx"
        if ruta.exists():
            df = pd.read_excel(ruta)
            df.columns = ["ini", "fin", "tit", "plays"] + list(df.columns[4:])
            t = df.tit.astype(str).str.upper()
            m = df[t.str.contains(PROGRAMA) & ~t.str.contains(r"\(R\)")]
            for _, r in m.iterrows():
                filas.append((pd.Timestamp(r.ini), pd.Timestamp(r.fin), int(r.plays)))
        else:
            print(f"⚠️  Falta el archivo de {d}")
        d += timedelta(days=7)
    return sorted(filas)


def escribir_excel(filas, hasta: date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Siempre hay un Chileno"
    ws["A1"] = "Siempre Hay un Chileno — Visualizaciones EN VIVO (S3)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Corresponde al en vivo del programa (excluye repetición). "
                f"Emisiones de los miércoles, hasta el {hasta:%d-%m-%Y}.")
    ws["A2"].font = Font(italic=True, color="555555")
    thin = Side(style="thin", color="BBBBBB")
    borde = Border(top=thin, bottom=thin, left=thin, right=thin)
    for i, h in enumerate(["Fecha", "Día", "Inicio", "Fin", "Plays en vivo (S3)"], 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    for j, (a, b, p) in enumerate(filas, 5):
        for i, v in enumerate([a.date(), "Miércoles", f"{a:%H:%M}", f"{b:%H:%M}", p], 1):
            c = ws.cell(row=j, column=i, value=v)
            c.border = borde
            c.alignment = Alignment(horizontal="center")
        ws.cell(row=j, column=1).number_format = "DD-MM-YYYY"
        ws.cell(row=j, column=5).number_format = "#,##0"
    n = 4 + len(filas)
    for k, (lbl, fn) in enumerate([("Total", "SUM"), ("Promedio", "AVERAGE")], 1):
        ws.cell(row=n + k, column=4, value=lbl).font = Font(bold=True)
        c = ws.cell(row=n + k, column=5, value=f"={fn}(E5:E{n})")
        c.number_format = "#,##0"
        c.font = Font(bold=True)
    for col, w in zip("ABCDE", [14, 12, 10, 10, 22]):
        ws.column_dimensions[col].width = w
    wb.save(SALIDA)


def main():
    if len(sys.argv) > 1:
        dia = date.fromisoformat(sys.argv[1])
    else:
        dia = date.today() - timedelta(days=1)
    if dia.weekday() != 2:
        print(f"{dia} no es miércoles; nada que actualizar.")
        return
    descargar_dia(dia)
    filas = emisiones_en_vivo(dia)
    try:
        escribir_excel(filas, dia)
    except PermissionError:
        print(f"ERROR: cierra {SALIDA.name} (está abierto en Excel) y vuelve a ejecutar.")
        sys.exit(1)
    print(f"\n✅ {SALIDA.name} actualizado: {len(filas)} emisiones, última {filas[-1][0]:%d-%m-%Y}"
          f" con {filas[-1][2]:,} plays")


if __name__ == "__main__":
    main()
