"""
corregir_titulos_finde.py
==========================
Los archivos *_programacion.xlsx de sábado y domingo a veces traen títulos
genéricos (ej. "CULTURA 13 PM", "CULTURA 13 ACCESS", "CULTURA 13 PRIME",
"SABADOS DE PELICULAS", "DOMINGOS DE PELICULA") en vez del nombre real del
programa que fue al aire esa franja.

Este script corrige esos títulos genéricos usando la grilla real de
programación de ese día (hora de inicio + nombre del programa).

Uso:
    python corregir_titulos_finde.py 2026-10-24 horario_sabado.txt
    python corregir_titulos_finde.py 2026-10-25 horario_domingo.txt

El archivo de horario debe tener una línea por programa, con la hora de
inicio (HH:MM) y el nombre real separados por tabulador o por el primer
espacio, en el orden en que van al aire. Las horas menores a 06:00 se
interpretan como la madrugada del día SIGUIENTE (igual que el archivo C13,
cuyo "día de transmisión" va de las 06:00 a las 06:00 del día después).

Ejemplo de archivo de horario:
    6:35\tTravesia 13C: La Gracia de Viajar (R)
    8:00\tT13 Finde
    13:00\tTeletrece Tarde
    15:00\tSábado de Películas: "Spiderman"
    21:00\tTeletrece
    0:15\tEl Clan (R)
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import re
import pandas as pd
from datetime import datetime, timedelta

CARPETA_DATOS = r"C:\procesos\ReporteCanal13\datos"

# Patrones de títulos genéricos que se deben reemplazar por el nombre real.
PATRONES_GENERICOS = [
    re.compile(r"^CULTURA\s*13", re.IGNORECASE),
    re.compile(r"^S[AÁ]BADOS?\s+DE\s+PEL[IÍ]CULAS?", re.IGNORECASE),
    re.compile(r"^DOMINGOS?\s+DE\s+PEL[IÍ]CULAS?", re.IGNORECASE),
    re.compile(r"^TRAVES[IÍ]A\s*13\s*C?\s*$", re.IGNORECASE),
]


def es_generico(titulo: str) -> bool:
    titulo = str(titulo).strip()
    return any(p.match(titulo) for p in PATRONES_GENERICOS)


def parsear_horario(ruta_txt: str, fecha_base: str):
    """Lee el archivo de horario y retorna lista de (datetime, programa) ordenada."""
    entradas = []
    fecha_dt = datetime.strptime(fecha_base, "%Y-%m-%d")

    with open(ruta_txt, encoding="utf-8") as f:
        for linea in f:
            linea = linea.rstrip("\n")
            if not linea.strip():
                continue
            if "\t" in linea:
                hora, programa = linea.split("\t", 1)
            else:
                hora, programa = linea.split(" ", 1)
            hora = hora.strip()
            programa = programa.strip()

            h, m = map(int, hora.split(":"))
            dt = fecha_dt.replace(hour=0, minute=0) + timedelta(hours=h, minutes=m)
            if h < 6:
                dt += timedelta(days=1)  # madrugada del día siguiente
            entradas.append((dt, programa))

    entradas.sort(key=lambda x: x[0])
    return entradas


def buscar_programa_real(horario, momento):
    """Retorna el programa de 'horario' cuyo inicio es el más reciente <= momento (floor match).

    Si 'momento' es anterior a toda la grilla (por redondeos de 1-2 min entre el
    horario real y el archivo C13), se usa la primera entrada de la grilla."""
    if momento < horario[0][0] and (horario[0][0] - momento) <= timedelta(minutes=5):
        return horario[0][1]
    candidato = None
    for dt, programa in horario:
        if dt <= momento:
            candidato = programa
        else:
            break
    return candidato


def corregir_archivo(fecha: str, ruta_horario: str):
    ruta_xlsx = f"{CARPETA_DATOS}\\{fecha}_programacion.xlsx"
    df = pd.read_excel(ruta_xlsx)
    tcol = df.columns[2]

    horario = parsear_horario(ruta_horario, fecha)
    if not horario:
        print("⚠️  El archivo de horario está vacío.")
        return

    cambios = 0
    for idx, row in df.iterrows():
        titulo = str(row[tcol]).strip()
        if titulo.upper() == "TOTAL" or not es_generico(titulo):
            continue
        inicial = row.get("Inicial")
        if pd.isna(inicial):
            continue
        momento = pd.to_datetime(inicial).to_pydatetime()
        real = buscar_programa_real(horario, momento)
        if real and real != titulo:
            print(f"  {momento}  '{titulo}' -> '{real}'")
            df.at[idx, tcol] = real
            cambios += 1

    if cambios:
        df.to_excel(ruta_xlsx, index=False)
        print(f"\n✅ {cambios} título(s) corregido(s) en {ruta_xlsx}")
    else:
        print("\nSin títulos genéricos para corregir (o ya estaban corregidos).")


def main():
    if len(sys.argv) != 3:
        print("Uso: python corregir_titulos_finde.py YYYY-MM-DD archivo_horario.txt")
        sys.exit(1)
    fecha = sys.argv[1]
    ruta_horario = sys.argv[2]
    corregir_archivo(fecha, ruta_horario)


if __name__ == "__main__":
    main()
