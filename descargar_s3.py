"""
descargar_s3.py — Descarga archivos de programación desde S3
=============================================================
Uso:
    python descargar_s3.py                          # te pregunta las fechas
    python descargar_s3.py 2026-04-01 2026-04-30   # rango por argumento

Requiere:
    pip install boto3
    AWS configurado: aws configure  (o variables AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY)
"""

import sys
import re
import boto3
from datetime import date, timedelta
from pathlib import Path
from botocore.exceptions import ClientError, NoCredentialsError

# ─── Configuración ────────────────────────────────────────────────────────────
BUCKET      = "datos-dps"
REGION      = "us-east-2"
CARPETA_DESTINO = Path(r"C:\procesos\ReporteCanal13\datos")

# Patrón de carpetas dentro del bucket: YYYY-MM-DD/
PATRON_CARPETA = re.compile(r"^(\d{4}-\d{2}-\d{2})/$")
# Patrón del archivo dentro de cada carpeta: YYYY-MM-DD_programacion.*
PATRON_ARCHIVO = re.compile(r"^\d{4}-\d{2}-\d{2}_programacion\.", re.IGNORECASE)


def pedir_fecha(mensaje: str) -> date:
    while True:
        texto = input(mensaje).strip()
        try:
            return date.fromisoformat(texto)
        except ValueError:
            print("  Formato incorrecto. Usa YYYY-MM-DD (ej. 2026-04-01)")


def fechas_en_rango(desde: date, hasta: date):
    d = desde
    while d <= hasta:
        yield d
        d += timedelta(days=1)


def descargar_rango(desde: date, hasta: date):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)

    try:
        s3 = boto3.client("s3", region_name=REGION)
    except NoCredentialsError:
        print("ERROR: No se encontraron credenciales AWS.")
        print("Ejecuta 'aws configure' o define AWS_ACCESS_KEY_ID y AWS_SECRET_ACCESS_KEY.")
        sys.exit(1)

    # Construir conjunto de fechas solicitadas como strings "YYYY-MM-DD"
    fechas = {str(f) for f in fechas_en_rango(desde, hasta)}

    print(f"\nBuscando archivos del {desde} al {hasta} en s3://{BUCKET}/...")

    # Listar todas las carpetas del bucket (prefijos de nivel 1)
    paginator = s3.get_paginator("list_objects_v2")
    pages = paginator.paginate(Bucket=BUCKET, Delimiter="/")

    carpetas_encontradas = []
    for page in pages:
        for prefix_obj in page.get("CommonPrefixes", []):
            prefix = prefix_obj["Prefix"]          # ej. "2026-04-01/"
            m = PATRON_CARPETA.match(prefix)
            if m and m.group(1) in fechas:
                carpetas_encontradas.append((m.group(1), prefix))

    if not carpetas_encontradas:
        print("No se encontraron carpetas para el rango indicado.")
        return

    carpetas_encontradas.sort()
    print(f"Carpetas encontradas: {len(carpetas_encontradas)}\n")

    descargados = 0
    omitidos    = 0
    errores     = 0

    for fecha_str, prefix in carpetas_encontradas:
        # Listar archivos dentro de la carpeta
        resp = s3.list_objects_v2(Bucket=BUCKET, Prefix=prefix)
        objetos = resp.get("Contents", [])

        archivos_prog = [
            o["Key"] for o in objetos
            if PATRON_ARCHIVO.match(Path(o["Key"]).name)
        ]

        if not archivos_prog:
            print(f"  [{fecha_str}] Sin archivo _programacion — omitido")
            omitidos += 1
            continue

        for key in archivos_prog:
            nombre_archivo = Path(key).name
            destino = CARPETA_DESTINO / nombre_archivo

            if destino.exists():
                print(f"  [{fecha_str}] Ya existe: {nombre_archivo} — omitido")
                omitidos += 1
                continue

            try:
                print(f"  [{fecha_str}] Descargando {nombre_archivo}...", end=" ", flush=True)
                s3.download_file(BUCKET, key, str(destino))
                print("OK")
                descargados += 1
            except ClientError as e:
                print(f"ERROR: {e}")
                errores += 1

    print(f"\nResumen: {descargados} descargados, {omitidos} omitidos, {errores} errores.")
    print(f"Destino: {CARPETA_DESTINO}")


def main():
    args = sys.argv[1:]

    if len(args) == 2:
        try:
            desde = date.fromisoformat(args[0])
            hasta = date.fromisoformat(args[1])
        except ValueError:
            print("Fechas inválidas. Usa formato YYYY-MM-DD.")
            sys.exit(1)
    elif len(args) == 0:
        print("=== Descarga de archivos S3 — Canal 13 ===\n")
        desde = pedir_fecha("Fecha DESDE (YYYY-MM-DD): ")
        hasta = pedir_fecha("Fecha HASTA  (YYYY-MM-DD): ")
    else:
        print("Uso: python descargar_s3.py [FECHA_DESDE FECHA_HASTA]")
        sys.exit(1)

    if desde > hasta:
        print("La fecha DESDE debe ser anterior o igual a HASTA.")
        sys.exit(1)

    descargar_rango(desde, hasta)


if __name__ == "__main__":
    main()
