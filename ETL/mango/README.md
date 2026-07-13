# ETL / Mango — Scripts de consulta a API Mango

Carpeta con scripts para extraer datos de audiencia digital desde la API Mango
y combinarlos con los datos de programación S3 de Canal 13.

---

## 1. `consulta_mango.py` — Consulta libre por señal y horario

### ¿Qué hace?
Consulta la API Mango para **cualquier señal** en un **rango de fechas y horario definido manualmente**.
Exporta los resultados a un Excel con Plays, Streams, Devices y Avg Time por día.

### ¿Cuándo usarlo?
Cuando ya sabes el horario exacto que quieres consultar y no necesitas cruzarlo con datos de S3.
Por ejemplo: "quiero ver la señal T13 todos los días de una semana entre 22:00 y 23:00".

### Cómo ejecutarlo
```
python ETL/mango/consulta_mango.py
```

### Preguntas que hace
| Pregunta | Ejemplo |
|----------|---------|
| Señal (lista o slug manual) | `1` → T13 |
| Fecha DESDE | `2026-06-11` |
| Fecha HASTA | `2026-06-16` |
| ¿Mismo horario para todos los días? | `s` |
| Hora INICIO | `22:25` |
| Hora FIN | `22:50` |

Si respondes `n` a "mismo horario", te pide inicio y fin para cada día por separado.

### Salida
Excel: `Mango_{SEÑAL}_{desde}_{hasta}.xlsx` en `C:\procesos\ReporteCanal13\datos`

| Columna | Descripción |
|---------|-------------|
| Fecha | Fecha del día |
| Día | Lun / Mar / ... |
| Inicio | Hora de inicio consultada |
| Fin | Hora de fin consultada |
| Duración | Minutos del bloque |
| Plays | Reproducciones totales en Mango |
| Streams | Streams simultáneos |
| Devices | Dispositivos únicos |
| Avg Time | Tiempo promedio de visionado |

---

## 2. `s3_vs_mango.py` — S3 vs Mango por programa

### ¿Qué hace?
Combina **dos fuentes de datos** para un programa específico:
- **S3**: lee los archivos `*_programacion.xlsx` ya descargados y extrae automáticamente el horario real (Inicial / Final) y los plays de cada día.
- **Mango**: consulta la API usando esos horarios reales y obtiene las métricas de la señal elegida.

Exporta un Excel comparativo con ambas fuentes lado a lado.

### ¿Cuándo usarlo?
Cuando quieres cruzar los datos de S3 con la señal de Mango para un programa concreto,
sin tener que buscar manualmente los horarios. El script los detecta solo desde los archivos S3.

### Cómo ejecutarlo
```
python ETL/mango/s3_vs_mango.py
```

### Preguntas que hace
| Pregunta | Ejemplo |
|----------|---------|
| Fecha DESDE | `2026-06-11` |
| Fecha HASTA | `2026-06-16` |
| Nombre o parte del programa | `D13` → encuentra `D13 MUNDIAL` |
| Si hay varias coincidencias, elige número | `1` |
| Señal Mango (lista o slug manual) | `1` → T13 |

### Salida
Excel: `S3_vs_{SEÑAL}_{PROGRAMA}_{desde}_{hasta}.xlsx` en `C:\procesos\ReporteCanal13\datos`

| Columna | Descripción |
|---------|-------------|
| Fecha | Fecha del día |
| Día | Lun / Mar / ... |
| Inicio | Hora de inicio tomada del archivo S3 |
| Fin | Hora de fin tomada del archivo S3 |
| Duración | Minutos del programa ese día |
| Plays S3 | Plays registrados en el archivo de programación |
| Plays Mango | Plays de la señal en ese bloque horario |
| Streams Mango | Streams simultáneos en Mango |
| Devices Mango | Dispositivos únicos en Mango |

---

## Señales conocidas

| Opción | Slug |
|--------|------|
| 1 | T13 |
| 2 | C13 |
| 3 | CNN_CHILE |
| 4 | T13R |
| 0 | Ingresar manualmente |

Si Mango incorpora nuevas señales, agrégalas en el diccionario `SENALES_CONOCIDAS`
al inicio de cada script.

---

## Requisitos
```
pip install requests pandas openpyxl
```

Los archivos S3 deben estar descargados previamente con `descargar_s3.py`.
