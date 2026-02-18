# Calculo de metricas operativas

Este repositorio incluye un script para calcular las metricas solicitadas desde el archivo base:

- `DispUEBD_AllRigs_010126-0000_170226-2100.csv`

## Formulas implementadas

- `UEBD = (Horas Efectivas / Horas Operativas) / 100`
- `Disponibilidad = (Horas Operativas / Horas Totales) / 100`

## Reglas usadas para construir horas

- `Horas Efectivas`: suma de `Duration` cuando `ShortCode == Efectivo`.
- `Horas Operativas`: suma de `Duration` cuando `ShortCode` es `Efectivo` o `Demora`.
- `Horas Totales`: suma de `Duration` para todos los registros.

> `Duration` se interpreta en segundos y se convierte a horas.

## Ejecucion

```bash
python3 calculate_operational_metrics.py
```

Salida por defecto:

- `metricas_operativas_por_rig.csv` (agrupado por `RigName`)

## Opciones utiles

```bash
python3 calculate_operational_metrics.py \
  --group-by RigName ShiftName WorkDayStarted \
  -o metricas_operativas_por_rig_turno_dia.csv
```

