# Analisis de tiempos operativos (Disponibilidad y UEBD)

Este script procesa archivos de eventos de perforadoras (CSV con separador `;`) y genera indicadores en **horas** para:

- Efectivo
- Reserva
- Mantencion Programada
- Mantencion No Programada

Ademas calcula:

- **Horas Operativas**
- **Disponibilidad** (ratio y %)
- **UEBD** (ratio y %)

---

## 1) Logica de turnos y dia operativo

Se considera que:

- **Turno A** inicia a las **21:00**
- **Turno B** inicia a las **09:00**
- Un dia operativo = Turno A + Turno B

Para asignar cada registro al dia operativo:

1. Se usa la columna `WorkDayStarted` si existe.
2. Si no existe, se calcula como `(Time - 21 horas).date()`.

---

## 2) Clasificacion de tiempos

Usa estas columnas:

- `ShortCode`
- `PlannedCodeName`
- `OnlyCodeName` (apoyo para Efectivo)

Reglas:

- `ShortCode = Efectivo` -> `horas_efectivo`
- `ShortCode = Reserva` -> `horas_reserva`
- `ShortCode = Mantencion` y `PlannedCodeName = Programada` -> `horas_mant_programada`
- `ShortCode = Mantencion` y otro valor -> `horas_mant_no_programada`
- Cualquier otro caso -> `horas_otras`

---

## 3) Formulas usadas

- `horas_operativas = horas_totales - horas_mant_programada - horas_mant_no_programada`
- `disponibilidad_ratio = horas_operativas / horas_totales`
- `uebd_ratio = horas_efectivo / horas_operativas`
- `disponibilidad_pct = disponibilidad_ratio * 100`
- `uebd_pct = uebd_ratio * 100`
- **Formula solicitada por usuario**:
  - `disponibilidad_formula_usuario = (horas_operativas / horas_totales) / 100`
  - `uebd_formula_usuario = (horas_efectivo / horas_operativas) / 100`

> Si el denominador es 0, el porcentaje se reporta como 0.
> Se incluye tambien `horas_disponibles` como alias de compatibilidad con el mismo valor de `horas_operativas`.

---

## 4) Ejecutar

Desde la carpeta del proyecto:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100.csv"
```

Opcionalmente puedes definir carpeta de salida:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100.csv" --output-dir "salidas_analisis"
```

---

## 5) Archivos de salida

En la carpeta de salida se generan:

1. `diario_por_perforadora.csv`
   - 1 fila por `fecha_operativa + perforadora`
   - horas por componente + disponibilidad + UEBD

2. `mensual_por_perforadora.csv`
   - resumen por `anio + mes + perforadora`
   - incluye **promedio diario** de cada componente

3. `anual_por_perforadora.csv`
   - resumen por `anio + perforadora`
   - incluye **promedio diario**

4. `mensual_flota.csv`
   - resumen mensual total de flota

5. `anual_flota.csv`
   - resumen anual total de flota

---

## 6) Encabezados esperados en el CSV

Columnas minimas requeridas:

- `RigName`
- `Time`
- `EndTime`
- `Duration`
- `ShortCode`
- `OnlyCodeName`
- `PlannedCodeName`

Si faltan, el script devuelve error indicando cuales faltan.
