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
- **Impacto por codigo** (para explicar caidas de Disponibilidad y UEBD)
- **Resumen ejecutivo del periodo**
- **Top dias criticos**
- **Reporte Excel extensivo con graficos**

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
python3 analisis_tiempos_operativos.py
```

> Si no pasas argumentos, el script busca automaticamente:
> `DispUEBD_AllRigs_010126-0000_170226-2100` (con o sin `.csv`).

Tambien puedes indicar el archivo explicitamente:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100"
```

> El script acepta el nombre con o sin extension `.csv`.

Opcionalmente puedes definir carpeta de salida:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100" --output-dir "salidas_analisis"
```

Top de codigos para cascada (por defecto 10):

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100" --top-n-codigos 12
```

Top de dias criticos (por defecto 20) y top de codigos por perforadora:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100" --top-dias-criticos 30 --top-n-codigos-rig 8
```

Si quieres omitir PNG y dejar solo CSV:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100" --sin-graficos-cascada
```

Si quieres omitir el Excel extensivo:

```bash
python3 analisis_tiempos_operativos.py "DispUEBD_AllRigs_010126-0000_170226-2100" --sin-excel
```

---

## 5) Archivos de salida

En la carpeta de salida se generan:

1. `diario_por_perforadora.csv`
   - 1 fila por `fecha_operativa + perforadora`
   - horas por componente + disponibilidad + UEBD

2. `diario_por_perforadora_y_turno.csv`
   - 1 fila por `fecha_operativa + perforadora + turno`

3. `mensual_por_perforadora.csv`
   - resumen por `anio + mes + perforadora`
   - incluye **promedio diario** de cada componente

4. `anual_por_perforadora.csv`
   - resumen por `anio + perforadora`
   - incluye **promedio diario**

5. `mensual_flota.csv`
   - resumen mensual total de flota

6. `anual_flota.csv`
   - resumen anual total de flota

7. `mensual_por_perforadora_y_turno.csv`
   - resumen mensual por perforadora y turno (A/B)

8. `resumen_ejecutivo_total_periodo.csv`
   - resumen total por perforadora + fila TOTAL

9. `impacto_codigos_disponibilidad.csv`
   - horas e impacto por codigo que reducen disponibilidad

10. `impacto_codigos_uebd.csv`
   - horas e impacto por codigo que reducen UEBD

11. `impacto_codigos_disponibilidad_por_perforadora.csv`
    - top codigos de disponibilidad por cada perforadora

12. `impacto_codigos_uebd_por_perforadora.csv`
    - top codigos de UEBD por cada perforadora

13. `top_dias_criticos_disponibilidad.csv`
    - dias con menor disponibilidad

14. `top_dias_criticos_uebd.csv`
    - dias con menor UEBD

15. `graficos/cascada_disponibilidad_top_codigos.png`
   - cascada Top N codigos de mayor impacto negativo en disponibilidad

16. `graficos/cascada_uebd_top_codigos.png`
   - cascada Top N codigos de mayor impacto negativo en UEBD

17. `reporte_extensivo_tiempos_operativos.xlsx` (nombre configurable)
    - libro Excel con multiples hojas y graficos:
      - Resumen Ejecutivo
      - Diario / Mensual / Anual
      - Turnos
      - Impactos por codigo
      - Top dias criticos

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

---

## 7) Dependencias para graficos y Excel

Los PNG de cascada usan `matplotlib` y el libro Excel usa `xlsxwriter`.

Instalacion:

```bash
python3 -m pip install matplotlib
```

```bash
python3 -m pip install xlsxwriter
```

Si quieres leer objetivos desde archivos mensuales `.xlsx`, instala tambien:

```bash
python3 -m pip install openpyxl
```

---

## 8) Analisis de brecha 2025 vs Feb-2026 (nuevo)

Se incluye un script adicional:

- `analisis_brecha_2025_vs_feb2026.py`

Este script hace exactamente el enfoque de brecha:

1. Baseline = promedio diario 2025 (toda la data 2025)
2. Comparado = promedio diario Feb-2026
3. Gap de UEBD y Disponibilidad vs objetivo mensual
4. Atribucion del gap a codigos por delta de horas/dia
5. Escalamiento para que la suma de impactos por codigo cierre exactamente el gap (pp)
6. Perdidas separadas por:
   - Disponibilidad
   - UEBD (por codigo)
   - Rendimiento F09 (filtrado por `DrillPlan` contiene `F09`)
   - Malla (manual, via parametro)

### Ejecucion minima (con nombres por defecto)

```bash
python3 analisis_brecha_2025_vs_feb2026.py
```

Por defecto busca:

- `DispUEBD_AllRigs_010125-0000_031225-2359` (con o sin `.csv`)
- `DispUEBD_AllRigs_010126-0000_170226-2100` (con o sin `.csv`)
- compara Feb-2026 (`--compare-year 2026 --compare-month 2`)

### Ejecucion con objetivo mensual explicito

Ejemplo (UEBD objetivo 58%, Disp objetivo 82%):

```bash
python3 analisis_brecha_2025_vs_feb2026.py --uebd-objetivo 58 --disp-objetivo 82
```

### Leer objetivo desde mensual 2026

```bash
python3 analisis_brecha_2025_vs_feb2026.py --mensual-2026 "MENSUAL 2026.xlsx" --mensual-rig "TOTAL"
```

### Agregar horas de malla (manual)

```bash
python3 analisis_brecha_2025_vs_feb2026.py --horas-perdida-malla 120
```

### Salidas del script de brecha

En `salidas_brecha_2025_vs_feb2026/`:

- `resumen_brecha_2025_vs_feb2026.csv`
- `aporte_gap_uebd_por_codigo.csv`
- `aporte_gap_disponibilidad_por_codigo.csv`
- `perdida_rendimiento_f09_por_codigo.csv`
- `diario_baseline_2025.csv`
- `diario_comparado_feb2026.csv`
- `graficos/cascada_brecha_uebd_codigos.png`
- `graficos/cascada_brecha_disponibilidad_codigos.png`
- `reporte_brecha_2025_vs_feb2026.xlsx`

---

## 9) Comparador mensual 2025 vs 2026 (formato A1)

Se incluye un script adicional para leer tus mensuales en formato:

- A1: `Equipo`
- B1: `Indices`
- C1: `Unidad`
- D1 en adelante: meses (`Enero`, `Febrero`, ...)

Script:

- `comparador_mensual_perforadoras.py`

### Que hace

1. Compara mensuales por perforadora entre 2025 y 2026.
2. Omite equipos excluidos (por defecto: `PF03, PFAR, PARR`).
3. Recalcula totales mensuales SIN esos equipos.
4. Entrega comparacion de:
   - Total recalculado 2025 vs total original 2025
   - Total recalculado 2026 vs total original 2026
   - Total recalculado 2025 vs 2026

### Uso (con tus nombres)

```bash
python3 comparador_mensual_perforadoras.py "MENSUAL 2025 (1)" "MENSUAL 2026"
```

### Especificar hoja (si no es la primera)

```bash
python3 comparador_mensual_perforadoras.py "MENSUAL 2025 (1).xlsx" "MENSUAL 2026.xlsx" --sheet-2025 "Hoja1" --sheet-2026 "Hoja1"
```

### Cambiar equipos excluidos

```bash
python3 comparador_mensual_perforadoras.py "MENSUAL 2025 (1)" "MENSUAL 2026" --exclude-equipos "PF03,PFAR"
```

### Archivos de salida

En `salidas_comparador_mensual/`:

- `mensual_comparado_por_perforadora.csv`
- `totales_recalculados_sin_excluidas_2025.csv`
- `totales_recalculados_sin_excluidas_2026.csv`
- `totales_2025_vs_original_2025.csv`
- `totales_2026_vs_original_2026.csv`
- `totales_recalculados_2025_vs_2026.csv`
- `reporte_mensual_comparado_sin_excluidas.xlsx`
