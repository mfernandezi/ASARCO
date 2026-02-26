"""
Análisis Mensual ASARCO - Perforadoras
=======================================
Script para:
1. Comparación mensual clara (Enero vs Febrero 2026) por equipo
2. Comparación Real 2026 vs Plan 2025
3. Análisis detallado de códigos UEBD y su impacto en Disponibilidad
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# =============================================================================
# 1. CARGAR DATOS
# =============================================================================
print("=" * 70)
print("  ANÁLISIS MENSUAL ASARCO - PERFORADORAS 2026")
print("=" * 70)

# Cargar datos UEBD del CSV
df = pd.read_csv(
    'DispUEBD_AllRigs_010126-0000_170226-2100.csv',
    sep=';', encoding='utf-8-sig'
)
df['Time'] = pd.to_datetime(df['Time'])
df['EndTime'] = pd.to_datetime(df['EndTime'])
df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')
df['Mes'] = df['Time'].dt.month
df['MesNombre'] = df['Time'].dt.month.map({1: 'Enero', 2: 'Febrero'})

# Normalizar RigName para match con Excel (PF-07 -> PF07, PF-21 -> PF21, etc.)
df['Equipo'] = df['RigName'].str.replace('-', '')

# Cargar datos del Excel mensual 2026 (real)
real_2026 = pd.read_excel('MENSUAL 2026.xlsx', sheet_name='Hoja1')

# Cargar planes mensuales 2025
plan_2025 = pd.read_excel('Planes MENSUALES 2025 (1).xlsx', sheet_name='Hoja1')

EQUIPOS = ['PF07', 'PF21', 'PF22', 'PF23', 'PF24', 'PF25', 'PF26']
MESES_DISP = ['Enero', 'Febrero']
INDICES = ['Disponibilidad', 'Utilización', 'Rendimiento', 'Metros', 'Horas Efectivas']

# =============================================================================
# 2. ANÁLISIS SIMPLE Y CLARO: Comparación Mensual Enero vs Febrero
# =============================================================================
print("\n")
print("=" * 70)
print("  PARTE 1: COMPARACIÓN MENSUAL - ENERO vs FEBRERO 2026")
print("=" * 70)

for indice in INDICES:
    rows_idx = real_2026[real_2026['Índices'] == indice]
    unidad = rows_idx['Unidad'].iloc[0] if len(rows_idx) > 0 else ''

    print(f"\n{'─' * 60}")
    if indice == 'Disponibilidad':
        print(f"  {indice} ({unidad})")
    elif indice == 'Utilización':
        print(f"  {indice} ({unidad})")
    elif indice == 'Rendimiento':
        print(f"  {indice} ({unidad})")
    elif indice == 'Metros':
        print(f"  {indice} ({unidad})")
    else:
        print(f"  {indice} ({unidad})")
    print(f"{'─' * 60}")
    print(f"  {'Equipo':<10} {'Enero':>12} {'Febrero':>12} {'Diferencia':>12} {'Var%':>8}")
    print(f"  {'─'*10} {'─'*12} {'─'*12} {'─'*12} {'─'*8}")

    for _, row in rows_idx.iterrows():
        equipo = row['Equipo']
        ene = row.get('Enero', None)
        feb = row.get('Febrero', None)

        if ene is None or feb is None or pd.isna(ene) or pd.isna(feb):
            continue

        diff = feb - ene
        var_pct = (diff / ene * 100) if ene != 0 else 0

        if indice in ['Disponibilidad', 'Utilización']:
            ene_str = f"{ene*100:.1f}%"
            feb_str = f"{feb*100:.1f}%"
            diff_str = f"{diff*100:+.1f}pp"
        elif indice == 'Rendimiento':
            ene_str = f"{ene:.1f}"
            feb_str = f"{feb:.1f}"
            diff_str = f"{diff:+.1f}"
        elif indice == 'Metros':
            ene_str = f"{ene:,.0f}"
            feb_str = f"{feb:,.0f}"
            diff_str = f"{diff:+,.0f}"
        else:
            ene_str = f"{ene:.1f}"
            feb_str = f"{feb:.1f}"
            diff_str = f"{diff:+.1f}"

        # Indicador visual
        if indice in ['Disponibilidad', 'Utilización', 'Rendimiento', 'Metros', 'Horas Efectivas']:
            indicator = " +" if diff > 0 else " -" if diff < 0 else " ="
        else:
            indicator = ""

        print(f"  {equipo:<10} {ene_str:>12} {feb_str:>12} {diff_str:>12} {var_pct:>+7.1f}%{indicator}")

# =============================================================================
# 3. COMPARACIÓN REAL 2026 vs PLAN 2025 (mes a mes)
# =============================================================================
print("\n\n")
print("=" * 70)
print("  PARTE 2: REAL 2026 vs PLAN 2025 (Comparación por Mes)")
print("=" * 70)
print("  Muestra la diferencia entre lo real y lo planificado para cada mes.")

for mes in MESES_DISP:
    print(f"\n{'━' * 60}")
    print(f"  MES: {mes.upper()}")
    print(f"{'━' * 60}")

    for indice in INDICES:
        real_rows = real_2026[real_2026['Índices'] == indice]
        plan_rows = plan_2025[plan_2025['Índices'] == indice]
        unidad = real_rows['Unidad'].iloc[0] if len(real_rows) > 0 else ''

        print(f"\n  {indice} ({unidad}):")
        print(f"    {'Equipo':<10} {'Plan 2025':>12} {'Real 2026':>12} {'Diferencia':>12} {'Cumple':>8}")
        print(f"    {'─'*10} {'─'*12} {'─'*12} {'─'*12} {'─'*8}")

        for equipo in EQUIPOS + ['TOTAL']:
            real_row = real_rows[real_rows['Equipo'] == equipo]
            plan_row = plan_rows[plan_rows['Equipo'] == equipo]

            if len(real_row) == 0 or len(plan_row) == 0:
                continue

            real_val = real_row[mes].values[0]
            plan_val = plan_row[mes].values[0]

            if pd.isna(real_val) or pd.isna(plan_val):
                continue

            diff = real_val - plan_val

            if indice in ['Disponibilidad', 'Utilización']:
                plan_str = f"{plan_val*100:.1f}%"
                real_str = f"{real_val*100:.1f}%"
                diff_str = f"{diff*100:+.1f}pp"
                cumple = "SI" if diff >= 0 else "NO"
            elif indice == 'Rendimiento':
                plan_str = f"{plan_val:.1f}"
                real_str = f"{real_val:.1f}"
                diff_str = f"{diff:+.1f}"
                cumple = "SI" if diff >= 0 else "NO"
            elif indice == 'Metros':
                plan_str = f"{plan_val:,.0f}"
                real_str = f"{real_val:,.0f}"
                diff_str = f"{diff:+,.0f}"
                cumple = "SI" if diff >= 0 else "NO"
            else:
                plan_str = f"{plan_val:.1f}"
                real_str = f"{real_val:.1f}"
                diff_str = f"{diff:+.1f}"
                cumple = "SI" if diff >= 0 else "NO"

            print(f"    {equipo:<10} {plan_str:>12} {real_str:>12} {diff_str:>12} {'  ' + cumple:>8}")


# =============================================================================
# 4. ANÁLISIS DE CÓDIGOS UEBD Y DISPONIBILIDAD
# =============================================================================
print("\n\n")
print("=" * 70)
print("  PARTE 3: ANÁLISIS DE CÓDIGOS UEBD")
print("=" * 70)

# Clasificar códigos UEBD
def clasificar_uebd(row):
    sc = row['ShortCode']
    if sc == 'Efectivo':
        return 'Efectivo'
    elif sc == 'Demora':
        return 'Demora Operacional'
    elif sc == 'Reserva':
        return 'Reserva'
    elif sc == 'Mantencion':
        return 'Mantención'
    else:
        code = row['OnlyCodeNumber']
        if isinstance(code, (int, float)) and not pd.isna(code):
            code = int(code)
            if 600 <= code <= 699:
                return 'Demora Operacional (BV)'
            elif 700 <= code <= 799:
                return 'Mantención (NP)'
        return 'Otro'

df['Categoria_UEBD'] = df.apply(clasificar_uebd, axis=1)

# Horas por categoría
df['Horas'] = df['Duration'] / 3600

# --- 4a. Distribución global de horas por categoría UEBD ---
print("\n  4a. Distribución Global de Horas por Categoría UEBD")
print(f"  {'─' * 55}")

cat_global = df.groupby('Categoria_UEBD')['Horas'].sum().sort_values(ascending=False)
total_horas = cat_global.sum()

print(f"    {'Categoría':<25} {'Horas':>10} {'%':>8}")
print(f"    {'─'*25} {'─'*10} {'─'*8}")
for cat, hrs in cat_global.items():
    pct = hrs / total_horas * 100
    bar = '█' * int(pct / 2)
    print(f"    {cat:<25} {hrs:>10,.1f} {pct:>7.1f}%  {bar}")
print(f"    {'TOTAL':<25} {total_horas:>10,.1f} {'100.0':>7}%")

# --- 4b. Top 15 Códigos que más horas consumen ---
print(f"\n\n  4b. Top 15 Códigos de Demora (Mayor Impacto en Horas)")
print(f"  {'─' * 65}")

# Excluir efectivo para ver solo demoras/mantención
df_no_efect = df[df['ShortCode'] != 'Efectivo']
top_codes = df_no_efect.groupby(['OnlyCodeNumber', 'OnlyCodeName', 'Categoria_UEBD']).agg(
    Horas=('Horas', 'sum'),
    Ocurrencias=('Horas', 'count')
).sort_values('Horas', ascending=False).head(15)

total_no_efect = df_no_efect['Horas'].sum()
print(f"    {'Código':<5} {'Nombre':<22} {'Categoría':<20} {'Horas':>8} {'%':>6} {'N':>6}")
print(f"    {'─'*5} {'─'*22} {'─'*20} {'─'*8} {'─'*6} {'─'*6}")
for (num, name, cat), row in top_codes.iterrows():
    pct = row['Horas'] / total_no_efect * 100
    print(f"    {int(num):<5} {name:<22} {cat:<20} {row['Horas']:>8,.1f} {pct:>5.1f}% {int(row['Ocurrencias']):>6}")

# --- 4c. Códigos por Equipo y Mes ---
print(f"\n\n  4c. Horas de Mantención por Equipo y Mes (Afectan Disponibilidad)")
print(f"  {'─' * 55}")

mant_codes = df[df['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]
mant_pivot = mant_codes.pivot_table(
    index='Equipo', columns='MesNombre', values='Horas', aggfunc='sum', fill_value=0
)

if 'Enero' in mant_pivot.columns and 'Febrero' in mant_pivot.columns:
    mant_pivot['Diferencia'] = mant_pivot['Febrero'] - mant_pivot['Enero']
    mant_pivot = mant_pivot.sort_values('Diferencia', ascending=True)

    print(f"    {'Equipo':<10} {'Enero':>10} {'Febrero':>10} {'Dif':>10} {'Impacto':<10}")
    print(f"    {'─'*10} {'─'*10} {'─'*10} {'─'*10} {'─'*10}")
    for equipo, row in mant_pivot.iterrows():
        diff = row.get('Diferencia', 0)
        impact = "MEJOR" if diff < 0 else "PEOR" if diff > 0 else "IGUAL"
        print(f"    {equipo:<10} {row['Enero']:>10.1f} {row['Febrero']:>10.1f} {diff:>+10.1f} {impact:<10}")


# =============================================================================
# 5. DISPONIBILIDAD CALCULADA DESDE UEBD vs MENSUAL
# =============================================================================
print("\n\n")
print("=" * 70)
print("  PARTE 4: DISPONIBILIDAD - Cálculo UEBD vs Reporte Mensual")
print("=" * 70)
print("  Disponibilidad = (Total Hrs - Mantención) / Total Hrs")
print("  Compara el cálculo directo desde UEBD con el dato reportado.\n")

for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
    print(f"  {'━' * 55}")
    print(f"  MES: {mes_nombre.upper()}")
    print(f"  {'━' * 55}")

    df_mes = df[df['Mes'] == mes_num]

    print(f"    {'Equipo':<10} {'Hrs Total':>10} {'Hrs Mant':>10} {'Disp UEBD':>10} {'Disp Rep':>10} {'Dif':>8}")
    print(f"    {'─'*10} {'─'*10} {'─'*10} {'─'*10} {'─'*10} {'─'*8}")

    for equipo in sorted(df_mes['Equipo'].unique()):
        df_eq = df_mes[df_mes['Equipo'] == equipo]
        total_hrs = df_eq['Horas'].sum()
        mant_hrs = df_eq[df_eq['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]['Horas'].sum()
        disp_uebd = (total_hrs - mant_hrs) / total_hrs if total_hrs > 0 else 0

        # Obtener dato reportado
        rep_row = real_2026[(real_2026['Equipo'] == equipo) & (real_2026['Índices'] == 'Disponibilidad')]
        disp_rep = rep_row[mes_nombre].values[0] if len(rep_row) > 0 else None

        diff = (disp_uebd - disp_rep) * 100 if disp_rep is not None else None
        disp_rep_str = f"{disp_rep*100:.1f}%" if disp_rep is not None else "N/A"
        diff_str = f"{diff:+.1f}pp" if diff is not None else "N/A"

        print(f"    {equipo:<10} {total_hrs:>10.1f} {mant_hrs:>10.1f} {disp_uebd*100:>9.1f}% {disp_rep_str:>10} {diff_str:>8}")


# =============================================================================
# 6. DESGLOSE DE DEMORAS NO PROGRAMADAS vs PROGRAMADAS
# =============================================================================
print("\n\n")
print("=" * 70)
print("  PARTE 5: DEMORAS PROGRAMADAS vs NO PROGRAMADAS por Equipo y Mes")
print("=" * 70)

for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
    print(f"\n  {'━' * 55}")
    print(f"  MES: {mes_nombre.upper()}")
    print(f"  {'━' * 55}")

    df_mes = df[df['Mes'] == mes_num]

    planned_pivot = df_mes.pivot_table(
        index='Equipo', columns='PlannedCodeName', values='Horas', aggfunc='sum', fill_value=0
    )

    print(f"    {'Equipo':<10} {'Programada':>12} {'No Program':>12} {'No Asignad':>12} {'% No Prog':>10}")
    print(f"    {'─'*10} {'─'*12} {'─'*12} {'─'*12} {'─'*10}")

    for equipo in sorted(planned_pivot.index):
        prog = planned_pivot.loc[equipo].get('Programada', 0)
        no_prog = planned_pivot.loc[equipo].get('No Programada', 0)
        no_asig = planned_pivot.loc[equipo].get('No Asignado', 0)
        total = prog + no_prog + no_asig
        pct_no_prog = no_prog / total * 100 if total > 0 else 0

        print(f"    {equipo:<10} {prog:>12.1f} {no_prog:>12.1f} {no_asig:>12.1f} {pct_no_prog:>9.1f}%")


# =============================================================================
# 7. GENERAR EXCEL DE REPORTE
# =============================================================================
print("\n\n  Generando reporte Excel...")

with pd.ExcelWriter('Reporte_Analisis_ASARCO_2026.xlsx', engine='xlsxwriter') as writer:
    workbook = writer.book

    # Formatos
    fmt_title = workbook.add_format({
        'bold': True, 'font_size': 14, 'align': 'center',
        'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1
    })
    fmt_header = workbook.add_format({
        'bold': True, 'font_size': 11, 'align': 'center',
        'bg_color': '#D6E4F0', 'border': 1, 'text_wrap': True
    })
    fmt_pct = workbook.add_format({'num_format': '0.0%', 'align': 'center', 'border': 1})
    fmt_num = workbook.add_format({'num_format': '#,##0.0', 'align': 'center', 'border': 1})
    fmt_int = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
    fmt_text = workbook.add_format({'align': 'left', 'border': 1})
    fmt_text_center = workbook.add_format({'align': 'center', 'border': 1})
    fmt_good = workbook.add_format({
        'num_format': '+0.0;-0.0', 'align': 'center', 'border': 1,
        'bg_color': '#C6EFCE', 'font_color': '#006100'
    })
    fmt_bad = workbook.add_format({
        'num_format': '+0.0;-0.0', 'align': 'center', 'border': 1,
        'bg_color': '#FFC7CE', 'font_color': '#9C0006'
    })
    fmt_section = workbook.add_format({
        'bold': True, 'font_size': 12, 'bg_color': '#4472C4',
        'font_color': 'white', 'border': 1
    })

    # =========================================================================
    # HOJA 1: Comparación Mensual Enero vs Febrero
    # =========================================================================
    ws1 = workbook.add_worksheet('Comparación Mensual')
    ws1.set_column('A:A', 12)
    ws1.set_column('B:B', 18)
    ws1.set_column('C:H', 14)

    row = 0
    ws1.merge_range(row, 0, row, 6, 'COMPARACIÓN MENSUAL - ENERO vs FEBRERO 2026', fmt_title)
    row += 2

    for indice in INDICES:
        real_rows = real_2026[real_2026['Índices'] == indice]
        unidad = real_rows['Unidad'].iloc[0] if len(real_rows) > 0 else ''

        ws1.merge_range(row, 0, row, 6, f'{indice} ({unidad})', fmt_section)
        row += 1

        headers = ['Equipo', 'Enero', 'Febrero', 'Diferencia', 'Var %', 'Tendencia', 'Estado']
        for c, h in enumerate(headers):
            ws1.write(row, c, h, fmt_header)
        row += 1

        for _, rdata in real_rows.iterrows():
            equipo = rdata['Equipo']
            ene = rdata.get('Enero', None)
            feb = rdata.get('Febrero', None)
            if pd.isna(ene) or pd.isna(feb):
                continue

            diff = feb - ene
            var_pct = (diff / ene) if ene != 0 else 0

            ws1.write(row, 0, equipo, fmt_text_center)

            if indice in ['Disponibilidad', 'Utilización']:
                ws1.write(row, 1, ene, fmt_pct)
                ws1.write(row, 2, feb, fmt_pct)
                ws1.write(row, 3, diff * 100, fmt_good if diff >= 0 else fmt_bad)
                ws1.write(row, 4, var_pct, fmt_pct)
            elif indice == 'Metros':
                ws1.write(row, 1, ene, fmt_int)
                ws1.write(row, 2, feb, fmt_int)
                ws1.write(row, 3, diff, fmt_good if diff >= 0 else fmt_bad)
                ws1.write(row, 4, var_pct, fmt_pct)
            else:
                ws1.write(row, 1, ene, fmt_num)
                ws1.write(row, 2, feb, fmt_num)
                ws1.write(row, 3, diff, fmt_good if diff >= 0 else fmt_bad)
                ws1.write(row, 4, var_pct, fmt_pct)

            ws1.write(row, 5, '▲' if diff > 0 else '▼' if diff < 0 else '═', fmt_text_center)
            ws1.write(row, 6, 'Mejora' if diff > 0 else 'Baja' if diff < 0 else 'Igual', fmt_text_center)
            row += 1

        row += 1

    # =========================================================================
    # HOJA 2: Real vs Plan por Mes
    # =========================================================================
    ws2 = workbook.add_worksheet('Real vs Plan')
    ws2.set_column('A:A', 12)
    ws2.set_column('B:F', 14)

    row = 0
    ws2.merge_range(row, 0, row, 5, 'REAL 2026 vs PLAN 2025 - POR MES', fmt_title)
    row += 2

    for mes in MESES_DISP:
        ws2.merge_range(row, 0, row, 5, f'MES: {mes.upper()}', fmt_section)
        row += 1

        for indice in INDICES:
            real_rows = real_2026[real_2026['Índices'] == indice]
            plan_rows = plan_2025[plan_2025['Índices'] == indice]
            unidad = real_rows['Unidad'].iloc[0] if len(real_rows) > 0 else ''

            headers = ['Equipo', f'Plan {mes}', f'Real {mes}', 'Diferencia', 'Cumple', f'{indice} ({unidad})']
            for c, h in enumerate(headers):
                ws2.write(row, c, h, fmt_header)
            row += 1

            for equipo in EQUIPOS + ['TOTAL']:
                real_row = real_rows[real_rows['Equipo'] == equipo]
                plan_row = plan_rows[plan_rows['Equipo'] == equipo]
                if len(real_row) == 0 or len(plan_row) == 0:
                    continue

                real_val = real_row[mes].values[0]
                plan_val = plan_row[mes].values[0]
                if pd.isna(real_val) or pd.isna(plan_val):
                    continue

                diff = real_val - plan_val
                cumple = 'SI' if diff >= 0 else 'NO'

                ws2.write(row, 0, equipo, fmt_text_center)
                if indice in ['Disponibilidad', 'Utilización']:
                    ws2.write(row, 1, plan_val, fmt_pct)
                    ws2.write(row, 2, real_val, fmt_pct)
                    ws2.write(row, 3, diff * 100, fmt_good if diff >= 0 else fmt_bad)
                elif indice == 'Metros':
                    ws2.write(row, 1, plan_val, fmt_int)
                    ws2.write(row, 2, real_val, fmt_int)
                    ws2.write(row, 3, diff, fmt_good if diff >= 0 else fmt_bad)
                else:
                    ws2.write(row, 1, plan_val, fmt_num)
                    ws2.write(row, 2, real_val, fmt_num)
                    ws2.write(row, 3, diff, fmt_good if diff >= 0 else fmt_bad)

                ws2.write(row, 4, cumple, fmt_good if cumple == 'SI' else fmt_bad)
                row += 1

            row += 1
        row += 1

    # =========================================================================
    # HOJA 3: Códigos UEBD
    # =========================================================================
    ws3 = workbook.add_worksheet('Códigos UEBD')
    ws3.set_column('A:A', 8)
    ws3.set_column('B:B', 22)
    ws3.set_column('C:C', 22)
    ws3.set_column('D:G', 14)

    row = 0
    ws3.merge_range(row, 0, row, 5, 'ANÁLISIS DE CÓDIGOS UEBD', fmt_title)
    row += 2

    # Distribución global
    ws3.merge_range(row, 0, row, 5, 'Distribución Global de Horas por Categoría', fmt_section)
    row += 1
    for c, h in enumerate(['Categoría', 'Horas', '% del Total']):
        ws3.write(row, c, h, fmt_header)
    row += 1

    for cat, hrs in cat_global.items():
        pct = hrs / total_horas
        ws3.write(row, 0, cat, fmt_text)
        ws3.write(row, 1, hrs, fmt_num)
        ws3.write(row, 2, pct, fmt_pct)
        row += 1

    row += 2

    # Top 15 códigos de demora
    ws3.merge_range(row, 0, row, 5, 'Top 15 Códigos de Demora (Mayor Impacto)', fmt_section)
    row += 1
    for c, h in enumerate(['Código', 'Nombre', 'Categoría', 'Horas', '% Demoras', 'Ocurrencias']):
        ws3.write(row, c, h, fmt_header)
    row += 1

    for (num, name, cat), rdata in top_codes.iterrows():
        pct = rdata['Horas'] / total_no_efect
        ws3.write(row, 0, int(num), fmt_text_center)
        ws3.write(row, 1, name, fmt_text)
        ws3.write(row, 2, cat, fmt_text)
        ws3.write(row, 3, rdata['Horas'], fmt_num)
        ws3.write(row, 4, pct, fmt_pct)
        ws3.write(row, 5, int(rdata['Ocurrencias']), fmt_int)
        row += 1

    # =========================================================================
    # HOJA 4: Disponibilidad UEBD vs Reporte
    # =========================================================================
    ws4 = workbook.add_worksheet('Disponibilidad UEBD')
    ws4.set_column('A:A', 12)
    ws4.set_column('B:G', 14)

    row = 0
    ws4.merge_range(row, 0, row, 6, 'DISPONIBILIDAD: UEBD vs REPORTE MENSUAL', fmt_title)
    row += 2

    for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
        ws4.merge_range(row, 0, row, 6, f'MES: {mes_nombre.upper()}', fmt_section)
        row += 1
        for c, h in enumerate(['Equipo', 'Hrs Total', 'Hrs Mant', 'Disp UEBD', 'Disp Reporte', 'Diferencia (pp)', 'Observación']):
            ws4.write(row, c, h, fmt_header)
        row += 1

        df_mes = df[df['Mes'] == mes_num]
        for equipo in sorted(df_mes['Equipo'].unique()):
            df_eq = df_mes[df_mes['Equipo'] == equipo]
            total_hrs = df_eq['Horas'].sum()
            mant_hrs = df_eq[df_eq['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]['Horas'].sum()
            disp_uebd = (total_hrs - mant_hrs) / total_hrs if total_hrs > 0 else 0

            rep_row = real_2026[(real_2026['Equipo'] == equipo) & (real_2026['Índices'] == 'Disponibilidad')]
            disp_rep = rep_row[mes_nombre].values[0] if len(rep_row) > 0 else None

            diff_pp = (disp_uebd - disp_rep) * 100 if disp_rep is not None else None

            ws4.write(row, 0, equipo, fmt_text_center)
            ws4.write(row, 1, total_hrs, fmt_num)
            ws4.write(row, 2, mant_hrs, fmt_num)
            ws4.write(row, 3, disp_uebd, fmt_pct)
            if disp_rep is not None:
                ws4.write(row, 4, disp_rep, fmt_pct)
                ws4.write(row, 5, diff_pp, fmt_good if abs(diff_pp) < 2 else fmt_bad)
                obs = 'Coincide' if abs(diff_pp) < 2 else 'Revisar diferencia'
                ws4.write(row, 6, obs, fmt_text_center)
            else:
                ws4.write(row, 4, 'N/A', fmt_text_center)
            row += 1

        row += 2

    # =========================================================================
    # HOJA 5: Mantención por Equipo (detalle)
    # =========================================================================
    ws5 = workbook.add_worksheet('Mantención Detalle')
    ws5.set_column('A:A', 12)
    ws5.set_column('B:B', 22)
    ws5.set_column('C:F', 14)

    row = 0
    ws5.merge_range(row, 0, row, 5, 'DETALLE DE MANTENCIÓN POR EQUIPO Y CÓDIGO', fmt_title)
    row += 2

    mant_detail = df[df['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]

    for equipo in sorted(mant_detail['Equipo'].unique()):
        ws5.merge_range(row, 0, row, 5, f'Equipo: {equipo}', fmt_section)
        row += 1
        for c, h in enumerate(['Código', 'Nombre', 'Hrs Enero', 'Hrs Febrero', 'Diferencia', 'Programado']):
            ws5.write(row, c, h, fmt_header)
        row += 1

        eq_mant = mant_detail[mant_detail['Equipo'] == equipo]
        eq_pivot = eq_mant.pivot_table(
            index=['OnlyCodeNumber', 'OnlyCodeName', 'PlannedCodeName'],
            columns='MesNombre', values='Horas', aggfunc='sum', fill_value=0
        ).reset_index()

        for _, rdata in eq_pivot.iterrows():
            ene_h = rdata.get('Enero', 0)
            feb_h = rdata.get('Febrero', 0)
            diff_h = feb_h - ene_h

            ws5.write(row, 0, int(rdata['OnlyCodeNumber']), fmt_text_center)
            ws5.write(row, 1, rdata['OnlyCodeName'], fmt_text)
            ws5.write(row, 2, ene_h, fmt_num)
            ws5.write(row, 3, feb_h, fmt_num)
            ws5.write(row, 4, diff_h, fmt_good if diff_h <= 0 else fmt_bad)
            ws5.write(row, 5, rdata['PlannedCodeName'], fmt_text_center)
            row += 1

        row += 1

    # =========================================================================
    # HOJA 6: Demoras Programadas vs No Programadas
    # =========================================================================
    ws6 = workbook.add_worksheet('Prog vs No Prog')
    ws6.set_column('A:A', 12)
    ws6.set_column('B:F', 14)

    row = 0
    ws6.merge_range(row, 0, row, 5, 'DEMORAS PROGRAMADAS vs NO PROGRAMADAS', fmt_title)
    row += 2

    for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
        ws6.merge_range(row, 0, row, 5, f'MES: {mes_nombre.upper()}', fmt_section)
        row += 1
        for c, h in enumerate(['Equipo', 'Programada (h)', 'No Prog (h)', 'No Asignado (h)', '% No Prog', 'Total (h)']):
            ws6.write(row, c, h, fmt_header)
        row += 1

        df_mes = df[df['Mes'] == mes_num]
        planned_pivot = df_mes.pivot_table(
            index='Equipo', columns='PlannedCodeName', values='Horas', aggfunc='sum', fill_value=0
        )

        for equipo in sorted(planned_pivot.index):
            prog = planned_pivot.loc[equipo].get('Programada', 0)
            no_prog = planned_pivot.loc[equipo].get('No Programada', 0)
            no_asig = planned_pivot.loc[equipo].get('No Asignado', 0)
            total = prog + no_prog + no_asig
            pct_no_prog = no_prog / total if total > 0 else 0

            ws6.write(row, 0, equipo, fmt_text_center)
            ws6.write(row, 1, prog, fmt_num)
            ws6.write(row, 2, no_prog, fmt_num)
            ws6.write(row, 3, no_asig, fmt_num)
            ws6.write(row, 4, pct_no_prog, fmt_pct)
            ws6.write(row, 5, total, fmt_num)
            row += 1

        row += 2

print("\n  Reporte generado: Reporte_Analisis_ASARCO_2026.xlsx")
print("=" * 70)
print("  ANÁLISIS COMPLETADO")
print("=" * 70)
