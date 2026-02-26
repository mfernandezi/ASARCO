"""
Análisis Mensual ASARCO - Generador de Reporte HTML
=====================================================
Genera un reporte HTML completo con:
1. Comparación mensual Enero vs Febrero 2026
2. Real 2026 vs Plan 2025 por mes
3. Análisis de códigos UEBD
4. Disponibilidad UEBD vs Reporte
5. Demoras programadas vs no programadas
6. Detalle de mantención por equipo
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# =============================================================================
# CARGAR DATOS
# =============================================================================
df = pd.read_csv(
    'DispUEBD_AllRigs_010126-0000_170226-2100.csv',
    sep=';', encoding='utf-8-sig'
)
df['Time'] = pd.to_datetime(df['Time'])
df['EndTime'] = pd.to_datetime(df['EndTime'])
df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')
df['Mes'] = df['Time'].dt.month
df['MesNombre'] = df['Time'].dt.month.map({1: 'Enero', 2: 'Febrero'})
df['Equipo'] = df['RigName'].str.replace('-', '')
df['Horas'] = df['Duration'] / 3600

real_2026 = pd.read_excel('MENSUAL 2026.xlsx', sheet_name='Hoja1')
plan_2025 = pd.read_excel('Planes MENSUALES 2025 (1).xlsx', sheet_name='Hoja1')

EQUIPOS = ['PF07', 'PF21', 'PF22', 'PF23', 'PF24', 'PF25', 'PF26']
MESES_DISP = ['Enero', 'Febrero']
INDICES = ['Disponibilidad', 'Utilización', 'Rendimiento', 'Metros', 'Horas Efectivas']

# Clasificar UEBD
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

# =============================================================================
# GENERAR HTML
# =============================================================================

def fmt_pct(v):
    return f"{v*100:.1f}%"

def fmt_num(v, decimals=1):
    return f"{v:,.{decimals}f}"

def fmt_int(v):
    return f"{v:,.0f}"

def diff_class(v, invert=False):
    if invert:
        v = -v
    if v > 0:
        return 'positive'
    elif v < 0:
        return 'negative'
    return 'neutral'

def arrow(v):
    if v > 0:
        return '▲'
    elif v < 0:
        return '▼'
    return '═'

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Análisis Mensual ASARCO 2026</title>
    <style>
        :root {{
            --primary: #1a365d;
            --primary-light: #2b6cb0;
            --accent: #ed8936;
            --green: #38a169;
            --red: #e53e3e;
            --bg: #f7fafc;
            --card-bg: #ffffff;
            --border: #e2e8f0;
            --text: #2d3748;
            --text-light: #718096;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: var(--bg);
            color: var(--text);
            line-height: 1.6;
        }}
        .header {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
            color: white;
            padding: 2rem;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .header h1 {{ font-size: 2rem; margin-bottom: 0.5rem; }}
        .header p {{ opacity: 0.85; font-size: 1.1rem; }}
        .container {{ max-width: 1200px; margin: 0 auto; padding: 1.5rem; }}
        nav {{
            background: var(--card-bg);
            border-radius: 12px;
            padding: 1rem 1.5rem;
            margin-bottom: 1.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            position: sticky;
            top: 0;
            z-index: 100;
        }}
        nav a {{
            display: inline-block;
            padding: 0.5rem 1rem;
            background: var(--bg);
            color: var(--primary);
            text-decoration: none;
            border-radius: 8px;
            font-size: 0.85rem;
            font-weight: 600;
            transition: all 0.2s;
            border: 1px solid var(--border);
        }}
        nav a:hover {{ background: var(--primary); color: white; }}
        .section {{
            background: var(--card-bg);
            border-radius: 12px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .section-title {{
            font-size: 1.3rem;
            color: var(--primary);
            border-bottom: 3px solid var(--accent);
            padding-bottom: 0.5rem;
            margin-bottom: 1rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }}
        .section-title .icon {{ font-size: 1.5rem; }}
        .subtitle {{
            font-size: 1.05rem;
            color: var(--primary-light);
            margin: 1.2rem 0 0.6rem 0;
            font-weight: 600;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 1rem;
            font-size: 0.9rem;
        }}
        th {{
            background: var(--primary);
            color: white;
            padding: 0.6rem 0.8rem;
            text-align: center;
            font-weight: 600;
            font-size: 0.82rem;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }}
        th:first-child {{ text-align: left; border-radius: 8px 0 0 0; }}
        th:last-child {{ border-radius: 0 8px 0 0; }}
        td {{
            padding: 0.5rem 0.8rem;
            text-align: center;
            border-bottom: 1px solid var(--border);
        }}
        td:first-child {{ text-align: left; font-weight: 600; }}
        tr:hover {{ background: #edf2f7; }}
        tr.total-row {{ background: #ebf4ff; font-weight: 700; }}
        tr.total-row:hover {{ background: #dbeafe; }}
        .positive {{ color: var(--green); font-weight: 600; }}
        .negative {{ color: var(--red); font-weight: 600; }}
        .neutral {{ color: var(--text-light); }}
        .badge {{
            display: inline-block;
            padding: 0.15rem 0.6rem;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 700;
        }}
        .badge-si {{ background: #c6f6d5; color: #22543d; }}
        .badge-no {{ background: #fed7d7; color: #9b2c2c; }}
        .badge-mejor {{ background: #c6f6d5; color: #22543d; }}
        .badge-peor {{ background: #fed7d7; color: #9b2c2c; }}
        .badge-revisar {{ background: #fefcbf; color: #744210; }}
        .badge-ok {{ background: #c6f6d5; color: #22543d; }}
        .bar-container {{
            width: 100%;
            background: #edf2f7;
            border-radius: 6px;
            overflow: hidden;
            height: 22px;
            position: relative;
        }}
        .bar {{
            height: 100%;
            border-radius: 6px;
            display: flex;
            align-items: center;
            padding-left: 6px;
            font-size: 0.72rem;
            font-weight: 700;
            color: white;
            transition: width 0.5s ease;
            min-width: fit-content;
        }}
        .bar-efectivo {{ background: var(--green); }}
        .bar-demora {{ background: var(--accent); }}
        .bar-mantencion {{ background: var(--red); }}
        .bar-reserva {{ background: var(--primary-light); }}
        .bar-otro {{ background: var(--text-light); }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 1rem;
            margin-bottom: 1.5rem;
        }}
        .kpi-card {{
            background: var(--bg);
            border-radius: 10px;
            padding: 1rem;
            text-align: center;
            border: 1px solid var(--border);
        }}
        .kpi-card .value {{ font-size: 1.8rem; font-weight: 700; color: var(--primary); }}
        .kpi-card .label {{ font-size: 0.8rem; color: var(--text-light); margin-top: 0.2rem; }}
        .mes-tab {{
            display: inline-block;
            background: var(--primary-light);
            color: white;
            padding: 0.3rem 1rem;
            border-radius: 8px 8px 0 0;
            font-weight: 700;
            font-size: 0.95rem;
            margin-bottom: -1px;
        }}
        .legend {{
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            margin: 0.8rem 0;
            font-size: 0.82rem;
        }}
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 0.3rem;
        }}
        .legend-dot {{
            width: 14px;
            height: 14px;
            border-radius: 4px;
            display: inline-block;
        }}
        .footer {{
            text-align: center;
            padding: 1.5rem;
            color: var(--text-light);
            font-size: 0.85rem;
        }}
    </style>
</head>
<body>

<div class="header">
    <h1>Análisis Mensual ASARCO 2026</h1>
    <p>Perforadoras - Reporte generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
</div>

<div class="container">

<nav>
    <a href="#comparacion">Comparación Mensual</a>
    <a href="#realvsplan">Real vs Plan</a>
    <a href="#uebd">Códigos UEBD</a>
    <a href="#disponibilidad">Disponibilidad UEBD</a>
    <a href="#demoras">Demoras Prog/No Prog</a>
    <a href="#mantencion">Mantención Detalle</a>
</nav>
"""

# =============================================================================
# KPI RESUMEN GLOBAL
# =============================================================================
total_row_ene = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Disponibilidad')]
total_row_feb = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Disponibilidad')]
disp_ene = total_row_ene['Enero'].values[0] if len(total_row_ene) > 0 else 0
disp_feb = total_row_feb['Febrero'].values[0] if len(total_row_feb) > 0 else 0

util_ene = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Utilización')]['Enero'].values[0]
util_feb = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Utilización')]['Febrero'].values[0]

metros_ene = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Metros')]['Enero'].values[0]
metros_feb = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Metros')]['Febrero'].values[0]

rend_ene = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Rendimiento')]['Enero'].values[0]
rend_feb = real_2026[(real_2026['Equipo'] == 'TOTAL') & (real_2026['Índices'] == 'Rendimiento')]['Febrero'].values[0]

html += f"""
<div class="section">
    <div class="section-title"><span class="icon">📊</span> Resumen KPI - Flota Total</div>
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="value">{disp_ene*100:.1f}%</div>
            <div class="label">Disponibilidad Enero</div>
        </div>
        <div class="kpi-card">
            <div class="value">{disp_feb*100:.1f}%</div>
            <div class="label">Disponibilidad Febrero</div>
        </div>
        <div class="kpi-card">
            <div class="value">{util_feb*100:.1f}%</div>
            <div class="label">Utilización Febrero</div>
        </div>
        <div class="kpi-card">
            <div class="value">{fmt_int(metros_ene + metros_feb)}</div>
            <div class="label">Metros Totales Ene+Feb</div>
        </div>
        <div class="kpi-card">
            <div class="value">{fmt_num(rend_feb)}</div>
            <div class="label">Rendimiento Feb (m/hr)</div>
        </div>
    </div>
</div>
"""

# =============================================================================
# SECCIÓN 1: COMPARACIÓN MENSUAL
# =============================================================================
html += """
<div class="section" id="comparacion">
    <div class="section-title"><span class="icon">📈</span> 1. Comparación Mensual - Enero vs Febrero 2026</div>
    <p style="color: var(--text-light); margin-bottom: 1rem;">Comparación directa de cada índice por equipo entre Enero y Febrero 2026.</p>
"""

for indice in INDICES:
    rows_idx = real_2026[real_2026['Índices'] == indice]
    unidad = rows_idx['Unidad'].iloc[0] if len(rows_idx) > 0 else ''
    is_pct = indice in ['Disponibilidad', 'Utilización']
    is_metros = indice == 'Metros'

    html += f'<div class="subtitle">{indice} ({unidad})</div>\n'
    html += """<table>
    <tr><th>Equipo</th><th>Enero</th><th>Febrero</th><th>Diferencia</th><th>Var %</th><th>Tendencia</th></tr>\n"""

    for _, row in rows_idx.iterrows():
        equipo = row['Equipo']
        ene = row.get('Enero', None)
        feb = row.get('Febrero', None)
        if pd.isna(ene) or pd.isna(feb):
            continue

        diff = feb - ene
        var_pct = (diff / ene * 100) if ene != 0 else 0
        cls = diff_class(diff)
        is_total = equipo == 'TOTAL'
        row_cls = ' class="total-row"' if is_total else ''

        if is_pct:
            ene_s, feb_s, diff_s = fmt_pct(ene), fmt_pct(feb), f"{diff*100:+.1f} pp"
        elif is_metros:
            ene_s, feb_s, diff_s = fmt_int(ene), fmt_int(feb), f"{diff:+,.0f}"
        else:
            ene_s, feb_s, diff_s = fmt_num(ene), fmt_num(feb), f"{diff:+.1f}"

        html += f"""    <tr{row_cls}>
        <td>{equipo}</td><td>{ene_s}</td><td>{feb_s}</td>
        <td class="{cls}">{diff_s}</td>
        <td class="{cls}">{var_pct:+.1f}%</td>
        <td class="{cls}">{arrow(diff)}</td>
    </tr>\n"""

    html += "</table>\n"

html += "</div>\n"

# =============================================================================
# SECCIÓN 2: REAL vs PLAN
# =============================================================================
html += """
<div class="section" id="realvsplan">
    <div class="section-title"><span class="icon">🎯</span> 2. Real 2026 vs Plan 2025 (por Mes)</div>
    <p style="color: var(--text-light); margin-bottom: 1rem;">Compara el valor real de cada mes contra el plan correspondiente del mismo mes.</p>
"""

for mes in MESES_DISP:
    html += f'<div class="mes-tab">{mes.upper()}</div>\n'

    for indice in INDICES:
        real_rows = real_2026[real_2026['Índices'] == indice]
        plan_rows = plan_2025[plan_2025['Índices'] == indice]
        unidad = real_rows['Unidad'].iloc[0] if len(real_rows) > 0 else ''
        is_pct = indice in ['Disponibilidad', 'Utilización']
        is_metros = indice == 'Metros'

        html += f'<div class="subtitle">{indice} ({unidad})</div>\n'
        html += f"""<table>
    <tr><th>Equipo</th><th>Plan 2025</th><th>Real 2026</th><th>Diferencia</th><th>Cumple</th></tr>\n"""

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
            cumple = diff >= 0
            cls = diff_class(diff)
            is_total = equipo == 'TOTAL'
            row_cls = ' class="total-row"' if is_total else ''

            if is_pct:
                plan_s, real_s, diff_s = fmt_pct(plan_val), fmt_pct(real_val), f"{diff*100:+.1f} pp"
            elif is_metros:
                plan_s, real_s, diff_s = fmt_int(plan_val), fmt_int(real_val), f"{diff:+,.0f}"
            else:
                plan_s, real_s, diff_s = fmt_num(plan_val), fmt_num(real_val), f"{diff:+.1f}"

            badge = '<span class="badge badge-si">SI</span>' if cumple else '<span class="badge badge-no">NO</span>'

            html += f"""    <tr{row_cls}>
        <td>{equipo}</td><td>{plan_s}</td><td>{real_s}</td>
        <td class="{cls}">{diff_s}</td><td>{badge}</td>
    </tr>\n"""

        html += "</table>\n"

html += "</div>\n"

# =============================================================================
# SECCIÓN 3: CÓDIGOS UEBD
# =============================================================================
cat_global = df.groupby('Categoria_UEBD')['Horas'].sum().sort_values(ascending=False)
total_horas = cat_global.sum()

cat_colors = {
    'Efectivo': '#38a169',
    'Demora Operacional': '#ed8936',
    'Mantención': '#e53e3e',
    'Reserva': '#2b6cb0',
    'Demora Operacional (BV)': '#d69e2e',
    'Mantención (NP)': '#c53030',
    'Otro': '#718096'
}
cat_css = {
    'Efectivo': 'bar-efectivo',
    'Demora Operacional': 'bar-demora',
    'Mantención': 'bar-mantencion',
    'Reserva': 'bar-reserva',
    'Demora Operacional (BV)': 'bar-otro',
    'Mantención (NP)': 'bar-mantencion',
    'Otro': 'bar-otro'
}

html += """
<div class="section" id="uebd">
    <div class="section-title"><span class="icon">🔧</span> 3. Análisis de Códigos UEBD</div>

    <div class="subtitle">Distribución Global de Horas por Categoría</div>
    <div class="legend">
        <div class="legend-item"><span class="legend-dot" style="background:#38a169"></span> Efectivo</div>
        <div class="legend-item"><span class="legend-dot" style="background:#ed8936"></span> Demora Operacional</div>
        <div class="legend-item"><span class="legend-dot" style="background:#e53e3e"></span> Mantención</div>
        <div class="legend-item"><span class="legend-dot" style="background:#2b6cb0"></span> Reserva</div>
    </div>
"""

# Stacked bar
html += '<div style="margin-bottom:1rem;">\n<div class="bar-container" style="height:36px; display:flex;">\n'
for cat, hrs in cat_global.items():
    pct = hrs / total_horas * 100
    css = cat_css.get(cat, 'bar-otro')
    if pct > 3:
        html += f'  <div class="bar {css}" style="width:{pct}%">{cat} {pct:.1f}%</div>\n'
    elif pct > 0.5:
        html += f'  <div class="bar {css}" style="width:{pct}%">{pct:.1f}%</div>\n'
html += '</div></div>\n'

html += """<table>
    <tr><th>Categoría</th><th>Horas</th><th>% del Total</th><th>Barra</th></tr>\n"""
for cat, hrs in cat_global.items():
    pct = hrs / total_horas * 100
    css = cat_css.get(cat, 'bar-otro')
    html += f"""    <tr>
        <td>{cat}</td><td>{fmt_num(hrs)}</td><td>{pct:.1f}%</td>
        <td><div class="bar-container"><div class="bar {css}" style="width:{pct}%">{pct:.1f}%</div></div></td>
    </tr>\n"""
html += f"""    <tr class="total-row">
        <td>TOTAL</td><td>{fmt_num(total_horas)}</td><td>100.0%</td><td></td>
    </tr>\n</table>\n"""

# Top 15 códigos
df_no_efect = df[df['ShortCode'] != 'Efectivo']
top_codes = df_no_efect.groupby(['OnlyCodeNumber', 'OnlyCodeName', 'Categoria_UEBD']).agg(
    Horas=('Horas', 'sum'),
    Ocurrencias=('Horas', 'count')
).sort_values('Horas', ascending=False).head(15)
total_no_efect = df_no_efect['Horas'].sum()

html += """<div class="subtitle">Top 15 Códigos de Demora (Mayor Impacto en Horas)</div>
<table>
    <tr><th>Código</th><th>Nombre</th><th>Categoría</th><th>Horas</th><th>% Demoras</th><th>Ocurrencias</th><th>Impacto</th></tr>\n"""

for (num, name, cat), rdata in top_codes.iterrows():
    pct = rdata['Horas'] / total_no_efect * 100
    css = cat_css.get(cat, 'bar-otro')
    html += f"""    <tr>
        <td style="text-align:center">{int(num)}</td><td>{name}</td><td>{cat}</td>
        <td>{fmt_num(rdata['Horas'])}</td><td>{pct:.1f}%</td><td>{fmt_int(rdata['Ocurrencias'])}</td>
        <td><div class="bar-container"><div class="bar {css}" style="width:{pct*3}%">{pct:.1f}%</div></div></td>
    </tr>\n"""

html += "</table>\n</div>\n"

# =============================================================================
# SECCIÓN 4: DISPONIBILIDAD UEBD vs REPORTE
# =============================================================================
html += """
<div class="section" id="disponibilidad">
    <div class="section-title"><span class="icon">⚙️</span> 4. Disponibilidad: UEBD vs Reporte Mensual</div>
    <p style="color: var(--text-light); margin-bottom: 1rem;">Disponibilidad = (Total Hrs − Mantención) / Total Hrs. Compara el cálculo directo desde datos UEBD contra el dato reportado en el Excel mensual.</p>
"""

for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
    html += f'<div class="mes-tab">{mes_nombre.upper()}</div>\n'
    html += """<table>
    <tr><th>Equipo</th><th>Hrs Total</th><th>Hrs Mant</th><th>Disp UEBD</th><th>Disp Reporte</th><th>Diferencia</th><th>Estado</th></tr>\n"""

    df_mes = df[df['Mes'] == mes_num]
    for equipo in sorted(df_mes['Equipo'].unique()):
        df_eq = df_mes[df_mes['Equipo'] == equipo]
        total_hrs = df_eq['Horas'].sum()
        mant_hrs = df_eq[df_eq['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]['Horas'].sum()
        disp_uebd = (total_hrs - mant_hrs) / total_hrs if total_hrs > 0 else 0

        rep_row = real_2026[(real_2026['Equipo'] == equipo) & (real_2026['Índices'] == 'Disponibilidad')]
        disp_rep = rep_row[mes_nombre].values[0] if len(rep_row) > 0 else None
        diff_pp = (disp_uebd - disp_rep) * 100 if disp_rep is not None else None

        disp_rep_s = fmt_pct(disp_rep) if disp_rep is not None else 'N/A'
        diff_s = f"{diff_pp:+.1f} pp" if diff_pp is not None else 'N/A'
        cls = diff_class(diff_pp) if diff_pp is not None else 'neutral'

        if diff_pp is not None and abs(diff_pp) < 2:
            badge = '<span class="badge badge-ok">OK</span>'
        elif diff_pp is not None:
            badge = '<span class="badge badge-revisar">Revisar</span>'
        else:
            badge = ''

        html += f"""    <tr>
        <td>{equipo}</td><td>{fmt_num(total_hrs)}</td><td>{fmt_num(mant_hrs)}</td>
        <td>{fmt_pct(disp_uebd)}</td><td>{disp_rep_s}</td>
        <td class="{cls}">{diff_s}</td><td>{badge}</td>
    </tr>\n"""

    html += "</table>\n"

html += "</div>\n"

# =============================================================================
# SECCIÓN 5: DEMORAS PROGRAMADAS vs NO PROGRAMADAS
# =============================================================================
html += """
<div class="section" id="demoras">
    <div class="section-title"><span class="icon">⏱️</span> 5. Demoras Programadas vs No Programadas</div>
    <p style="color: var(--text-light); margin-bottom: 1rem;">Distribución del tiempo por tipo de planificación para cada equipo y mes.</p>
"""

for mes_num, mes_nombre in [(1, 'Enero'), (2, 'Febrero')]:
    html += f'<div class="mes-tab">{mes_nombre.upper()}</div>\n'
    html += """<table>
    <tr><th>Equipo</th><th>Programada (h)</th><th>No Programada (h)</th><th>No Asignado (h)</th><th>% No Prog</th><th>Distribución</th></tr>\n"""

    df_mes = df[df['Mes'] == mes_num]
    planned_pivot = df_mes.pivot_table(
        index='Equipo', columns='PlannedCodeName', values='Horas', aggfunc='sum', fill_value=0
    )

    for equipo in sorted(planned_pivot.index):
        prog = planned_pivot.loc[equipo].get('Programada', 0)
        no_prog = planned_pivot.loc[equipo].get('No Programada', 0)
        no_asig = planned_pivot.loc[equipo].get('No Asignado', 0)
        total = prog + no_prog + no_asig
        pct_prog = prog / total * 100 if total > 0 else 0
        pct_no_prog = no_prog / total * 100 if total > 0 else 0
        pct_no_asig = no_asig / total * 100 if total > 0 else 0

        bar_html = f"""<div class="bar-container" style="display:flex">
            <div class="bar bar-efectivo" style="width:{pct_prog}%" title="Programada {pct_prog:.0f}%">{pct_prog:.0f}%</div>
            <div class="bar bar-mantencion" style="width:{pct_no_prog}%" title="No Programada {pct_no_prog:.0f}%">{pct_no_prog:.0f}%</div>
            <div class="bar bar-reserva" style="width:{pct_no_asig}%" title="No Asignado {pct_no_asig:.0f}%">{pct_no_asig:.0f}%</div>
        </div>"""

        html += f"""    <tr>
        <td>{equipo}</td><td>{fmt_num(prog)}</td><td>{fmt_num(no_prog)}</td><td>{fmt_num(no_asig)}</td>
        <td>{pct_no_prog:.1f}%</td><td>{bar_html}</td>
    </tr>\n"""

    html += "</table>\n"
    html += """<div class="legend" style="margin-bottom:1.5rem">
        <div class="legend-item"><span class="legend-dot" style="background:#38a169"></span> Programada</div>
        <div class="legend-item"><span class="legend-dot" style="background:#e53e3e"></span> No Programada</div>
        <div class="legend-item"><span class="legend-dot" style="background:#2b6cb0"></span> No Asignado</div>
    </div>\n"""

html += "</div>\n"

# =============================================================================
# SECCIÓN 6: MANTENCIÓN DETALLE POR EQUIPO
# =============================================================================
html += """
<div class="section" id="mantencion">
    <div class="section-title"><span class="icon">🔩</span> 6. Detalle de Mantención por Equipo y Código</div>
    <p style="color: var(--text-light); margin-bottom: 1rem;">Desglose de horas de mantención por código para cada equipo, comparando Enero vs Febrero.</p>
"""

mant_detail = df[df['Categoria_UEBD'].isin(['Mantención', 'Mantención (NP)'])]

for equipo in sorted(mant_detail['Equipo'].unique()):
    html += f'<div class="subtitle">{equipo}</div>\n'
    html += """<table>
    <tr><th>Código</th><th>Nombre</th><th>Hrs Enero</th><th>Hrs Febrero</th><th>Diferencia</th><th>Tipo</th><th>Impacto</th></tr>\n"""

    eq_mant = mant_detail[mant_detail['Equipo'] == equipo]
    eq_pivot = eq_mant.pivot_table(
        index=['OnlyCodeNumber', 'OnlyCodeName', 'PlannedCodeName'],
        columns='MesNombre', values='Horas', aggfunc='sum', fill_value=0
    ).reset_index()

    eq_pivot_sorted = eq_pivot.copy()
    ene_col = 'Enero' if 'Enero' in eq_pivot.columns else None
    feb_col = 'Febrero' if 'Febrero' in eq_pivot.columns else None
    if ene_col and feb_col:
        eq_pivot_sorted['_total'] = eq_pivot_sorted[ene_col] + eq_pivot_sorted[feb_col]
    elif ene_col:
        eq_pivot_sorted['_total'] = eq_pivot_sorted[ene_col]
    elif feb_col:
        eq_pivot_sorted['_total'] = eq_pivot_sorted[feb_col]
    else:
        eq_pivot_sorted['_total'] = 0
    eq_pivot_sorted = eq_pivot_sorted.sort_values('_total', ascending=False)

    for _, rdata in eq_pivot_sorted.iterrows():
        ene_h = rdata.get('Enero', 0)
        feb_h = rdata.get('Febrero', 0)
        diff_h = feb_h - ene_h
        cls = diff_class(diff_h, invert=True)  # menos mantención = mejor

        if diff_h < -5:
            badge = '<span class="badge badge-mejor">Mejor</span>'
        elif diff_h > 5:
            badge = '<span class="badge badge-peor">Peor</span>'
        else:
            badge = '<span class="badge badge-ok">Estable</span>'

        html += f"""    <tr>
        <td style="text-align:center">{int(rdata['OnlyCodeNumber'])}</td><td>{rdata['OnlyCodeName']}</td>
        <td>{fmt_num(ene_h)}</td><td>{fmt_num(feb_h)}</td>
        <td class="{cls}">{diff_h:+.1f}</td>
        <td>{rdata['PlannedCodeName']}</td><td>{badge}</td>
    </tr>\n"""

    html += "</table>\n"

html += "</div>\n"

# =============================================================================
# FOOTER
# =============================================================================
html += f"""
<div class="footer">
    Reporte generado automáticamente | ASARCO Perforadoras | {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>

</div>
</body>
</html>
"""

# Escribir archivo
with open('Reporte_Analisis_ASARCO_2026.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Reporte HTML generado: Reporte_Analisis_ASARCO_2026.html")
