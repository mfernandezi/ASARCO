#!/usr/bin/env python3
"""
ASARCO Drilling Rig Dashboard Generator
Processes CSV operational data and Excel monthly plans to generate
a self-contained interactive HTML dashboard.
"""

import pandas as pd
import json
from datetime import datetime, timedelta
import math

# ─── 1. LOAD AND PROCESS DATA ───────────────────────────────────────────────

print("Loading CSV data...")
df = pd.read_csv(
    "DispUEBD_AllRigs_010126-0000_170226-2100.csv",
    sep=";",
    encoding="utf-8-sig",
)

df["Time"] = pd.to_datetime(df["Time"])
df["EndTime"] = pd.to_datetime(df["EndTime"])
df["Date"] = df["Time"].dt.date.astype(str)
df["Duration"] = pd.to_numeric(df["Duration"], errors="coerce").fillna(0)
df["DurationHrs"] = df["Duration"] / 3600.0

# Clean rig names for consistency
df["Rig"] = df["RigName"].str.replace("-", "")

rigs = sorted(df["Rig"].unique().tolist())
dates = sorted(df["Date"].unique().tolist())
date_min = dates[0]
date_max = dates[-1]
shifts = sorted(df["ShiftName"].dropna().unique().tolist())

print(f"  Rigs: {rigs}")
print(f"  Date range: {date_min} to {date_max}")
print(f"  Records: {len(df)}")

# ─── 2. AGGREGATE: DAILY KPI PER RIG ────────────────────────────────────────

print("Aggregating daily KPIs...")
daily_by_rig = []
for (date, rig), g in df.groupby(["Date", "Rig"]):
    total_hrs = g["DurationHrs"].sum()
    efectivo_hrs = g.loc[g["ShortCode"] == "Efectivo", "DurationHrs"].sum()
    demora_hrs = g.loc[g["ShortCode"] == "Demora", "DurationHrs"].sum()
    mantencion_hrs = g.loc[g["ShortCode"] == "Mantencion", "DurationHrs"].sum()
    reserva_hrs = g.loc[g["ShortCode"] == "Reserva", "DurationHrs"].sum()

    disponible_hrs = total_hrs - mantencion_hrs
    disponibilidad = (disponible_hrs / total_hrs * 100) if total_hrs > 0 else 0
    utilizacion = (efectivo_hrs / disponible_hrs * 100) if disponible_hrs > 0 else 0

    daily_by_rig.append({
        "date": date,
        "rig": rig,
        "totalHrs": round(total_hrs, 4),
        "efectivoHrs": round(efectivo_hrs, 4),
        "demoraHrs": round(demora_hrs, 4),
        "mantencionHrs": round(mantencion_hrs, 4),
        "reservaHrs": round(reserva_hrs, 4),
        "disponibilidad": round(disponibilidad, 2),
        "utilizacion": round(utilizacion, 2),
    })

# ─── 3. AGGREGATE: DELAY BREAKDOWN PER RIG PER DATE ─────────────────────────

print("Aggregating delay codes...")
delay_data = []
delays_df = df[df["ShortCode"] == "Demora"].copy()
for (date, rig, code), g in delays_df.groupby(["Date", "Rig", "CodeName"]):
    delay_data.append({
        "date": date,
        "rig": rig,
        "code": str(code),
        "hrs": round(g["DurationHrs"].sum(), 4),
        "count": len(g),
        "planned": g["PlannedCodeName"].iloc[0] if len(g) > 0 else "No Asignado",
    })

# ─── 4. AGGREGATE: STATUS BREAKDOWN PER RIG PER DATE ────────────────────────

print("Aggregating status breakdown...")
status_data = []
for (date, rig, status), g in df.groupby(["Date", "Rig", "Status"]):
    status_data.append({
        "date": date,
        "rig": rig,
        "status": status,
        "hrs": round(g["DurationHrs"].sum(), 4),
    })

# ─── 5. AGGREGATE: SHIFT DATA ───────────────────────────────────────────────

print("Aggregating shift data...")
shift_data = []
for (date, rig, shift), g in df.groupby(["Date", "Rig", "ShiftName"]):
    total_hrs = g["DurationHrs"].sum()
    efectivo_hrs = g.loc[g["ShortCode"] == "Efectivo", "DurationHrs"].sum()
    shift_data.append({
        "date": date,
        "rig": rig,
        "shift": shift,
        "totalHrs": round(total_hrs, 4),
        "efectivoHrs": round(efectivo_hrs, 4),
    })

# ─── 6. AGGREGATE: OPERATOR DATA ────────────────────────────────────────────

print("Aggregating operator data...")
operator_data = []
for (date, rig, op), g in df.groupby(["Date", "Rig", "Operator"]):
    if pd.isna(op) or op == "No Data":
        continue
    total_hrs = g["DurationHrs"].sum()
    efectivo_hrs = g.loc[g["ShortCode"] == "Efectivo", "DurationHrs"].sum()
    operator_data.append({
        "date": date,
        "rig": rig,
        "operator": str(op),
        "totalHrs": round(total_hrs, 4),
        "efectivoHrs": round(efectivo_hrs, 4),
    })

# ─── 7. LOAD EXCEL MONTHLY DATA ─────────────────────────────────────────────

print("Loading Excel data...")
monthly_2026 = []
try:
    xl26 = pd.read_excel("MENSUAL 2026.xlsx")
    for _, row in xl26.iterrows():
        equipo = row.iloc[0]
        indice = row.iloc[1]
        unidad = row.iloc[2]
        for col_idx in range(3, len(xl26.columns)):
            month_name = xl26.columns[col_idx]
            val = row.iloc[col_idx]
            if pd.notna(val):
                monthly_2026.append({
                    "equipo": str(equipo),
                    "indice": str(indice),
                    "unidad": str(unidad),
                    "mes": str(month_name),
                    "valor": round(float(val), 4),
                    "year": 2026,
                })
except Exception as e:
    print(f"  Warning reading 2026 Excel: {e}")

monthly_2025 = []
try:
    xl25 = pd.read_excel("Planes MENSUALES 2025 (1).xlsx")
    for _, row in xl25.iterrows():
        equipo = row.iloc[0]
        indice = row.iloc[1]
        unidad = row.iloc[2]
        for col_idx in range(3, len(xl25.columns)):
            month_name = xl25.columns[col_idx]
            val = row.iloc[col_idx]
            if pd.notna(val):
                monthly_2025.append({
                    "equipo": str(equipo),
                    "indice": str(indice),
                    "unidad": str(unidad),
                    "mes": str(month_name),
                    "valor": round(float(val), 4),
                    "year": 2025,
                })
except Exception as e:
    print(f"  Warning reading 2025 Excel: {e}")

# ─── 8. BUILD EMBEDDED DATA OBJECT ──────────────────────────────────────────

embedded_data = {
    "rigs": rigs,
    "dates": dates,
    "dateMin": date_min,
    "dateMax": date_max,
    "shifts": shifts,
    "dailyByRig": daily_by_rig,
    "delayData": delay_data,
    "statusData": status_data,
    "shiftData": shift_data,
    "operatorData": operator_data,
    "monthly2026": monthly_2026,
    "monthly2025": monthly_2025,
}

data_json = json.dumps(embedded_data, ensure_ascii=False)
print(f"  Embedded data size: {len(data_json) / 1024:.0f} KB")

# ─── 9. GENERATE HTML ───────────────────────────────────────────────────────

print("Generating HTML...")

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ASARCO - Dashboard Perforadoras</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<style>
:root {{
  --bg: #0f172a;
  --surface: #1e293b;
  --surface2: #334155;
  --border: #475569;
  --text: #e2e8f0;
  --text-muted: #94a3b8;
  --accent: #3b82f6;
  --accent2: #f59e0b;
  --green: #22c55e;
  --red: #ef4444;
  --orange: #f97316;
  --purple: #a855f7;
  --rangeA: #3b82f6;
  --rangeB: #f59e0b;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
  font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
}}
.header {{
  background: linear-gradient(135deg, #1e3a5f 0%, #0f172a 100%);
  padding: 1.5rem 2rem;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 1rem;
}}
.header h1 {{
  font-size: 1.6rem;
  font-weight: 700;
  color: #fff;
}}
.header h1 span {{ color: var(--accent); }}
.header-info {{
  font-size: 0.85rem;
  color: var(--text-muted);
}}

/* Controls */
.controls {{
  background: var(--surface);
  padding: 1.2rem 2rem;
  border-bottom: 1px solid var(--border);
  display: flex;
  flex-wrap: wrap;
  gap: 1.5rem;
  align-items: flex-end;
}}
.control-group {{
  display: flex;
  flex-direction: column;
  gap: 0.3rem;
}}
.control-group label {{
  font-size: 0.75rem;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--text-muted);
}}
.control-group.range-a label {{ color: var(--rangeA); }}
.control-group.range-b label {{ color: var(--rangeB); }}

input[type="date"], select {{
  background: var(--surface2);
  border: 1px solid var(--border);
  color: var(--text);
  padding: 0.5rem 0.7rem;
  border-radius: 6px;
  font-size: 0.9rem;
  outline: none;
}}
input[type="date"]:focus, select:focus {{
  border-color: var(--accent);
  box-shadow: 0 0 0 2px rgba(59,130,246,0.2);
}}
.btn {{
  padding: 0.5rem 1.2rem;
  border: none;
  border-radius: 6px;
  font-size: 0.85rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}}
.btn-primary {{
  background: var(--accent);
  color: #fff;
}}
.btn-primary:hover {{ background: #2563eb; }}
.btn-sm {{
  padding: 0.3rem 0.7rem;
  font-size: 0.75rem;
  border-radius: 4px;
}}
.btn-outline {{
  background: transparent;
  border: 1px solid var(--border);
  color: var(--text-muted);
}}
.btn-outline:hover {{
  background: var(--surface2);
  color: var(--text);
}}

/* Rig checkboxes */
.rig-selector {{
  display: flex;
  gap: 0.5rem;
  flex-wrap: wrap;
  align-items: center;
}}
.rig-cb {{
  display: flex;
  align-items: center;
  gap: 0.3rem;
  background: var(--surface2);
  padding: 0.35rem 0.6rem;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.82rem;
  user-select: none;
  border: 1px solid transparent;
  transition: all 0.15s;
}}
.rig-cb:hover {{ border-color: var(--accent); }}
.rig-cb input {{ accent-color: var(--accent); cursor: pointer; }}
.rig-cb.checked {{ background: rgba(59,130,246,0.15); border-color: var(--accent); }}

/* KPI cards */
.kpi-row {{
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 1rem;
  padding: 1.5rem 2rem;
}}
.kpi-card {{
  background: var(--surface);
  border-radius: 10px;
  padding: 1.2rem;
  border: 1px solid var(--border);
  position: relative;
  overflow: hidden;
}}
.kpi-card::before {{
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
}}
.kpi-card.blue::before {{ background: var(--accent); }}
.kpi-card.green::before {{ background: var(--green); }}
.kpi-card.orange::before {{ background: var(--orange); }}
.kpi-card.purple::before {{ background: var(--purple); }}
.kpi-card.red::before {{ background: var(--red); }}
.kpi-card.yellow::before {{ background: var(--accent2); }}
.kpi-label {{
  font-size: 0.75rem;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--text-muted);
  margin-bottom: 0.5rem;
}}
.kpi-values {{
  display: flex;
  align-items: baseline;
  gap: 1rem;
}}
.kpi-val {{
  font-size: 1.8rem;
  font-weight: 700;
}}
.kpi-val.a {{ color: var(--rangeA); }}
.kpi-val.b {{ color: var(--rangeB); }}
.kpi-unit {{
  font-size: 0.8rem;
  color: var(--text-muted);
}}
.kpi-vs {{
  font-size: 0.85rem;
  color: var(--text-muted);
  margin: 0 0.2rem;
}}

/* Chart containers */
.dashboard {{
  padding: 0 2rem 2rem;
}}
.charts-grid {{
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(550px, 1fr));
  gap: 1.5rem;
  margin-bottom: 1.5rem;
}}
.chart-card {{
  background: var(--surface);
  border-radius: 10px;
  border: 1px solid var(--border);
  padding: 1.2rem;
  position: relative;
}}
.chart-card.full-width {{
  grid-column: 1 / -1;
}}
.chart-header {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 1rem;
}}
.chart-title {{
  font-size: 1rem;
  font-weight: 600;
}}
.chart-actions {{
  display: flex;
  gap: 0.4rem;
}}
.chart-body {{
  position: relative;
  min-height: 300px;
}}
.chart-body canvas {{
  max-height: 400px;
}}

/* Tables */
.table-card {{
  background: var(--surface);
  border-radius: 10px;
  border: 1px solid var(--border);
  padding: 1.2rem;
  margin-bottom: 1.5rem;
  overflow-x: auto;
}}
.table-header {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 1rem;
  flex-wrap: wrap;
  gap: 0.5rem;
}}
.table-title {{
  font-size: 1rem;
  font-weight: 600;
}}
table {{
  width: 100%;
  border-collapse: collapse;
  font-size: 0.85rem;
}}
thead th {{
  background: var(--surface2);
  padding: 0.7rem 0.8rem;
  text-align: left;
  font-weight: 600;
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-muted);
  border-bottom: 2px solid var(--border);
  white-space: nowrap;
}}
thead th.range-a {{ color: var(--rangeA); }}
thead th.range-b {{ color: var(--rangeB); }}
tbody td {{
  padding: 0.6rem 0.8rem;
  border-bottom: 1px solid rgba(71,85,105,0.4);
  white-space: nowrap;
}}
tbody tr:hover {{ background: rgba(59,130,246,0.05); }}
.badge {{
  display: inline-block;
  padding: 0.15rem 0.5rem;
  border-radius: 4px;
  font-size: 0.75rem;
  font-weight: 600;
}}
.badge-green {{ background: rgba(34,197,94,0.15); color: var(--green); }}
.badge-red {{ background: rgba(239,68,68,0.15); color: var(--red); }}
.badge-orange {{ background: rgba(249,115,22,0.15); color: var(--orange); }}
.badge-blue {{ background: rgba(59,130,246,0.15); color: var(--accent); }}
.badge-yellow {{ background: rgba(245,158,11,0.15); color: var(--accent2); }}

/* Legend */
.legend-row {{
  display: flex;
  gap: 1.5rem;
  padding: 0.5rem 0;
  flex-wrap: wrap;
}}
.legend-item {{
  display: flex;
  align-items: center;
  gap: 0.4rem;
  font-size: 0.82rem;
  color: var(--text-muted);
}}
.legend-dot {{
  width: 12px;
  height: 12px;
  border-radius: 3px;
}}

/* Tabs */
.tabs {{
  display: flex;
  gap: 0;
  margin-bottom: 1.5rem;
  border-bottom: 2px solid var(--border);
}}
.tab {{
  padding: 0.7rem 1.5rem;
  cursor: pointer;
  font-size: 0.9rem;
  font-weight: 500;
  color: var(--text-muted);
  border-bottom: 2px solid transparent;
  margin-bottom: -2px;
  transition: all 0.2s;
}}
.tab:hover {{ color: var(--text); }}
.tab.active {{
  color: var(--accent);
  border-bottom-color: var(--accent);
}}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}

/* Responsive */
@media (max-width: 768px) {{
  .charts-grid {{ grid-template-columns: 1fr; }}
  .kpi-row {{ grid-template-columns: repeat(2, 1fr); }}
  .controls {{ padding: 1rem; }}
  .dashboard {{ padding: 0 1rem 1rem; }}
}}

/* Download overlay */
.download-icon {{
  width: 14px; height: 14px;
  display: inline-block;
  vertical-align: middle;
}}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div>
    <h1><span>ASARCO</span> &mdash; Dashboard Perforadoras</h1>
    <div class="header-info">Datos operacionales &bull; {date_min} al {date_max} &bull; {len(rigs)} equipos</div>
  </div>
  <div style="display:flex;gap:0.5rem;">
    <button class="btn btn-outline btn-sm" onclick="exportAllTablesCSV()">&#11015; Exportar Todas las Tablas</button>
  </div>
</div>

<!-- CONTROLS -->
<div class="controls">
  <div class="control-group range-a">
    <label>&#9632; Rango A &mdash; Desde</label>
    <input type="date" id="rangeAStart" value="{date_min}">
  </div>
  <div class="control-group range-a">
    <label>&#9632; Rango A &mdash; Hasta</label>
    <input type="date" id="rangeAEnd" value="{date_max}">
  </div>
  <div class="control-group range-b">
    <label>&#9632; Rango B &mdash; Desde</label>
    <input type="date" id="rangeBStart" value="{date_min}">
  </div>
  <div class="control-group range-b">
    <label>&#9632; Rango B &mdash; Hasta</label>
    <input type="date" id="rangeBEnd" value="{date_max}">
  </div>
  <div class="control-group">
    <label>Equipos</label>
    <div class="rig-selector" id="rigSelector"></div>
  </div>
  <div class="control-group" style="justify-content:flex-end;">
    <label>&nbsp;</label>
    <button class="btn btn-primary" onclick="updateDashboard()">Actualizar</button>
  </div>
</div>

<!-- KPI CARDS -->
<div class="kpi-row" id="kpiRow"></div>

<!-- TABS -->
<div class="dashboard">
  <div class="tabs">
    <div class="tab active" onclick="switchTab('charts')">Gr&aacute;ficos</div>
    <div class="tab" onclick="switchTab('tables')">Tablas Detalladas</div>
    <div class="tab" onclick="switchTab('delays')">An&aacute;lisis Demoras</div>
    <div class="tab" onclick="switchTab('monthly')">Plan Mensual</div>
  </div>

  <!-- CHARTS TAB -->
  <div class="tab-content active" id="tab-charts">
    <div class="legend-row">
      <div class="legend-item"><div class="legend-dot" style="background:var(--rangeA)"></div> Rango A</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--rangeB)"></div> Rango B</div>
    </div>
    <div class="charts-grid">
      <div class="chart-card" id="card-disponibilidad">
        <div class="chart-header">
          <div class="chart-title">Disponibilidad por Equipo (%)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartDisponibilidad','disponibilidad')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartDisponibilidad"></canvas></div>
      </div>
      <div class="chart-card" id="card-utilizacion">
        <div class="chart-header">
          <div class="chart-title">Utilizaci&oacute;n por Equipo (%)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartUtilizacion','utilizacion')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartUtilizacion"></canvas></div>
      </div>
      <div class="chart-card" id="card-hrsEfectivas">
        <div class="chart-header">
          <div class="chart-title">Horas Efectivas por Equipo</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartHrsEfectivas','hrsEfectivas')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartHrsEfectivas"></canvas></div>
      </div>
      <div class="chart-card" id="card-statusDist">
        <div class="chart-header">
          <div class="chart-title">Distribuci&oacute;n de Horas por Categor&iacute;a</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartStatusDist','distribucion_categorias')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartStatusDist"></canvas></div>
      </div>
      <div class="chart-card full-width" id="card-dailyTrend">
        <div class="chart-header">
          <div class="chart-title">Tendencia Diaria &mdash; Disponibilidad (%)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartDailyTrend','tendencia_diaria')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body" style="min-height:350px;"><canvas id="chartDailyTrend"></canvas></div>
      </div>
      <div class="chart-card" id="card-shiftComp">
        <div class="chart-header">
          <div class="chart-title">Comparaci&oacute;n por Turno (Hrs Efectivas)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartShiftComp','turnos')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartShiftComp"></canvas></div>
      </div>
      <div class="chart-card" id="card-topDelays">
        <div class="chart-header">
          <div class="chart-title">Top 10 Demoras (Hrs)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartTopDelays','top_demoras')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body" style="min-height:350px;"><canvas id="chartTopDelays"></canvas></div>
      </div>
    </div>
  </div>

  <!-- TABLES TAB -->
  <div class="tab-content" id="tab-tables">
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">KPI por Equipo &mdash; Comparaci&oacute;n de Rangos</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('kpiTable','kpi_por_equipo')">&#11015; CSV</button>
      </div>
      <table id="kpiTable"></table>
    </div>
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">Distribuci&oacute;n de Horas por Equipo</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('hoursTable','horas_por_equipo')">&#11015; CSV</button>
      </div>
      <table id="hoursTable"></table>
    </div>
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">Rendimiento por Operador</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('operatorTable','operadores')">&#11015; CSV</button>
      </div>
      <table id="operatorTable"></table>
    </div>
  </div>

  <!-- DELAYS TAB -->
  <div class="tab-content" id="tab-delays">
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">Detalle de Demoras por C&oacute;digo &mdash; Comparaci&oacute;n</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('delayTable','demoras_detalle')">&#11015; CSV</button>
      </div>
      <table id="delayTable"></table>
    </div>
    <div class="charts-grid" style="margin-top:1.5rem;">
      <div class="chart-card" id="card-delayPie">
        <div class="chart-header">
          <div class="chart-title">Demoras Programadas vs No Programadas</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartDelayPie','demoras_prog_vs_noprog')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartDelayPie"></canvas></div>
      </div>
      <div class="chart-card" id="card-delayByRig">
        <div class="chart-header">
          <div class="chart-title">Demoras por Equipo (Hrs)</div>
          <div class="chart-actions">
            <button class="btn btn-outline btn-sm" onclick="downloadChart('chartDelayByRig','demoras_por_equipo')">&#128247; PNG</button>
          </div>
        </div>
        <div class="chart-body"><canvas id="chartDelayByRig"></canvas></div>
      </div>
    </div>
  </div>

  <!-- MONTHLY TAB -->
  <div class="tab-content" id="tab-monthly">
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">Plan Mensual 2026 (Real)</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('monthly2026Table','plan_mensual_2026')">&#11015; CSV</button>
      </div>
      <table id="monthly2026Table"></table>
    </div>
    <div class="table-card">
      <div class="table-header">
        <div class="table-title">Plan Mensual 2025 (Referencia)</div>
        <button class="btn btn-outline btn-sm" onclick="downloadTableCSV('monthly2025Table','plan_mensual_2025')">&#11015; CSV</button>
      </div>
      <table id="monthly2025Table"></table>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════════════════
// EMBEDDED DATA
// ═══════════════════════════════════════════════════════════════════
const DATA = {data_json};

// ═══════════════════════════════════════════════════════════════════
// CHART.JS DEFAULTS
// ═══════════════════════════════════════════════════════════════════
Chart.defaults.color = '#94a3b8';
Chart.defaults.borderColor = 'rgba(71,85,105,0.4)';
Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";

const COLORS_A = ['#3b82f6','#60a5fa','#93c5fd','#2563eb','#1d4ed8','#3b82f6','#60a5fa'];
const COLORS_B = ['#f59e0b','#fbbf24','#fcd34d','#d97706','#b45309','#f59e0b','#fbbf24'];
const STATUS_COLORS = {{
  'Efectivo': '#22c55e',
  'Demora': '#ef4444',
  'Mantencion': '#f59e0b',
  'Reserva': '#8b5cf6',
}};
const CATEGORY_NAMES = {{
  'efectivoHrs': 'Efectivo',
  'demoraHrs': 'Demora',
  'mantencionHrs': 'Mantenci\\u00f3n',
  'reservaHrs': 'Reserva',
}};

let charts = {{}};
let selectedRigs = [...DATA.rigs];

// ═══════════════════════════════════════════════════════════════════
// INITIALIZATION
// ═══════════════════════════════════════════════════════════════════
function init() {{
  buildRigSelector();
  updateDashboard();
}}

function buildRigSelector() {{
  const container = document.getElementById('rigSelector');
  DATA.rigs.forEach(rig => {{
    const lbl = document.createElement('label');
    lbl.className = 'rig-cb checked';
    lbl.innerHTML = `<input type="checkbox" checked value="${{rig}}"> ${{rig}}`;
    lbl.querySelector('input').addEventListener('change', e => {{
      if (e.target.checked) {{
        lbl.classList.add('checked');
      }} else {{
        lbl.classList.remove('checked');
      }}
    }});
    container.appendChild(lbl);
  }});
}}

function getSelectedRigs() {{
  return [...document.querySelectorAll('#rigSelector input:checked')].map(i => i.value);
}}

function getRange(prefix) {{
  return {{
    start: document.getElementById(prefix + 'Start').value,
    end: document.getElementById(prefix + 'End').value,
  }};
}}

// ═══════════════════════════════════════════════════════════════════
// DATA FILTERING
// ═══════════════════════════════════════════════════════════════════
function filterDaily(range, rigs) {{
  return DATA.dailyByRig.filter(d =>
    d.date >= range.start && d.date <= range.end && rigs.includes(d.rig)
  );
}}

function filterDelays(range, rigs) {{
  return DATA.delayData.filter(d =>
    d.date >= range.start && d.date <= range.end && rigs.includes(d.rig)
  );
}}

function filterShifts(range, rigs) {{
  return DATA.shiftData.filter(d =>
    d.date >= range.start && d.date <= range.end && rigs.includes(d.rig)
  );
}}

function filterOperators(range, rigs) {{
  return DATA.operatorData.filter(d =>
    d.date >= range.start && d.date <= range.end && rigs.includes(d.rig)
  );
}}

// Aggregation helpers
function aggByRig(data, field) {{
  const map = {{}};
  data.forEach(d => {{
    if (!map[d.rig]) map[d.rig] = 0;
    map[d.rig] += d[field] || 0;
  }});
  return map;
}}

function aggKpiByRig(data) {{
  const map = {{}};
  data.forEach(d => {{
    if (!map[d.rig]) map[d.rig] = {{ totalHrs: 0, efectivoHrs: 0, demoraHrs: 0, mantencionHrs: 0, reservaHrs: 0 }};
    const r = map[d.rig];
    r.totalHrs += d.totalHrs;
    r.efectivoHrs += d.efectivoHrs;
    r.demoraHrs += d.demoraHrs;
    r.mantencionHrs += d.mantencionHrs;
    r.reservaHrs += d.reservaHrs;
  }});
  Object.keys(map).forEach(rig => {{
    const r = map[rig];
    const disp = r.totalHrs > 0 ? (r.totalHrs - r.mantencionHrs) / r.totalHrs * 100 : 0;
    const dispHrs = r.totalHrs - r.mantencionHrs;
    const util = dispHrs > 0 ? r.efectivoHrs / dispHrs * 100 : 0;
    r.disponibilidad = disp;
    r.utilizacion = util;
  }});
  return map;
}}

// ═══════════════════════════════════════════════════════════════════
// UPDATE DASHBOARD
// ═══════════════════════════════════════════════════════════════════
function updateDashboard() {{
  const rigs = getSelectedRigs();
  const rangeA = getRange('rangeA');
  const rangeB = getRange('rangeB');

  const dailyA = filterDaily(rangeA, rigs);
  const dailyB = filterDaily(rangeB, rigs);
  const kpiA = aggKpiByRig(dailyA);
  const kpiB = aggKpiByRig(dailyB);

  updateKPICards(kpiA, kpiB, rigs);
  updateCharts(dailyA, dailyB, kpiA, kpiB, rigs, rangeA, rangeB);
  updateTables(dailyA, dailyB, kpiA, kpiB, rigs, rangeA, rangeB);
}}

// ═══════════════════════════════════════════════════════════════════
// KPI CARDS
// ═══════════════════════════════════════════════════════════════════
function updateKPICards(kpiA, kpiB, rigs) {{
  // Global averages
  const avgA = calcGlobalKpi(kpiA);
  const avgB = calcGlobalKpi(kpiB);

  const row = document.getElementById('kpiRow');
  row.innerHTML = `
    ${{kpiCard('Disponibilidad', avgA.disponibilidad, avgB.disponibilidad, '%', 'blue')}}
    ${{kpiCard('Utilizaci\\u00f3n', avgA.utilizacion, avgB.utilizacion, '%', 'green')}}
    ${{kpiCard('Hrs Efectivas', avgA.efectivoHrs, avgB.efectivoHrs, 'hrs', 'orange')}}
    ${{kpiCard('Hrs Demora', avgA.demoraHrs, avgB.demoraHrs, 'hrs', 'red')}}
    ${{kpiCard('Hrs Mantenci\\u00f3n', avgA.mantencionHrs, avgB.mantencionHrs, 'hrs', 'yellow')}}
    ${{kpiCard('Hrs Totales', avgA.totalHrs, avgB.totalHrs, 'hrs', 'purple')}}
  `;
}}

function calcGlobalKpi(kpiMap) {{
  let t = 0, e = 0, d = 0, m = 0, re = 0;
  Object.values(kpiMap).forEach(r => {{
    t += r.totalHrs; e += r.efectivoHrs; d += r.demoraHrs; m += r.mantencionHrs; re += r.reservaHrs;
  }});
  const dispHrs = t - m;
  return {{
    totalHrs: t,
    efectivoHrs: e,
    demoraHrs: d,
    mantencionHrs: m,
    reservaHrs: re,
    disponibilidad: t > 0 ? (t - m) / t * 100 : 0,
    utilizacion: dispHrs > 0 ? e / dispHrs * 100 : 0,
  }};
}}

function kpiCard(label, valA, valB, unit, color) {{
  const fmtA = unit === '%' ? valA.toFixed(1) : valA.toFixed(1);
  const fmtB = unit === '%' ? valB.toFixed(1) : valB.toFixed(1);
  return `
    <div class="kpi-card ${{color}}">
      <div class="kpi-label">${{label}}</div>
      <div class="kpi-values">
        <span class="kpi-val a">${{fmtA}}</span>
        <span class="kpi-vs">vs</span>
        <span class="kpi-val b">${{fmtB}}</span>
        <span class="kpi-unit">${{unit}}</span>
      </div>
    </div>`;
}}

// ═══════════════════════════════════════════════════════════════════
// CHARTS
// ═══════════════════════════════════════════════════════════════════
function destroyChart(id) {{
  if (charts[id]) {{ charts[id].destroy(); delete charts[id]; }}
}}

function updateCharts(dailyA, dailyB, kpiA, kpiB, rigs, rangeA, rangeB) {{
  chartDisponibilidad(kpiA, kpiB, rigs);
  chartUtilizacion(kpiA, kpiB, rigs);
  chartHrsEfectivas(kpiA, kpiB, rigs);
  chartStatusDist(kpiA, kpiB);
  chartDailyTrend(dailyA, dailyB, rigs, rangeA, rangeB);
  chartShiftComp(rangeA, rangeB, rigs);
  chartTopDelays(rangeA, rangeB, rigs);
  chartDelayPie(rangeA, rangeB, rigs);
  chartDelayByRig(rangeA, rangeB, rigs);
}}

function chartDisponibilidad(kpiA, kpiB, rigs) {{
  destroyChart('chartDisponibilidad');
  const ctx = document.getElementById('chartDisponibilidad').getContext('2d');
  charts['chartDisponibilidad'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: rigs,
      datasets: [
        {{ label: 'Rango A', data: rigs.map(r => (kpiA[r]||{{}}).disponibilidad||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: rigs.map(r => (kpiB[r]||{{}}).disponibilidad||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, max: 100, ticks: {{ callback: v => v+'%' }} }} }} }}
  }});
}}

function chartUtilizacion(kpiA, kpiB, rigs) {{
  destroyChart('chartUtilizacion');
  const ctx = document.getElementById('chartUtilizacion').getContext('2d');
  charts['chartUtilizacion'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: rigs,
      datasets: [
        {{ label: 'Rango A', data: rigs.map(r => (kpiA[r]||{{}}).utilizacion||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: rigs.map(r => (kpiB[r]||{{}}).utilizacion||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, max: 100, ticks: {{ callback: v => v+'%' }} }} }} }}
  }});
}}

function chartHrsEfectivas(kpiA, kpiB, rigs) {{
  destroyChart('chartHrsEfectivas');
  const ctx = document.getElementById('chartHrsEfectivas').getContext('2d');
  charts['chartHrsEfectivas'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: rigs,
      datasets: [
        {{ label: 'Rango A', data: rigs.map(r => (kpiA[r]||{{}}).efectivoHrs||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: rigs.map(r => (kpiB[r]||{{}}).efectivoHrs||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Horas' }} }} }} }}
  }});
}}

function chartStatusDist(kpiA, kpiB) {{
  destroyChart('chartStatusDist');
  const ctx = document.getElementById('chartStatusDist').getContext('2d');
  const gA = calcGlobalKpi(kpiA);
  const gB = calcGlobalKpi(kpiB);
  const cats = ['efectivoHrs','demoraHrs','mantencionHrs','reservaHrs'];
  const colors = ['#22c55e','#ef4444','#f59e0b','#8b5cf6'];
  charts['chartStatusDist'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: cats.map(c => CATEGORY_NAMES[c]),
      datasets: [
        {{ label: 'Rango A', data: cats.map(c => gA[c]), backgroundColor: cats.map((_, i) => colors[i]+'B3'), borderColor: colors, borderWidth: 1 }},
        {{ label: 'Rango B', data: cats.map(c => gB[c]), backgroundColor: cats.map((_, i) => colors[i]+'66'), borderColor: colors.map(c => c+'99'), borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Horas' }} }} }} }}
  }});
}}

function chartDailyTrend(dailyA, dailyB, rigs, rangeA, rangeB) {{
  destroyChart('chartDailyTrend');
  const ctx = document.getElementById('chartDailyTrend').getContext('2d');

  // Aggregate daily across rigs
  function aggDaily(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.date]) map[d.date] = {{ totalHrs: 0, mantencionHrs: 0 }};
      map[d.date].totalHrs += d.totalHrs;
      map[d.date].mantencionHrs += d.mantencionHrs;
    }});
    const result = [];
    Object.keys(map).sort().forEach(date => {{
      const t = map[date].totalHrs;
      const m = map[date].mantencionHrs;
      result.push({{ date, disponibilidad: t > 0 ? (t - m) / t * 100 : 0 }});
    }});
    return result;
  }}

  const trendA = aggDaily(dailyA);
  const trendB = aggDaily(dailyB);

  const datasets = [];
  if (trendA.length > 0) {{
    datasets.push({{
      label: `Rango A (${{rangeA.start}} a ${{rangeA.end}})`,
      data: trendA.map(d => ({{ x: d.date, y: d.disponibilidad }})),
      borderColor: '#3b82f6',
      backgroundColor: 'rgba(59,130,246,0.1)',
      fill: true,
      tension: 0.3,
      pointRadius: 3,
    }});
  }}
  if (trendB.length > 0) {{
    datasets.push({{
      label: `Rango B (${{rangeB.start}} a ${{rangeB.end}})`,
      data: trendB.map(d => ({{ x: d.date, y: d.disponibilidad }})),
      borderColor: '#f59e0b',
      backgroundColor: 'rgba(245,158,11,0.1)',
      fill: true,
      tension: 0.3,
      pointRadius: 3,
    }});
  }}

  charts['chartDailyTrend'] = new Chart(ctx, {{
    type: 'line',
    data: {{ datasets }},
    options: {{
      responsive: true,
      plugins: {{ legend: {{ display: true }} }},
      scales: {{
        x: {{ type: 'category', title: {{ display: true, text: 'Fecha' }}, ticks: {{ maxRotation: 45 }} }},
        y: {{ beginAtZero: true, max: 100, title: {{ display: true, text: 'Disponibilidad (%)' }}, ticks: {{ callback: v => v+'%' }} }}
      }}
    }}
  }});
}}

function chartShiftComp(rangeA, rangeB, rigs) {{
  destroyChart('chartShiftComp');
  const ctx = document.getElementById('chartShiftComp').getContext('2d');
  const shiftA = filterShifts(rangeA, rigs);
  const shiftB = filterShifts(rangeB, rigs);

  function aggShift(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.shift]) map[d.shift] = 0;
      map[d.shift] += d.efectivoHrs;
    }});
    return map;
  }}

  const sA = aggShift(shiftA);
  const sB = aggShift(shiftB);
  const shiftLabels = [...new Set([...Object.keys(sA), ...Object.keys(sB)])].sort();

  charts['chartShiftComp'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: shiftLabels,
      datasets: [
        {{ label: 'Rango A', data: shiftLabels.map(s => sA[s]||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: shiftLabels.map(s => sB[s]||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Horas Efectivas' }} }} }} }}
  }});
}}

function chartTopDelays(rangeA, rangeB, rigs) {{
  destroyChart('chartTopDelays');
  const ctx = document.getElementById('chartTopDelays').getContext('2d');
  const delA = filterDelays(rangeA, rigs);
  const delB = filterDelays(rangeB, rigs);

  function aggDelayCode(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.code]) map[d.code] = 0;
      map[d.code] += d.hrs;
    }});
    return map;
  }}

  const dA = aggDelayCode(delA);
  const dB = aggDelayCode(delB);
  const allCodes = [...new Set([...Object.keys(dA), ...Object.keys(dB)])];
  allCodes.sort((a, b) => ((dA[b]||0) + (dB[b]||0)) - ((dA[a]||0) + (dB[a]||0)));
  const top10 = allCodes.slice(0, 10);

  charts['chartTopDelays'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: top10.map(c => c.length > 25 ? c.substring(0,25)+'...' : c),
      datasets: [
        {{ label: 'Rango A', data: top10.map(c => dA[c]||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: top10.map(c => dB[c]||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      plugins: {{ legend: {{ display: true }} }},
      scales: {{ x: {{ beginAtZero: true, title: {{ display: true, text: 'Horas' }} }} }}
    }}
  }});
}}

function chartDelayPie(rangeA, rangeB, rigs) {{
  destroyChart('chartDelayPie');
  const ctx = document.getElementById('chartDelayPie').getContext('2d');
  const delA = filterDelays(rangeA, rigs);
  const delB = filterDelays(rangeB, rigs);

  function aggPlanned(data) {{
    const map = {{ 'Programada': 0, 'No Programada': 0, 'No Asignado': 0 }};
    data.forEach(d => {{
      const key = d.planned || 'No Asignado';
      if (!map[key]) map[key] = 0;
      map[key] += d.hrs;
    }});
    return map;
  }}

  const pA = aggPlanned(delA);
  const pB = aggPlanned(delB);
  const labels = Object.keys(pA);
  const colors = ['#22c55e','#ef4444','#94a3b8'];

  charts['chartDelayPie'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: labels,
      datasets: [
        {{ label: 'Rango A', data: labels.map(l => pA[l]||0), backgroundColor: colors.map(c => c+'B3'), borderColor: colors, borderWidth: 1 }},
        {{ label: 'Rango B', data: labels.map(l => pB[l]||0), backgroundColor: colors.map(c => c+'66'), borderColor: colors.map(c => c+'99'), borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Horas' }} }} }} }}
  }});
}}

function chartDelayByRig(rangeA, rangeB, rigs) {{
  destroyChart('chartDelayByRig');
  const ctx = document.getElementById('chartDelayByRig').getContext('2d');
  const delA = filterDelays(rangeA, rigs);
  const delB = filterDelays(rangeB, rigs);

  function aggRig(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.rig]) map[d.rig] = 0;
      map[d.rig] += d.hrs;
    }});
    return map;
  }}

  const rA = aggRig(delA);
  const rB = aggRig(delB);

  charts['chartDelayByRig'] = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: rigs,
      datasets: [
        {{ label: 'Rango A', data: rigs.map(r => rA[r]||0), backgroundColor: 'rgba(59,130,246,0.7)', borderColor: '#3b82f6', borderWidth: 1 }},
        {{ label: 'Rango B', data: rigs.map(r => rB[r]||0), backgroundColor: 'rgba(245,158,11,0.7)', borderColor: '#f59e0b', borderWidth: 1 }},
      ]
    }},
    options: {{ responsive: true, plugins: {{ legend: {{ display: true }} }}, scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Horas Demora' }} }} }} }}
  }});
}}

// ═══════════════════════════════════════════════════════════════════
// TABLES
// ═══════════════════════════════════════════════════════════════════
function updateTables(dailyA, dailyB, kpiA, kpiB, rigs, rangeA, rangeB) {{
  buildKpiTable(kpiA, kpiB, rigs);
  buildHoursTable(kpiA, kpiB, rigs);
  buildOperatorTable(rangeA, rangeB, rigs);
  buildDelayTable(rangeA, rangeB, rigs);
  buildMonthlyTable('monthly2026Table', DATA.monthly2026, 2026);
  buildMonthlyTable('monthly2025Table', DATA.monthly2025, 2025);
}}

function buildKpiTable(kpiA, kpiB, rigs) {{
  const t = document.getElementById('kpiTable');
  let html = `<thead><tr>
    <th>Equipo</th>
    <th class="range-a">Disp. A (%)</th><th class="range-b">Disp. B (%)</th><th>&#916;</th>
    <th class="range-a">Util. A (%)</th><th class="range-b">Util. B (%)</th><th>&#916;</th>
    <th class="range-a">Hrs Efect. A</th><th class="range-b">Hrs Efect. B</th><th>&#916;</th>
  </tr></thead><tbody>`;

  rigs.forEach(rig => {{
    const a = kpiA[rig] || {{ disponibilidad: 0, utilizacion: 0, efectivoHrs: 0 }};
    const b = kpiB[rig] || {{ disponibilidad: 0, utilizacion: 0, efectivoHrs: 0 }};
    const dDisp = (a.disponibilidad - b.disponibilidad).toFixed(1);
    const dUtil = (a.utilizacion - b.utilizacion).toFixed(1);
    const dHrs = (a.efectivoHrs - b.efectivoHrs).toFixed(1);
    html += `<tr>
      <td><strong>${{rig}}</strong></td>
      <td>${{a.disponibilidad.toFixed(1)}}</td><td>${{b.disponibilidad.toFixed(1)}}</td><td>${{deltaBadge(dDisp)}}</td>
      <td>${{a.utilizacion.toFixed(1)}}</td><td>${{b.utilizacion.toFixed(1)}}</td><td>${{deltaBadge(dUtil)}}</td>
      <td>${{a.efectivoHrs.toFixed(1)}}</td><td>${{b.efectivoHrs.toFixed(1)}}</td><td>${{deltaBadge(dHrs)}}</td>
    </tr>`;
  }});

  // Totals
  const gA = calcGlobalKpi(kpiA);
  const gB = calcGlobalKpi(kpiB);
  html += `<tr style="font-weight:700;border-top:2px solid var(--border);">
    <td>TOTAL</td>
    <td>${{gA.disponibilidad.toFixed(1)}}</td><td>${{gB.disponibilidad.toFixed(1)}}</td><td>${{deltaBadge((gA.disponibilidad - gB.disponibilidad).toFixed(1))}}</td>
    <td>${{gA.utilizacion.toFixed(1)}}</td><td>${{gB.utilizacion.toFixed(1)}}</td><td>${{deltaBadge((gA.utilizacion - gB.utilizacion).toFixed(1))}}</td>
    <td>${{gA.efectivoHrs.toFixed(1)}}</td><td>${{gB.efectivoHrs.toFixed(1)}}</td><td>${{deltaBadge((gA.efectivoHrs - gB.efectivoHrs).toFixed(1))}}</td>
  </tr>`;

  html += '</tbody>';
  t.innerHTML = html;
}}

function buildHoursTable(kpiA, kpiB, rigs) {{
  const t = document.getElementById('hoursTable');
  let html = `<thead><tr>
    <th>Equipo</th>
    <th class="range-a">Efectivo A</th><th class="range-b">Efectivo B</th>
    <th class="range-a">Demora A</th><th class="range-b">Demora B</th>
    <th class="range-a">Mant. A</th><th class="range-b">Mant. B</th>
    <th class="range-a">Reserva A</th><th class="range-b">Reserva B</th>
    <th class="range-a">Total A</th><th class="range-b">Total B</th>
  </tr></thead><tbody>`;

  rigs.forEach(rig => {{
    const a = kpiA[rig] || {{ efectivoHrs:0, demoraHrs:0, mantencionHrs:0, reservaHrs:0, totalHrs:0 }};
    const b = kpiB[rig] || {{ efectivoHrs:0, demoraHrs:0, mantencionHrs:0, reservaHrs:0, totalHrs:0 }};
    html += `<tr>
      <td><strong>${{rig}}</strong></td>
      <td>${{a.efectivoHrs.toFixed(1)}}</td><td>${{b.efectivoHrs.toFixed(1)}}</td>
      <td>${{a.demoraHrs.toFixed(1)}}</td><td>${{b.demoraHrs.toFixed(1)}}</td>
      <td>${{a.mantencionHrs.toFixed(1)}}</td><td>${{b.mantencionHrs.toFixed(1)}}</td>
      <td>${{a.reservaHrs.toFixed(1)}}</td><td>${{b.reservaHrs.toFixed(1)}}</td>
      <td>${{a.totalHrs.toFixed(1)}}</td><td>${{b.totalHrs.toFixed(1)}}</td>
    </tr>`;
  }});

  html += '</tbody>';
  t.innerHTML = html;
}}

function buildOperatorTable(rangeA, rangeB, rigs) {{
  const t = document.getElementById('operatorTable');
  const opA = filterOperators(rangeA, rigs);
  const opB = filterOperators(rangeB, rigs);

  function aggOp(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.operator]) map[d.operator] = {{ totalHrs: 0, efectivoHrs: 0 }};
      map[d.operator].totalHrs += d.totalHrs;
      map[d.operator].efectivoHrs += d.efectivoHrs;
    }});
    return map;
  }}

  const mA = aggOp(opA);
  const mB = aggOp(opB);
  const operators = [...new Set([...Object.keys(mA), ...Object.keys(mB)])].sort();

  let html = `<thead><tr>
    <th>Operador</th>
    <th class="range-a">Hrs Efect. A</th><th class="range-b">Hrs Efect. B</th>
    <th class="range-a">Hrs Total A</th><th class="range-b">Hrs Total B</th>
    <th class="range-a">% Efect. A</th><th class="range-b">% Efect. B</th>
  </tr></thead><tbody>`;

  operators.forEach(op => {{
    const a = mA[op] || {{ totalHrs: 0, efectivoHrs: 0 }};
    const b = mB[op] || {{ totalHrs: 0, efectivoHrs: 0 }};
    const pctA = a.totalHrs > 0 ? (a.efectivoHrs / a.totalHrs * 100).toFixed(1) : '0.0';
    const pctB = b.totalHrs > 0 ? (b.efectivoHrs / b.totalHrs * 100).toFixed(1) : '0.0';
    html += `<tr>
      <td>${{op.replace(/_/g,' ')}}</td>
      <td>${{a.efectivoHrs.toFixed(1)}}</td><td>${{b.efectivoHrs.toFixed(1)}}</td>
      <td>${{a.totalHrs.toFixed(1)}}</td><td>${{b.totalHrs.toFixed(1)}}</td>
      <td>${{pctA}}</td><td>${{pctB}}</td>
    </tr>`;
  }});

  html += '</tbody>';
  t.innerHTML = html;
}}

function buildDelayTable(rangeA, rangeB, rigs) {{
  const t = document.getElementById('delayTable');
  const delA = filterDelays(rangeA, rigs);
  const delB = filterDelays(rangeB, rigs);

  function aggDel(data) {{
    const map = {{}};
    data.forEach(d => {{
      if (!map[d.code]) map[d.code] = {{ hrs: 0, count: 0, planned: d.planned }};
      map[d.code].hrs += d.hrs;
      map[d.code].count += d.count;
    }});
    return map;
  }}

  const mA = aggDel(delA);
  const mB = aggDel(delB);
  const codes = [...new Set([...Object.keys(mA), ...Object.keys(mB)])];
  codes.sort((a, b) => ((mA[b]||{{hrs:0}}).hrs + (mB[b]||{{hrs:0}}).hrs) - ((mA[a]||{{hrs:0}}).hrs + (mB[a]||{{hrs:0}}).hrs));

  let html = `<thead><tr>
    <th>C&oacute;digo Demora</th><th>Tipo</th>
    <th class="range-a">Hrs A</th><th class="range-b">Hrs B</th><th>&#916;</th>
    <th class="range-a">Eventos A</th><th class="range-b">Eventos B</th>
  </tr></thead><tbody>`;

  codes.forEach(code => {{
    const a = mA[code] || {{ hrs: 0, count: 0, planned: 'No Asignado' }};
    const b = mB[code] || {{ hrs: 0, count: 0, planned: 'No Asignado' }};
    const delta = (a.hrs - b.hrs).toFixed(1);
    const planned = a.planned || b.planned || 'No Asignado';
    const badgeClass = planned === 'Programada' ? 'badge-green' : planned === 'No Programada' ? 'badge-red' : 'badge-blue';
    html += `<tr>
      <td>${{code}}</td>
      <td><span class="badge ${{badgeClass}}">${{planned}}</span></td>
      <td>${{a.hrs.toFixed(1)}}</td><td>${{b.hrs.toFixed(1)}}</td><td>${{deltaBadge(delta)}}</td>
      <td>${{a.count}}</td><td>${{b.count}}</td>
    </tr>`;
  }});

  html += '</tbody>';
  t.innerHTML = html;
}}

function buildMonthlyTable(tableId, data, year) {{
  const t = document.getElementById(tableId);
  if (!data || data.length === 0) {{
    t.innerHTML = '<tbody><tr><td style="padding:2rem;color:var(--text-muted);">No hay datos disponibles</td></tr></tbody>';
    return;
  }}

  const months = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const equipos = [...new Set(data.map(d => d.equipo))];
  const indices = [...new Set(data.map(d => d.indice))];

  // Build lookup
  const lookup = {{}};
  data.forEach(d => {{
    const key = `${{d.equipo}}|${{d.indice}}|${{d.mes}}`;
    lookup[key] = d.valor;
  }});

  let html = `<thead><tr><th>Equipo</th><th>&Iacute;ndice</th><th>Unidad</th>`;
  months.forEach(m => {{ html += `<th>${{m.substring(0,3)}}</th>`; }});
  html += '</tr></thead><tbody>';

  equipos.forEach(eq => {{
    indices.forEach(idx => {{
      const unidad = (data.find(d => d.equipo === eq && d.indice === idx) || {{}}).unidad || '';
      html += `<tr><td><strong>${{eq}}</strong></td><td>${{idx}}</td><td>${{unidad}}</td>`;
      months.forEach(m => {{
        const val = lookup[`${{eq}}|${{idx}}|${{m}}`];
        if (val !== undefined) {{
          if (unidad === '%') {{
            html += `<td>${{(val * 100).toFixed(1)}}%</td>`;
          }} else {{
            html += `<td>${{val.toFixed(1)}}</td>`;
          }}
        }} else {{
          html += '<td style="color:var(--text-muted)">-</td>';
        }}
      }});
      html += '</tr>';
    }});
  }});

  html += '</tbody>';
  t.innerHTML = html;
}}

function deltaBadge(val) {{
  const n = parseFloat(val);
  if (n > 0) return `<span class="badge badge-green">+${{val}}</span>`;
  if (n < 0) return `<span class="badge badge-red">${{val}}</span>`;
  return `<span class="badge badge-blue">0</span>`;
}}

// ═══════════════════════════════════════════════════════════════════
// DOWNLOAD FUNCTIONS
// ═══════════════════════════════════════════════════════════════════
function downloadChart(canvasId, filename) {{
  const card = document.getElementById(canvasId).closest('.chart-card');
  html2canvas(card, {{
    backgroundColor: '#1e293b',
    scale: 2,
  }}).then(canvas => {{
    const link = document.createElement('a');
    link.download = `${{filename}}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  }});
}}

function downloadTableCSV(tableId, filename) {{
  const table = document.getElementById(tableId);
  const rows = table.querySelectorAll('tr');
  let csv = '';
  rows.forEach(row => {{
    const cols = row.querySelectorAll('th, td');
    const rowData = [];
    cols.forEach(col => {{
      let text = col.innerText.replace(/"/g, '""');
      rowData.push(`"${{text}}"`);
    }});
    csv += rowData.join(',') + '\\n';
  }});
  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csv], {{ type: 'text/csv;charset=utf-8;' }});
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${{filename}}.csv`;
  link.click();
}}

function exportAllTablesCSV() {{
  const tables = [
    ['kpiTable', 'kpi_por_equipo'],
    ['hoursTable', 'horas_por_equipo'],
    ['operatorTable', 'operadores'],
    ['delayTable', 'demoras_detalle'],
    ['monthly2026Table', 'plan_mensual_2026'],
    ['monthly2025Table', 'plan_mensual_2025'],
  ];
  tables.forEach(([id, name]) => {{
    const t = document.getElementById(id);
    if (t && t.innerHTML.trim()) {{
      setTimeout(() => downloadTableCSV(id, name), 200);
    }}
  }});
}}

// ═══════════════════════════════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════════════════════════════
function switchTab(tabName) {{
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + tabName).classList.add('active');
  event.target.classList.add('active');

  // Resize charts after tab switch (Chart.js needs this)
  setTimeout(() => {{
    Object.values(charts).forEach(c => c.resize());
  }}, 50);
}}

// ═══════════════════════════════════════════════════════════════════
// START
// ═══════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', init);
</script>
</body>
</html>"""

# ─── 10. WRITE HTML FILE ────────────────────────────────────────────────────

output_path = "dashboard_asarco.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html)

file_size_mb = len(html) / (1024 * 1024)
print(f"\nDashboard generated: {output_path} ({file_size_mb:.1f} MB)")
print("Done!")
