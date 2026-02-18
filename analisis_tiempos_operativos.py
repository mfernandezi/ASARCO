#!/usr/bin/env python3
"""
Analisis de tiempos operativos por perforadora.

Lee un archivo CSV con separador ';' y genera:
  - Detalle diario por perforadora (dia operativo A+B)
  - Resumen mensual por perforadora
  - Resumen anual por perforadora
  - Resumen mensual de flota
  - Resumen anual de flota
  - Impacto por codigo para Disponibilidad y UEBD
  - Graficos de cascada (Top N codigos de mayor impacto negativo)

Definicion de dia operativo:
  - Turno A inicia a las 21:00
  - Turno B inicia a las 09:00
  - Un dia operativo incluye Turno A + Turno B
  - Se utiliza WorkDayStarted cuando existe; en caso contrario, se calcula
    como (Time - 21 horas).date().

Formulas:
  Horas_operativas      = Horas_totales - Mantencion_programada - Mantencion_no_programada
  Disponibilidad_ratio  = Horas_operativas / Horas_totales
  UEBD_ratio            = Horas_efectivo / Horas_operativas
  Disponibilidad(%)     = Disponibilidad_ratio * 100
  UEBD(%)               = UEBD_ratio * 100
"""

from __future__ import annotations

import argparse
import csv
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

METRIC_KEYS = [
    "horas_totales",
    "horas_efectivo",
    "horas_reserva",
    "horas_mant_programada",
    "horas_mant_no_programada",
    "horas_otras",
]

DAILY_FIELDNAMES = [
    "fecha_operativa",
    "anio",
    "mes",
    "perforadora",
    "horas_totales",
    "horas_efectivo",
    "horas_reserva",
    "horas_mant_programada",
    "horas_mant_no_programada",
    "horas_otras",
    "horas_operativas",
    "horas_disponibles",
    "disponibilidad_ratio",
    "disponibilidad_pct",
    "disponibilidad_formula_usuario",
    "uebd_ratio",
    "uebd_pct",
    "uebd_formula_usuario",
]

SUMMARY_BASE_FIELDS = [
    "anio",
    "mes",
    "perforadora",
    "dias_con_datos",
    "horas_totales",
    "horas_efectivo",
    "horas_reserva",
    "horas_mant_programada",
    "horas_mant_no_programada",
    "horas_otras",
    "horas_operativas",
    "horas_disponibles",
    "promedio_diario_efectivo_h",
    "promedio_diario_reserva_h",
    "promedio_diario_mant_programada_h",
    "promedio_diario_mant_no_programada_h",
    "disponibilidad_ratio",
    "disponibilidad_pct",
    "disponibilidad_formula_usuario",
    "uebd_ratio",
    "uebd_pct",
    "uebd_formula_usuario",
]

IMPACT_FIELDNAMES = [
    "metrica",
    "ranking",
    "codigo",
    "horas",
    "impacto_ratio",
    "impacto_pct_points",
    "denominador_horas",
    "valor_final_ratio",
    "valor_final_pct",
]

REQUIRED_COLUMNS = {
    "RigName",
    "Time",
    "EndTime",
    "Duration",
    "ShortCode",
    "OnlyCodeName",
    "PlannedCodeName",
}


def normalize_text(value: Optional[str]) -> str:
    text = (value or "").strip().lower()
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def parse_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    raw = value.strip()
    if not raw:
        return None

    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M",
    )
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def parse_date(value: Optional[str]) -> Optional[date]:
    if not value:
        return None
    raw = value.strip()
    if not raw:
        return None

    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        pass

    dt = parse_datetime(raw)
    if dt is not None:
        return dt.date()
    return None


def to_float(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    raw = value.strip()
    if not raw:
        return None
    try:
        return float(raw.replace(",", "."))
    except ValueError:
        return None


def build_daily_row(fecha_operativa: date, perforadora: str) -> Dict[str, float]:
    row: Dict[str, float] = {
        "fecha_operativa": fecha_operativa.isoformat(),
        "anio": fecha_operativa.year,
        "mes": fecha_operativa.month,
        "perforadora": perforadora,
    }
    for key in METRIC_KEYS:
        row[key] = 0.0
    row["horas_operativas"] = 0.0
    row["horas_disponibles"] = 0.0
    row["disponibilidad_ratio"] = 0.0
    row["disponibilidad_pct"] = 0.0
    row["disponibilidad_formula_usuario"] = 0.0
    row["uebd_ratio"] = 0.0
    row["uebd_pct"] = 0.0
    row["uebd_formula_usuario"] = 0.0
    return row


def classify_bucket(row: Dict[str, str]) -> str:
    short_code = normalize_text(row.get("ShortCode"))
    planned = normalize_text(row.get("PlannedCodeName"))
    only_code_name = normalize_text(row.get("OnlyCodeName"))

    if short_code == "efectivo" or only_code_name.startswith("efectivo_"):
        return "efectivo"
    if short_code == "reserva":
        return "reserva"
    if short_code == "mantencion":
        if planned == "programada":
            return "mant_programada"
        return "mant_no_programada"

    return "otras"


def build_code_label(row: Dict[str, str]) -> str:
    code_number = (row.get("OnlyCodeNumber") or "").strip()
    code_name = (row.get("OnlyCodeName") or "").strip()
    code_name_alt = (row.get("CodeName") or "").strip()
    delay_data = (row.get("DelayData") or "").strip()

    if code_number and code_name:
        return f"{code_number}_{code_name}"
    if code_name:
        return code_name
    if code_name_alt:
        return code_name_alt
    if delay_data:
        return delay_data
    return "SIN_CODIGO"


def classify_and_add_hours(bucket: str, metrics: Dict[str, float], hours: float) -> None:
    metrics["horas_totales"] += hours

    if bucket == "efectivo":
        metrics["horas_efectivo"] += hours
        return
    if bucket == "reserva":
        metrics["horas_reserva"] += hours
        return
    if bucket == "mant_programada":
        metrics["horas_mant_programada"] += hours
        return
    if bucket == "mant_no_programada":
        metrics["horas_mant_no_programada"] += hours
        return

    metrics["horas_otras"] += hours


def get_operational_day(row: Dict[str, str], start_dt: Optional[datetime]) -> Optional[date]:
    work_day_started = parse_date(row.get("WorkDayStarted"))
    if work_day_started is not None:
        return work_day_started

    if start_dt is None:
        return None

    # Si no existe WorkDayStarted, se calcula con inicio de dia operativo a las 21:00.
    return (start_dt - timedelta(hours=21)).date()


def finalize_metrics(row: Dict[str, float]) -> None:
    operational_hours = (
        row["horas_totales"] - row["horas_mant_programada"] - row["horas_mant_no_programada"]
    )
    row["horas_operativas"] = max(operational_hours, 0.0)
    # Alias de compatibilidad (mismo valor que horas_operativas)
    row["horas_disponibles"] = row["horas_operativas"]
    row["disponibilidad_ratio"] = (
        row["horas_operativas"] / row["horas_totales"] if row["horas_totales"] > 0 else 0.0
    )
    row["disponibilidad_pct"] = row["disponibilidad_ratio"] * 100.0
    # Formula indicada por usuario: (Horas Operativas / Horas Totales) / 100
    row["disponibilidad_formula_usuario"] = row["disponibilidad_ratio"] / 100.0
    row["uebd_ratio"] = (
        row["horas_efectivo"] / row["horas_operativas"] if row["horas_operativas"] > 0 else 0.0
    )
    row["uebd_pct"] = row["uebd_ratio"] * 100.0
    # Formula indicada por usuario: (Horas Efectivas / Horas Operativas) / 100
    row["uebd_formula_usuario"] = row["uebd_ratio"] / 100.0


def load_daily_metrics(
    input_csv: Path,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> Tuple[List[Dict[str, float]], Dict[str, int], Dict[str, Any]]:
    daily: Dict[Tuple[date, str], Dict[str, float]] = {}
    availability_impact_hours_by_code: Dict[str, float] = defaultdict(float)
    uebd_impact_hours_by_code: Dict[str, float] = defaultdict(float)
    impact_totals = {
        "horas_totales": 0.0,
        "horas_operativas": 0.0,
        "horas_efectivo": 0.0,
    }
    stats = {
        "rows_total": 0,
        "rows_without_operational_day": 0,
        "rows_duration_fallback": 0,
    }

    with input_csv.open("r", encoding=encoding, newline="") as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        if reader.fieldnames is None:
            raise ValueError("El archivo no tiene encabezados.")

        missing_cols = REQUIRED_COLUMNS.difference(reader.fieldnames)
        if missing_cols:
            missing = ", ".join(sorted(missing_cols))
            raise ValueError(f"Faltan columnas requeridas: {missing}")

        for row in reader:
            stats["rows_total"] += 1

            rig_name = (row.get("RigName") or "").strip() or "SIN_RIG"
            start_dt = parse_datetime(row.get("Time"))
            end_dt = parse_datetime(row.get("EndTime"))

            duration_seconds = to_float(row.get("Duration"))
            if duration_seconds is None and start_dt is not None and end_dt is not None:
                duration_seconds = (end_dt - start_dt).total_seconds()
                stats["rows_duration_fallback"] += 1
            if duration_seconds is None:
                duration_seconds = 0.0
            duration_seconds = max(duration_seconds, 0.0)
            duration_hours = duration_seconds / 3600.0

            operational_day = get_operational_day(row, start_dt)
            if operational_day is None:
                stats["rows_without_operational_day"] += 1
                continue

            bucket = classify_bucket(row)
            code_label = build_code_label(row)
            impact_totals["horas_totales"] += duration_hours

            if bucket in {"mant_programada", "mant_no_programada"}:
                availability_impact_hours_by_code[code_label] += duration_hours
            else:
                impact_totals["horas_operativas"] += duration_hours
                if bucket == "efectivo":
                    impact_totals["horas_efectivo"] += duration_hours
                else:
                    uebd_impact_hours_by_code[code_label] += duration_hours

            key = (operational_day, rig_name)
            if key not in daily:
                daily[key] = build_daily_row(operational_day, rig_name)

            classify_and_add_hours(bucket, daily[key], duration_hours)

    rows = sorted(
        daily.values(),
        key=lambda r: (r["fecha_operativa"], r["perforadora"]),
    )
    for row in rows:
        finalize_metrics(row)

    impact_data: Dict[str, Any] = {
        "availability_impact_hours_by_code": dict(availability_impact_hours_by_code),
        "uebd_impact_hours_by_code": dict(uebd_impact_hours_by_code),
        "totals": impact_totals,
    }

    return rows, stats, impact_data


def aggregate_period(
    daily_rows: Iterable[Dict[str, float]],
    granularity: str,
    by_rig: bool,
) -> List[Dict[str, float]]:
    if granularity not in {"mensual", "anual"}:
        raise ValueError("granularity debe ser 'mensual' o 'anual'.")

    grouped: Dict[Tuple, Dict[str, float]] = {}

    for row in daily_rows:
        if granularity == "mensual":
            key_base: Tuple = (row["anio"], row["mes"])
        else:
            key_base = (row["anio"],)

        if by_rig:
            key = key_base + (row["perforadora"],)
        else:
            key = key_base + ("FLOTA",)

        if key not in grouped:
            record: Dict[str, float] = {
                "anio": int(row["anio"]),
                "mes": int(row["mes"]) if granularity == "mensual" else "",
                "perforadora": row["perforadora"] if by_rig else "FLOTA",
                "dias_con_datos": 0,
            }
            for metric in METRIC_KEYS:
                record[metric] = 0.0
            record["horas_operativas"] = 0.0
            record["horas_disponibles"] = 0.0
            record["promedio_diario_efectivo_h"] = 0.0
            record["promedio_diario_reserva_h"] = 0.0
            record["promedio_diario_mant_programada_h"] = 0.0
            record["promedio_diario_mant_no_programada_h"] = 0.0
            record["disponibilidad_ratio"] = 0.0
            record["disponibilidad_pct"] = 0.0
            record["disponibilidad_formula_usuario"] = 0.0
            record["uebd_ratio"] = 0.0
            record["uebd_pct"] = 0.0
            record["uebd_formula_usuario"] = 0.0
            grouped[key] = record

        rec = grouped[key]
        rec["dias_con_datos"] += 1
        for metric in METRIC_KEYS:
            rec[metric] += row[metric]

    result = sorted(
        grouped.values(),
        key=lambda r: (
            r["anio"],
            0 if r["mes"] == "" else r["mes"],
            r["perforadora"],
        ),
    )

    for rec in result:
        operational_hours = (
            rec["horas_totales"] - rec["horas_mant_programada"] - rec["horas_mant_no_programada"]
        )
        rec["horas_operativas"] = max(operational_hours, 0.0)
        # Alias de compatibilidad (mismo valor que horas_operativas)
        rec["horas_disponibles"] = rec["horas_operativas"]

        days = max(int(rec["dias_con_datos"]), 1)
        rec["promedio_diario_efectivo_h"] = rec["horas_efectivo"] / days
        rec["promedio_diario_reserva_h"] = rec["horas_reserva"] / days
        rec["promedio_diario_mant_programada_h"] = rec["horas_mant_programada"] / days
        rec["promedio_diario_mant_no_programada_h"] = rec["horas_mant_no_programada"] / days

        rec["disponibilidad_ratio"] = (
            rec["horas_operativas"] / rec["horas_totales"] if rec["horas_totales"] > 0 else 0.0
        )
        rec["disponibilidad_pct"] = rec["disponibilidad_ratio"] * 100.0
        # Formula indicada por usuario: (Horas Operativas / Horas Totales) / 100
        rec["disponibilidad_formula_usuario"] = rec["disponibilidad_ratio"] / 100.0
        rec["uebd_ratio"] = (
            rec["horas_efectivo"] / rec["horas_operativas"] if rec["horas_operativas"] > 0 else 0.0
        )
        rec["uebd_pct"] = rec["uebd_ratio"] * 100.0
        # Formula indicada por usuario: (Horas Efectivas / Horas Operativas) / 100
        rec["uebd_formula_usuario"] = rec["uebd_ratio"] / 100.0

    return result


def build_impact_rows(
    metrica: str,
    hours_by_code: Dict[str, float],
    denominator_hours: float,
    final_ratio: float,
) -> List[Dict[str, float]]:
    if denominator_hours <= 0:
        return []

    sorted_items = sorted(hours_by_code.items(), key=lambda item: item[1], reverse=True)
    rows: List[Dict[str, float]] = []
    for idx, (code, hours) in enumerate(sorted_items, start=1):
        impact_ratio = hours / denominator_hours
        rows.append(
            {
                "metrica": metrica,
                "ranking": idx,
                "codigo": code,
                "horas": hours,
                "impacto_ratio": impact_ratio,
                "impacto_pct_points": impact_ratio * 100.0,
                "denominador_horas": denominator_hours,
                "valor_final_ratio": final_ratio,
                "valor_final_pct": final_ratio * 100.0,
            }
        )
    return rows


def clip_label(text: str, max_len: int = 28) -> str:
    clean = text.strip()
    if len(clean) <= max_len:
        return clean
    return f"{clean[: max_len - 3]}..."


def get_top_contributions(
    hours_by_code: Dict[str, float],
    denominator_hours: float,
    top_n: int,
) -> List[Tuple[str, float]]:
    if denominator_hours <= 0:
        return []

    sorted_items = sorted(hours_by_code.items(), key=lambda item: item[1], reverse=True)
    top = sorted_items[:top_n]
    other_hours = sum(hours for _, hours in sorted_items[top_n:])

    contributions: List[Tuple[str, float]] = []
    for code, hours in top:
        # Contribucion en puntos porcentuales a la disminucion (signo negativo).
        impact_pp = (hours / denominator_hours) * 100.0
        contributions.append((clip_label(code), -impact_pp))

    if other_hours > 0:
        impact_pp = (other_hours / denominator_hours) * 100.0
        contributions.append(("Otros", -impact_pp))

    return contributions


def try_import_pyplot():
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        return plt, ""
    except Exception as exc:  # pragma: no cover
        return None, str(exc)


def generate_waterfall_chart(
    output_path: Path,
    title: str,
    subtitle: str,
    contributions: List[Tuple[str, float]],
    final_ratio: float,
) -> Tuple[bool, str]:
    plt, err = try_import_pyplot()
    if plt is None:
        return False, err

    labels = ["Base 100%"] + [label for label, _ in contributions] + ["Resultado"]
    final_pct = final_ratio * 100.0
    fig_width = max(10.0, len(labels) * 0.9)
    fig, ax = plt.subplots(figsize=(fig_width, 6.0))

    x_base = 0
    ax.bar(x_base, 100.0, bottom=0.0, color="#2ca02c", edgecolor="black")
    ax.text(x_base, 101.0, "100.0", ha="center", va="bottom", fontsize=8)

    running = 100.0
    for idx, (_label, delta_pp) in enumerate(contributions, start=1):
        next_value = running + delta_pp
        bottom = min(running, next_value)
        height = abs(delta_pp)
        color = "#d62728" if delta_pp < 0 else "#2ca02c"
        ax.bar(idx, height, bottom=bottom, color=color, edgecolor="black")
        ax.text(
            idx,
            bottom + height + 0.8,
            f"{delta_pp:.2f}",
            ha="center",
            va="bottom",
            fontsize=8,
        )
        ax.plot([idx - 1 + 0.35, idx - 0.35], [running, running], color="gray", linewidth=0.8)
        running = next_value

    final_idx = len(labels) - 1
    ax.plot(
        [final_idx - 1 + 0.35, final_idx - 0.35],
        [running, running],
        color="gray",
        linewidth=0.8,
    )
    ax.bar(final_idx, final_pct, bottom=0.0, color="#1f77b4", edgecolor="black")
    ax.text(final_idx, final_pct + 1.0, f"{final_pct:.2f}", ha="center", va="bottom", fontsize=8)

    ax.set_xticks(list(range(len(labels))))
    ax.set_xticklabels(labels, rotation=40, ha="right")
    ax.set_ylabel("Puntos porcentuales (%)")
    ax.set_title(f"{title}\n{subtitle}")
    ax.axhline(0.0, color="black", linewidth=1.0)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    fig.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150)
    plt.close(fig)
    return True, ""


def format_value(value) -> str:
    if isinstance(value, float):
        return f"{value:.4f}"
    return str(value)


def write_csv(path: Path, rows: List[Dict[str, float]], fieldnames: List[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: format_value(row.get(field, "")) for field in fieldnames})


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analiza tiempos operativos por dia/mes/anio con disponibilidad y UEBD."
    )
    parser.add_argument(
        "input_csv",
        type=Path,
        help="Ruta al archivo CSV de entrada.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("salidas_analisis"),
        help="Carpeta de salida para los archivos CSV generados.",
    )
    parser.add_argument(
        "--delimiter",
        default=";",
        help="Separador del CSV de entrada (por defecto ';').",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8-sig",
        help="Encoding del CSV de entrada (por defecto utf-8-sig).",
    )
    parser.add_argument(
        "--top-n-codigos",
        type=int,
        default=10,
        help="Cantidad de codigos a mostrar en los graficos de cascada.",
    )
    parser.add_argument(
        "--sin-graficos-cascada",
        action="store_true",
        help="No generar graficos de cascada (solo CSV de impacto por codigo).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_csv: Path = args.input_csv
    output_dir: Path = args.output_dir
    top_n_codigos = max(int(args.top_n_codigos), 1)

    if not input_csv.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {input_csv}")

    daily_rows, stats, impact_data = load_daily_metrics(
        input_csv=input_csv,
        delimiter=args.delimiter,
        encoding=args.encoding,
    )
    monthly_by_rig = aggregate_period(daily_rows, granularity="mensual", by_rig=True)
    yearly_by_rig = aggregate_period(daily_rows, granularity="anual", by_rig=True)
    monthly_fleet = aggregate_period(daily_rows, granularity="mensual", by_rig=False)
    yearly_fleet = aggregate_period(daily_rows, granularity="anual", by_rig=False)

    monthly_fields = SUMMARY_BASE_FIELDS.copy()
    yearly_fields = [field for field in SUMMARY_BASE_FIELDS if field != "mes"]

    write_csv(output_dir / "diario_por_perforadora.csv", daily_rows, DAILY_FIELDNAMES)
    write_csv(output_dir / "mensual_por_perforadora.csv", monthly_by_rig, monthly_fields)
    write_csv(output_dir / "anual_por_perforadora.csv", yearly_by_rig, yearly_fields)
    write_csv(output_dir / "mensual_flota.csv", monthly_fleet, monthly_fields)
    write_csv(output_dir / "anual_flota.csv", yearly_fleet, yearly_fields)

    totals = impact_data["totals"]
    total_hours = float(totals.get("horas_totales", 0.0))
    operational_hours = float(totals.get("horas_operativas", 0.0))
    effective_hours = float(totals.get("horas_efectivo", 0.0))

    disponibilidad_ratio = operational_hours / total_hours if total_hours > 0 else 0.0
    uebd_ratio = effective_hours / operational_hours if operational_hours > 0 else 0.0

    availability_impact_rows = build_impact_rows(
        metrica="disponibilidad",
        hours_by_code=impact_data["availability_impact_hours_by_code"],
        denominator_hours=total_hours,
        final_ratio=disponibilidad_ratio,
    )
    uebd_impact_rows = build_impact_rows(
        metrica="uebd",
        hours_by_code=impact_data["uebd_impact_hours_by_code"],
        denominator_hours=operational_hours,
        final_ratio=uebd_ratio,
    )
    write_csv(
        output_dir / "impacto_codigos_disponibilidad.csv",
        availability_impact_rows,
        IMPACT_FIELDNAMES,
    )
    write_csv(
        output_dir / "impacto_codigos_uebd.csv",
        uebd_impact_rows,
        IMPACT_FIELDNAMES,
    )

    chart_messages: List[str] = []
    if not args.sin_graficos_cascada:
        availability_contrib = get_top_contributions(
            impact_data["availability_impact_hours_by_code"],
            total_hours,
            top_n_codigos,
        )
        uebd_contrib = get_top_contributions(
            impact_data["uebd_impact_hours_by_code"],
            operational_hours,
            top_n_codigos,
        )

        ok_disp, err_disp = generate_waterfall_chart(
            output_path=output_dir / "graficos" / "cascada_disponibilidad_top_codigos.png",
            title="Cascada de codigos que reducen Disponibilidad",
            subtitle=f"Top {top_n_codigos} + Otros | Base: 100%",
            contributions=availability_contrib,
            final_ratio=disponibilidad_ratio,
        )
        ok_uebd, err_uebd = generate_waterfall_chart(
            output_path=output_dir / "graficos" / "cascada_uebd_top_codigos.png",
            title="Cascada de codigos que reducen UEBD",
            subtitle=f"Top {top_n_codigos} + Otros | Base: 100%",
            contributions=uebd_contrib,
            final_ratio=uebd_ratio,
        )

        if ok_disp:
            chart_messages.append("Grafico cascada disponibilidad: generado")
        else:
            chart_messages.append(f"Grafico cascada disponibilidad: omitido ({err_disp})")
        if ok_uebd:
            chart_messages.append("Grafico cascada UEBD: generado")
        else:
            chart_messages.append(f"Grafico cascada UEBD: omitido ({err_uebd})")
    else:
        chart_messages.append("Graficos de cascada deshabilitados por parametro.")

    print("Analisis finalizado.")
    print(f"Archivo de entrada: {input_csv}")
    print(f"Carpeta de salida: {output_dir}")
    print(f"Filas leidas: {stats['rows_total']}")
    print(f"Filas sin dia operativo: {stats['rows_without_operational_day']}")
    print(f"Filas con fallback de duracion: {stats['rows_duration_fallback']}")
    print(f"Registros diarios (dia+perforadora): {len(daily_rows)}")
    print(f"Resumen mensual por perforadora: {len(monthly_by_rig)} filas")
    print(f"Resumen anual por perforadora: {len(yearly_by_rig)} filas")
    print(f"Impacto codigos disponibilidad: {len(availability_impact_rows)} filas")
    print(f"Impacto codigos UEBD: {len(uebd_impact_rows)} filas")
    for msg in chart_messages:
        print(msg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
