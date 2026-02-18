#!/usr/bin/env python3
"""
Analisis de tiempos operativos por perforadora.

Lee un archivo CSV con separador ';' y genera:
  - Detalle diario por perforadora (dia operativo A+B)
  - Resumen mensual por perforadora
  - Resumen anual por perforadora
  - Resumen mensual de flota
  - Resumen anual de flota

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
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

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
    "uebd_ratio",
    "uebd_pct",
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
    "uebd_ratio",
    "uebd_pct",
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
    row["uebd_ratio"] = 0.0
    row["uebd_pct"] = 0.0
    return row


def classify_and_add_hours(row: Dict[str, str], metrics: Dict[str, float], hours: float) -> None:
    short_code = normalize_text(row.get("ShortCode"))
    planned = normalize_text(row.get("PlannedCodeName"))
    only_code_name = normalize_text(row.get("OnlyCodeName"))

    metrics["horas_totales"] += hours

    if short_code == "efectivo" or only_code_name.startswith("efectivo_"):
        metrics["horas_efectivo"] += hours
        return
    if short_code == "reserva":
        metrics["horas_reserva"] += hours
        return
    if short_code == "mantencion":
        if planned == "programada":
            metrics["horas_mant_programada"] += hours
        else:
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
    row["uebd_ratio"] = (
        row["horas_efectivo"] / row["horas_operativas"] if row["horas_operativas"] > 0 else 0.0
    )
    row["uebd_pct"] = row["uebd_ratio"] * 100.0


def load_daily_metrics(
    input_csv: Path,
    delimiter: str = ";",
    encoding: str = "utf-8-sig",
) -> Tuple[List[Dict[str, float]], Dict[str, int]]:
    daily: Dict[Tuple[date, str], Dict[str, float]] = {}
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

            key = (operational_day, rig_name)
            if key not in daily:
                daily[key] = build_daily_row(operational_day, rig_name)

            classify_and_add_hours(row, daily[key], duration_hours)

    rows = sorted(
        daily.values(),
        key=lambda r: (r["fecha_operativa"], r["perforadora"]),
    )
    for row in rows:
        finalize_metrics(row)

    return rows, stats


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
            record["uebd_ratio"] = 0.0
            record["uebd_pct"] = 0.0
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
        rec["uebd_ratio"] = (
            rec["horas_efectivo"] / rec["horas_operativas"] if rec["horas_operativas"] > 0 else 0.0
        )
        rec["uebd_pct"] = rec["uebd_ratio"] * 100.0

    return result


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
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_csv: Path = args.input_csv
    output_dir: Path = args.output_dir

    if not input_csv.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {input_csv}")

    daily_rows, stats = load_daily_metrics(
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

    print("Analisis finalizado.")
    print(f"Archivo de entrada: {input_csv}")
    print(f"Carpeta de salida: {output_dir}")
    print(f"Filas leidas: {stats['rows_total']}")
    print(f"Filas sin dia operativo: {stats['rows_without_operational_day']}")
    print(f"Filas con fallback de duracion: {stats['rows_duration_fallback']}")
    print(f"Registros diarios (dia+perforadora): {len(daily_rows)}")
    print(f"Resumen mensual por perforadora: {len(monthly_by_rig)} filas")
    print(f"Resumen anual por perforadora: {len(yearly_by_rig)} filas")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
