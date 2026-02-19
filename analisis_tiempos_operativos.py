#!/usr/bin/env python3
"""
Analisis de tiempos operativos por perforadora.

Lee un archivo CSV con separador ';' y genera:
  - Detalle diario por perforadora (dia operativo A+B)
  - Detalle diario por perforadora y turno
  - Resumen mensual por perforadora
  - Resumen anual por perforadora
  - Resumen mensual por perforadora y turno
  - Resumen mensual de flota
  - Resumen anual de flota
  - Resumen ejecutivo total del periodo
  - Impacto por codigo para Disponibilidad y UEBD
  - Impacto por codigo por perforadora
  - Top dias criticos (baja disponibilidad y baja UEBD)
  - Graficos de cascada (Top N codigos de mayor impacto negativo)
  - Excel extensivo con hojas y graficos

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

SHIFT_DAILY_FIELDNAMES = [
    "fecha_operativa",
    "anio",
    "mes",
    "perforadora",
    "turno",
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

SHIFT_MONTHLY_FIELDNAMES = [
    "anio",
    "mes",
    "perforadora",
    "turno",
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

RESUMEN_EJECUTIVO_FIELDNAMES = [
    "perforadora",
    "dias_con_datos",
    "horas_efectivas",
    "horas_operativas",
    "horas_totales",
    "uebd_formula_usuario",
    "disponibilidad_formula_usuario",
    "uebd_ratio",
    "uebd_pct",
    "disponibilidad_ratio",
    "disponibilidad_pct",
    "horas_reserva",
    "horas_mant_programada",
    "horas_mant_no_programada",
    "horas_otras",
]

TOP_DIAS_FIELDNAMES = [
    "ranking",
    "metric",
    "perforadora",
    "fecha_operativa",
    "valor_ratio",
    "valor_pct",
    "horas_efectivo",
    "horas_operativas",
    "horas_totales",
    "horas_reserva",
    "horas_mant_programada",
    "horas_mant_no_programada",
    "horas_otras",
]

IMPACT_BY_RIG_FIELDNAMES = [
    "metrica",
    "perforadora",
    "ranking",
    "codigo",
    "horas",
    "impacto_ratio",
    "impacto_pct_points",
    "denominador_horas",
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


def build_shift_daily_row(fecha_operativa: date, perforadora: str, turno: str) -> Dict[str, float]:
    row = build_daily_row(fecha_operativa, perforadora)
    row["turno"] = turno
    return row


def normalize_shift_name(value: Optional[str]) -> str:
    shift = (value or "").strip()
    if not shift:
        return "SIN_TURNO"
    normalized = normalize_text(shift)
    if normalized in {"turno a", "a"}:
        return "Turno A"
    if normalized in {"turno b", "b"}:
        return "Turno B"
    return shift


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
) -> Tuple[List[Dict[str, float]], List[Dict[str, float]], Dict[str, int], Dict[str, Any]]:
    daily: Dict[Tuple[date, str], Dict[str, float]] = {}
    shift_daily: Dict[Tuple[date, str, str], Dict[str, float]] = {}
    availability_impact_hours_by_code: Dict[str, float] = defaultdict(float)
    uebd_impact_hours_by_code: Dict[str, float] = defaultdict(float)
    availability_impact_hours_by_code_rig: Dict[Tuple[str, str], float] = defaultdict(float)
    uebd_impact_hours_by_code_rig: Dict[Tuple[str, str], float] = defaultdict(float)
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
            shift_name = normalize_shift_name(row.get("ShiftName"))
            impact_totals["horas_totales"] += duration_hours

            if bucket in {"mant_programada", "mant_no_programada"}:
                availability_impact_hours_by_code[code_label] += duration_hours
                availability_impact_hours_by_code_rig[(rig_name, code_label)] += duration_hours
            else:
                impact_totals["horas_operativas"] += duration_hours
                if bucket == "efectivo":
                    impact_totals["horas_efectivo"] += duration_hours
                else:
                    uebd_impact_hours_by_code[code_label] += duration_hours
                    uebd_impact_hours_by_code_rig[(rig_name, code_label)] += duration_hours

            key = (operational_day, rig_name)
            if key not in daily:
                daily[key] = build_daily_row(operational_day, rig_name)

            classify_and_add_hours(bucket, daily[key], duration_hours)

            shift_key = (operational_day, rig_name, shift_name)
            if shift_key not in shift_daily:
                shift_daily[shift_key] = build_shift_daily_row(operational_day, rig_name, shift_name)
            classify_and_add_hours(bucket, shift_daily[shift_key], duration_hours)

    rows = sorted(
        daily.values(),
        key=lambda r: (r["fecha_operativa"], r["perforadora"]),
    )
    for row in rows:
        finalize_metrics(row)

    shift_rows = sorted(
        shift_daily.values(),
        key=lambda r: (r["fecha_operativa"], r["perforadora"], r["turno"]),
    )
    for row in shift_rows:
        finalize_metrics(row)

    impact_data: Dict[str, Any] = {
        "availability_impact_hours_by_code": dict(availability_impact_hours_by_code),
        "uebd_impact_hours_by_code": dict(uebd_impact_hours_by_code),
        "availability_impact_hours_by_code_rig": dict(availability_impact_hours_by_code_rig),
        "uebd_impact_hours_by_code_rig": dict(uebd_impact_hours_by_code_rig),
        "totals": impact_totals,
    }

    return rows, shift_rows, stats, impact_data


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


def aggregate_shift_monthly(shift_daily_rows: Iterable[Dict[str, float]]) -> List[Dict[str, float]]:
    grouped: Dict[Tuple[int, int, str, str], Dict[str, float]] = {}

    for row in shift_daily_rows:
        key = (int(row["anio"]), int(row["mes"]), str(row["perforadora"]), str(row["turno"]))
        if key not in grouped:
            rec: Dict[str, float] = {
                "anio": int(row["anio"]),
                "mes": int(row["mes"]),
                "perforadora": row["perforadora"],
                "turno": row["turno"],
                "dias_con_datos": 0,
            }
            for metric in METRIC_KEYS:
                rec[metric] = 0.0
            rec["horas_operativas"] = 0.0
            rec["horas_disponibles"] = 0.0
            rec["promedio_diario_efectivo_h"] = 0.0
            rec["promedio_diario_reserva_h"] = 0.0
            rec["promedio_diario_mant_programada_h"] = 0.0
            rec["promedio_diario_mant_no_programada_h"] = 0.0
            rec["disponibilidad_ratio"] = 0.0
            rec["disponibilidad_pct"] = 0.0
            rec["disponibilidad_formula_usuario"] = 0.0
            rec["uebd_ratio"] = 0.0
            rec["uebd_pct"] = 0.0
            rec["uebd_formula_usuario"] = 0.0
            grouped[key] = rec

        rec = grouped[key]
        rec["dias_con_datos"] += 1
        for metric in METRIC_KEYS:
            rec[metric] += row[metric]

    result = sorted(grouped.values(), key=lambda r: (r["anio"], r["mes"], r["perforadora"], r["turno"]))
    for rec in result:
        finalize_metrics(rec)
        days = max(int(rec["dias_con_datos"]), 1)
        rec["promedio_diario_efectivo_h"] = rec["horas_efectivo"] / days
        rec["promedio_diario_reserva_h"] = rec["horas_reserva"] / days
        rec["promedio_diario_mant_programada_h"] = rec["horas_mant_programada"] / days
        rec["promedio_diario_mant_no_programada_h"] = rec["horas_mant_no_programada"] / days

    return result


def aggregate_all_period_by_rig(daily_rows: Iterable[Dict[str, float]]) -> List[Dict[str, float]]:
    grouped: Dict[str, Dict[str, float]] = {}

    for row in daily_rows:
        rig = str(row["perforadora"])
        if rig not in grouped:
            grouped[rig] = {
                "perforadora": rig,
                "dias_con_datos": 0,
                "horas_efectivas": 0.0,
                "horas_operativas": 0.0,
                "horas_totales": 0.0,
                "uebd_formula_usuario": 0.0,
                "disponibilidad_formula_usuario": 0.0,
                "uebd_ratio": 0.0,
                "uebd_pct": 0.0,
                "disponibilidad_ratio": 0.0,
                "disponibilidad_pct": 0.0,
                "horas_reserva": 0.0,
                "horas_mant_programada": 0.0,
                "horas_mant_no_programada": 0.0,
                "horas_otras": 0.0,
            }

        rec = grouped[rig]
        rec["dias_con_datos"] += 1
        rec["horas_efectivas"] += row["horas_efectivo"]
        rec["horas_operativas"] += row["horas_operativas"]
        rec["horas_totales"] += row["horas_totales"]
        rec["horas_reserva"] += row["horas_reserva"]
        rec["horas_mant_programada"] += row["horas_mant_programada"]
        rec["horas_mant_no_programada"] += row["horas_mant_no_programada"]
        rec["horas_otras"] += row["horas_otras"]

    result = sorted(grouped.values(), key=lambda r: r["perforadora"])
    for rec in result:
        total = rec["horas_totales"]
        op = rec["horas_operativas"]
        eff = rec["horas_efectivas"]
        rec["disponibilidad_ratio"] = op / total if total > 0 else 0.0
        rec["disponibilidad_pct"] = rec["disponibilidad_ratio"] * 100.0
        rec["disponibilidad_formula_usuario"] = rec["disponibilidad_ratio"] / 100.0
        rec["uebd_ratio"] = eff / op if op > 0 else 0.0
        rec["uebd_pct"] = rec["uebd_ratio"] * 100.0
        rec["uebd_formula_usuario"] = rec["uebd_ratio"] / 100.0

    total_row = {
        "perforadora": "TOTAL",
        "dias_con_datos": 0,
        "horas_efectivas": 0.0,
        "horas_operativas": 0.0,
        "horas_totales": 0.0,
        "uebd_formula_usuario": 0.0,
        "disponibilidad_formula_usuario": 0.0,
        "uebd_ratio": 0.0,
        "uebd_pct": 0.0,
        "disponibilidad_ratio": 0.0,
        "disponibilidad_pct": 0.0,
        "horas_reserva": 0.0,
        "horas_mant_programada": 0.0,
        "horas_mant_no_programada": 0.0,
        "horas_otras": 0.0,
    }
    for rec in result:
        total_row["dias_con_datos"] += rec["dias_con_datos"]
        total_row["horas_efectivas"] += rec["horas_efectivas"]
        total_row["horas_operativas"] += rec["horas_operativas"]
        total_row["horas_totales"] += rec["horas_totales"]
        total_row["horas_reserva"] += rec["horas_reserva"]
        total_row["horas_mant_programada"] += rec["horas_mant_programada"]
        total_row["horas_mant_no_programada"] += rec["horas_mant_no_programada"]
        total_row["horas_otras"] += rec["horas_otras"]

    total = total_row["horas_totales"]
    op = total_row["horas_operativas"]
    eff = total_row["horas_efectivas"]
    total_row["disponibilidad_ratio"] = op / total if total > 0 else 0.0
    total_row["disponibilidad_pct"] = total_row["disponibilidad_ratio"] * 100.0
    total_row["disponibilidad_formula_usuario"] = total_row["disponibilidad_ratio"] / 100.0
    total_row["uebd_ratio"] = eff / op if op > 0 else 0.0
    total_row["uebd_pct"] = total_row["uebd_ratio"] * 100.0
    total_row["uebd_formula_usuario"] = total_row["uebd_ratio"] / 100.0

    result.append(total_row)
    return result


def build_top_critical_days(
    daily_rows: Iterable[Dict[str, float]],
    metric_key: str,
    metric_name: str,
    top_n: int,
) -> List[Dict[str, float]]:
    valid_rows = [row for row in daily_rows if row["horas_totales"] > 0]
    sorted_rows = sorted(valid_rows, key=lambda r: (r[metric_key], r["fecha_operativa"], r["perforadora"]))
    result: List[Dict[str, float]] = []
    for idx, row in enumerate(sorted_rows[:top_n], start=1):
        result.append(
            {
                "ranking": idx,
                "metric": metric_name,
                "perforadora": row["perforadora"],
                "fecha_operativa": row["fecha_operativa"],
                "valor_ratio": row[metric_key],
                "valor_pct": row[metric_key] * 100.0,
                "horas_efectivo": row["horas_efectivo"],
                "horas_operativas": row["horas_operativas"],
                "horas_totales": row["horas_totales"],
                "horas_reserva": row["horas_reserva"],
                "horas_mant_programada": row["horas_mant_programada"],
                "horas_mant_no_programada": row["horas_mant_no_programada"],
                "horas_otras": row["horas_otras"],
            }
        )
    return result


def build_impact_rows_by_rig(
    metrica: str,
    hours_by_rig_code: Dict[Tuple[str, str], float],
    denominator_by_rig: Dict[str, float],
    top_n_per_rig: int,
) -> List[Dict[str, float]]:
    grouped: Dict[str, List[Tuple[str, float]]] = defaultdict(list)
    for (rig, code), hours in hours_by_rig_code.items():
        grouped[rig].append((code, hours))

    result: List[Dict[str, float]] = []
    for rig in sorted(grouped.keys()):
        denom = float(denominator_by_rig.get(rig, 0.0))
        if denom <= 0:
            continue

        sorted_items = sorted(grouped[rig], key=lambda item: item[1], reverse=True)[:top_n_per_rig]
        for idx, (code, hours) in enumerate(sorted_items, start=1):
            impact_ratio = hours / denom
            result.append(
                {
                    "metrica": metrica,
                    "perforadora": rig,
                    "ranking": idx,
                    "codigo": code,
                    "horas": hours,
                    "impacto_ratio": impact_ratio,
                    "impacto_pct_points": impact_ratio * 100.0,
                    "denominador_horas": denom,
                }
            )
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


def try_import_xlsxwriter():
    try:
        import xlsxwriter

        return xlsxwriter, ""
    except Exception as exc:  # pragma: no cover
        return None, str(exc)


def excel_write_sheet(
    workbook,
    sheet_name: str,
    fieldnames: List[str],
    rows: List[Dict[str, Any]],
):
    ws = workbook.add_worksheet(sheet_name[:31])
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    number_fmt = workbook.add_format({"num_format": "0.0000"})
    pct_fmt = workbook.add_format({"num_format": "0.00"})

    for col_idx, col_name in enumerate(fieldnames):
        ws.write(0, col_idx, col_name, header_fmt)

    for row_idx, row in enumerate(rows, start=1):
        for col_idx, col_name in enumerate(fieldnames):
            value = row.get(col_name, "")
            if isinstance(value, float):
                if col_name.endswith("_pct") or col_name.endswith("pct_points"):
                    ws.write_number(row_idx, col_idx, value, pct_fmt)
                else:
                    ws.write_number(row_idx, col_idx, value, number_fmt)
            elif isinstance(value, int):
                ws.write_number(row_idx, col_idx, value)
            else:
                ws.write(row_idx, col_idx, str(value))

    for col_idx, col_name in enumerate(fieldnames):
        width = max(12, min(40, len(col_name) + 2))
        ws.set_column(col_idx, col_idx, width)

    return ws


def write_excel_report(
    output_path: Path,
    resumen_ejecutivo_rows: List[Dict[str, float]],
    diario_rows: List[Dict[str, float]],
    mensual_por_rig_rows: List[Dict[str, float]],
    anual_por_rig_rows: List[Dict[str, float]],
    mensual_flota_rows: List[Dict[str, float]],
    anual_flota_rows: List[Dict[str, float]],
    shift_mensual_rows: List[Dict[str, float]],
    impacto_disp_rows: List[Dict[str, float]],
    impacto_uebd_rows: List[Dict[str, float]],
    impacto_disp_rig_rows: List[Dict[str, float]],
    impacto_uebd_rig_rows: List[Dict[str, float]],
    top_dias_disp_rows: List[Dict[str, float]],
    top_dias_uebd_rows: List[Dict[str, float]],
    waterfall_disp_png: Optional[Path] = None,
    waterfall_uebd_png: Optional[Path] = None,
) -> Tuple[bool, str]:
    xlsxwriter, err = try_import_xlsxwriter()
    if xlsxwriter is None:
        return False, err

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output_path))

    ws_resumen = excel_write_sheet(
        workbook,
        "Resumen_Ejecutivo",
        RESUMEN_EJECUTIVO_FIELDNAMES,
        resumen_ejecutivo_rows,
    )
    excel_write_sheet(workbook, "Diario_Perforadora", DAILY_FIELDNAMES, diario_rows)
    ws_mensual_rig = excel_write_sheet(
        workbook,
        "Mensual_Perforadora",
        SUMMARY_BASE_FIELDS,
        mensual_por_rig_rows,
    )
    excel_write_sheet(
        workbook,
        "Anual_Perforadora",
        [field for field in SUMMARY_BASE_FIELDS if field != "mes"],
        anual_por_rig_rows,
    )
    ws_mensual_flota = excel_write_sheet(workbook, "Mensual_Flota", SUMMARY_BASE_FIELDS, mensual_flota_rows)
    excel_write_sheet(
        workbook,
        "Anual_Flota",
        [field for field in SUMMARY_BASE_FIELDS if field != "mes"],
        anual_flota_rows,
    )
    excel_write_sheet(workbook, "Turnos_Mensual", SHIFT_MONTHLY_FIELDNAMES, shift_mensual_rows)
    excel_write_sheet(workbook, "Impacto_Disp_Cod", IMPACT_FIELDNAMES, impacto_disp_rows)
    excel_write_sheet(workbook, "Impacto_UEBD_Cod", IMPACT_FIELDNAMES, impacto_uebd_rows)
    excel_write_sheet(
        workbook,
        "Impacto_Disp_Cod_Rig",
        IMPACT_BY_RIG_FIELDNAMES,
        impacto_disp_rig_rows,
    )
    excel_write_sheet(
        workbook,
        "Impacto_UEBD_Cod_Rig",
        IMPACT_BY_RIG_FIELDNAMES,
        impacto_uebd_rig_rows,
    )
    excel_write_sheet(workbook, "Top_Dias_Disp", TOP_DIAS_FIELDNAMES, top_dias_disp_rows)
    excel_write_sheet(workbook, "Top_Dias_UEBD", TOP_DIAS_FIELDNAMES, top_dias_uebd_rows)

    if len(resumen_ejecutivo_rows) > 1:
        last_row = len(resumen_ejecutivo_rows) - 1
        cat_col = RESUMEN_EJECUTIVO_FIELDNAMES.index("perforadora")
        disp_col = RESUMEN_EJECUTIVO_FIELDNAMES.index("disponibilidad_pct")
        uebd_col = RESUMEN_EJECUTIVO_FIELDNAMES.index("uebd_pct")

        chart = workbook.add_chart({"type": "column"})
        chart.add_series(
            {
                "name": "Disponibilidad %",
                "categories": ["Resumen_Ejecutivo", 1, cat_col, last_row, cat_col],
                "values": ["Resumen_Ejecutivo", 1, disp_col, last_row, disp_col],
            }
        )
        chart.add_series(
            {
                "name": "UEBD %",
                "categories": ["Resumen_Ejecutivo", 1, cat_col, last_row, cat_col],
                "values": ["Resumen_Ejecutivo", 1, uebd_col, last_row, uebd_col],
            }
        )
        chart.set_title({"name": "KPI por perforadora (%)"})
        chart.set_y_axis({"name": "%"})
        chart.set_legend({"position": "bottom"})
        ws_resumen.insert_chart("Q2", chart, {"x_scale": 1.2, "y_scale": 1.2})

    if mensual_flota_rows:
        last_row = len(mensual_flota_rows)
        month_col = SUMMARY_BASE_FIELDS.index("mes")
        disp_col = SUMMARY_BASE_FIELDS.index("disponibilidad_pct")
        uebd_col = SUMMARY_BASE_FIELDS.index("uebd_pct")

        chart_line = workbook.add_chart({"type": "line"})
        chart_line.add_series(
            {
                "name": "Disponibilidad %",
                "categories": ["Mensual_Flota", 1, month_col, last_row, month_col],
                "values": ["Mensual_Flota", 1, disp_col, last_row, disp_col],
            }
        )
        chart_line.add_series(
            {
                "name": "UEBD %",
                "categories": ["Mensual_Flota", 1, month_col, last_row, month_col],
                "values": ["Mensual_Flota", 1, uebd_col, last_row, uebd_col],
            }
        )
        chart_line.set_title({"name": "Tendencia mensual flota (%)"})
        chart_line.set_y_axis({"name": "%"})
        chart_line.set_legend({"position": "bottom"})
        ws_mensual_flota.insert_chart("Q2", chart_line, {"x_scale": 1.2, "y_scale": 1.2})

    if waterfall_disp_png is not None and waterfall_disp_png.exists():
        ws_mensual_rig.insert_image("Q2", str(waterfall_disp_png))
    if waterfall_uebd_png is not None and waterfall_uebd_png.exists():
        ws_mensual_rig.insert_image("Q28", str(waterfall_uebd_png))

    workbook.close()
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
    parser.add_argument(
        "--top-dias-criticos",
        type=int,
        default=20,
        help="Cantidad de dias criticos a listar para disponibilidad y UEBD.",
    )
    parser.add_argument(
        "--top-n-codigos-rig",
        type=int,
        default=5,
        help="Top de codigos por perforadora en impactos por metrica.",
    )
    parser.add_argument(
        "--sin-excel",
        action="store_true",
        help="No generar libro Excel extensivo.",
    )
    parser.add_argument(
        "--excel-filename",
        default="reporte_extensivo_tiempos_operativos.xlsx",
        help="Nombre del archivo Excel extensivo de salida.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_csv: Path = args.input_csv
    output_dir: Path = args.output_dir
    top_n_codigos = max(int(args.top_n_codigos), 1)
    top_dias_criticos = max(int(args.top_dias_criticos), 1)
    top_n_codigos_rig = max(int(args.top_n_codigos_rig), 1)

    if not input_csv.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {input_csv}")

    daily_rows, shift_daily_rows, stats, impact_data = load_daily_metrics(
        input_csv=input_csv,
        delimiter=args.delimiter,
        encoding=args.encoding,
    )
    monthly_by_rig = aggregate_period(daily_rows, granularity="mensual", by_rig=True)
    yearly_by_rig = aggregate_period(daily_rows, granularity="anual", by_rig=True)
    monthly_fleet = aggregate_period(daily_rows, granularity="mensual", by_rig=False)
    yearly_fleet = aggregate_period(daily_rows, granularity="anual", by_rig=False)
    shift_monthly = aggregate_shift_monthly(shift_daily_rows)
    resumen_ejecutivo_rows = aggregate_all_period_by_rig(daily_rows)

    monthly_fields = SUMMARY_BASE_FIELDS.copy()
    yearly_fields = [field for field in SUMMARY_BASE_FIELDS if field != "mes"]

    write_csv(output_dir / "diario_por_perforadora.csv", daily_rows, DAILY_FIELDNAMES)
    write_csv(
        output_dir / "diario_por_perforadora_y_turno.csv",
        shift_daily_rows,
        SHIFT_DAILY_FIELDNAMES,
    )
    write_csv(output_dir / "mensual_por_perforadora.csv", monthly_by_rig, monthly_fields)
    write_csv(output_dir / "anual_por_perforadora.csv", yearly_by_rig, yearly_fields)
    write_csv(output_dir / "mensual_flota.csv", monthly_fleet, monthly_fields)
    write_csv(output_dir / "anual_flota.csv", yearly_fleet, yearly_fields)
    write_csv(
        output_dir / "mensual_por_perforadora_y_turno.csv",
        shift_monthly,
        SHIFT_MONTHLY_FIELDNAMES,
    )
    write_csv(
        output_dir / "resumen_ejecutivo_total_periodo.csv",
        resumen_ejecutivo_rows,
        RESUMEN_EJECUTIVO_FIELDNAMES,
    )

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

    denominator_total_by_rig = {
        str(row["perforadora"]): float(row["horas_totales"])
        for row in resumen_ejecutivo_rows
        if str(row["perforadora"]) != "TOTAL"
    }
    denominator_oper_by_rig = {
        str(row["perforadora"]): float(row["horas_operativas"])
        for row in resumen_ejecutivo_rows
        if str(row["perforadora"]) != "TOTAL"
    }

    availability_impact_rows_by_rig = build_impact_rows_by_rig(
        metrica="disponibilidad",
        hours_by_rig_code=impact_data["availability_impact_hours_by_code_rig"],
        denominator_by_rig=denominator_total_by_rig,
        top_n_per_rig=top_n_codigos_rig,
    )
    uebd_impact_rows_by_rig = build_impact_rows_by_rig(
        metrica="uebd",
        hours_by_rig_code=impact_data["uebd_impact_hours_by_code_rig"],
        denominator_by_rig=denominator_oper_by_rig,
        top_n_per_rig=top_n_codigos_rig,
    )
    write_csv(
        output_dir / "impacto_codigos_disponibilidad_por_perforadora.csv",
        availability_impact_rows_by_rig,
        IMPACT_BY_RIG_FIELDNAMES,
    )
    write_csv(
        output_dir / "impacto_codigos_uebd_por_perforadora.csv",
        uebd_impact_rows_by_rig,
        IMPACT_BY_RIG_FIELDNAMES,
    )

    top_dias_disp = build_top_critical_days(
        daily_rows,
        metric_key="disponibilidad_ratio",
        metric_name="disponibilidad",
        top_n=top_dias_criticos,
    )
    top_dias_uebd = build_top_critical_days(
        daily_rows,
        metric_key="uebd_ratio",
        metric_name="uebd",
        top_n=top_dias_criticos,
    )
    write_csv(
        output_dir / "top_dias_criticos_disponibilidad.csv",
        top_dias_disp,
        TOP_DIAS_FIELDNAMES,
    )
    write_csv(
        output_dir / "top_dias_criticos_uebd.csv",
        top_dias_uebd,
        TOP_DIAS_FIELDNAMES,
    )

    chart_messages: List[str] = []
    waterfall_disp_path = output_dir / "graficos" / "cascada_disponibilidad_top_codigos.png"
    waterfall_uebd_path = output_dir / "graficos" / "cascada_uebd_top_codigos.png"
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
            output_path=waterfall_disp_path,
            title="Cascada de codigos que reducen Disponibilidad",
            subtitle=f"Top {top_n_codigos} + Otros | Base: 100%",
            contributions=availability_contrib,
            final_ratio=disponibilidad_ratio,
        )
        ok_uebd, err_uebd = generate_waterfall_chart(
            output_path=waterfall_uebd_path,
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

    excel_messages: List[str] = []
    if not args.sin_excel:
        excel_path = output_dir / args.excel_filename
        ok_excel, err_excel = write_excel_report(
            output_path=excel_path,
            resumen_ejecutivo_rows=resumen_ejecutivo_rows,
            diario_rows=daily_rows,
            mensual_por_rig_rows=monthly_by_rig,
            anual_por_rig_rows=yearly_by_rig,
            mensual_flota_rows=monthly_fleet,
            anual_flota_rows=yearly_fleet,
            shift_mensual_rows=shift_monthly,
            impacto_disp_rows=availability_impact_rows,
            impacto_uebd_rows=uebd_impact_rows,
            impacto_disp_rig_rows=availability_impact_rows_by_rig,
            impacto_uebd_rig_rows=uebd_impact_rows_by_rig,
            top_dias_disp_rows=top_dias_disp,
            top_dias_uebd_rows=top_dias_uebd,
            waterfall_disp_png=waterfall_disp_path if waterfall_disp_path.exists() else None,
            waterfall_uebd_png=waterfall_uebd_path if waterfall_uebd_path.exists() else None,
        )
        if ok_excel:
            excel_messages.append(f"Excel extensivo: generado ({excel_path})")
        else:
            excel_messages.append(f"Excel extensivo: omitido ({err_excel})")
    else:
        excel_messages.append("Excel extensivo deshabilitado por parametro.")

    print("Analisis finalizado.")
    print(f"Archivo de entrada: {input_csv}")
    print(f"Carpeta de salida: {output_dir}")
    print(f"Filas leidas: {stats['rows_total']}")
    print(f"Filas sin dia operativo: {stats['rows_without_operational_day']}")
    print(f"Filas con fallback de duracion: {stats['rows_duration_fallback']}")
    print(f"Registros diarios (dia+perforadora): {len(daily_rows)}")
    print(f"Registros diarios (dia+perforadora+turno): {len(shift_daily_rows)}")
    print(f"Resumen mensual por perforadora: {len(monthly_by_rig)} filas")
    print(f"Resumen anual por perforadora: {len(yearly_by_rig)} filas")
    print(f"Resumen mensual por perforadora+turno: {len(shift_monthly)} filas")
    print(f"Resumen ejecutivo total periodo: {len(resumen_ejecutivo_rows)} filas")
    print(f"Impacto codigos disponibilidad: {len(availability_impact_rows)} filas")
    print(f"Impacto codigos UEBD: {len(uebd_impact_rows)} filas")
    print(f"Impacto codigos disponibilidad por perforadora: {len(availability_impact_rows_by_rig)} filas")
    print(f"Impacto codigos UEBD por perforadora: {len(uebd_impact_rows_by_rig)} filas")
    print(f"Top dias criticos disponibilidad: {len(top_dias_disp)} filas")
    print(f"Top dias criticos UEBD: {len(top_dias_uebd)} filas")
    for msg in chart_messages:
        print(msg)
    for msg in excel_messages:
        print(msg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
