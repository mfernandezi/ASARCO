#!/usr/bin/env python3
"""Calcula metricas operativas (UEBD y Disponibilidad) desde un CSV base.

Formulas implementadas:
    UEBD = (Horas Efectivas / Horas Operativas) / 100
    Disponibilidad = (Horas Operativas / Horas Totales) / 100
"""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Calcula Horas Efectivas/Operativas/Totales, UEBD y Disponibilidad."
    )
    parser.add_argument(
        "input_csv",
        type=Path,
        nargs="?",
        default=Path("DispUEBD_AllRigs_010126-0000_170226-2100.csv"),
        help="Ruta del CSV de entrada (default: archivo de la carpeta actual).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("metricas_operativas_por_rig.csv"),
        help="Ruta del CSV de salida.",
    )
    parser.add_argument(
        "--group-by",
        nargs="+",
        default=["RigName"],
        help="Columnas para agrupar (default: RigName).",
    )
    parser.add_argument(
        "--effective-shortcode",
        default="Efectivo",
        help="Valor de ShortCode considerado como hora efectiva.",
    )
    parser.add_argument(
        "--operative-shortcodes",
        nargs="+",
        default=["Efectivo", "Demora"],
        help="Valores de ShortCode considerados como horas operativas.",
    )
    parser.add_argument(
        "--hours-decimals",
        type=int,
        default=4,
        help="Cantidad de decimales para columnas de horas.",
    )
    parser.add_argument(
        "--ratio-decimals",
        type=int,
        default=8,
        help="Cantidad de decimales para UEBD y Disponibilidad.",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8-sig",
        help="Codificacion de entrada/salida del CSV.",
    )
    parser.add_argument(
        "--delimiter",
        default=";",
        help="Delimitador del CSV (default: ';').",
    )
    parser.add_argument(
        "--no-total-row",
        action="store_true",
        help="Si se usa, no agrega fila de total general.",
    )
    return parser.parse_args()


def to_float(value: str | None) -> float:
    if value is None:
        return 0.0
    text = str(value).strip()
    if not text:
        return 0.0
    # Soporta coma o punto decimal en valores puntuales.
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def format_number(value: float, decimals: int) -> str:
    return f"{value:.{decimals}f}"


def ensure_required_columns(fieldnames: Sequence[str], required: Iterable[str]) -> None:
    missing = [column for column in required if column not in fieldnames]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {', '.join(missing)}")


def build_group_key(row: Dict[str, str], group_by: Sequence[str]) -> Tuple[str, ...]:
    return tuple((row.get(column) or "").strip() for column in group_by)


def main() -> None:
    args = parse_args()
    operative_shortcodes = set(args.operative_shortcodes)

    accumulators: Dict[Tuple[str, ...], Dict[str, float]] = defaultdict(
        lambda: {"Horas Efectivas": 0.0, "Horas Operativas": 0.0, "Horas Totales": 0.0}
    )

    with args.input_csv.open("r", encoding=args.encoding, newline="") as in_file:
        reader = csv.DictReader(in_file, delimiter=args.delimiter)
        if not reader.fieldnames:
            raise ValueError("El CSV de entrada no contiene encabezados.")

        required_columns = list(args.group_by) + ["Duration", "ShortCode"]
        ensure_required_columns(reader.fieldnames, required_columns)

        for row in reader:
            key = build_group_key(row, args.group_by)
            duration_hours = to_float(row.get("Duration")) / 3600.0
            short_code = (row.get("ShortCode") or "").strip()

            accumulators[key]["Horas Totales"] += duration_hours
            if short_code == args.effective_shortcode:
                accumulators[key]["Horas Efectivas"] += duration_hours
            if short_code in operative_shortcodes:
                accumulators[key]["Horas Operativas"] += duration_hours

    output_rows: List[Dict[str, str]] = []
    totals = {"Horas Efectivas": 0.0, "Horas Operativas": 0.0, "Horas Totales": 0.0}

    for key in sorted(accumulators):
        values = accumulators[key]
        horas_efectivas = values["Horas Efectivas"]
        horas_operativas = values["Horas Operativas"]
        horas_totales = values["Horas Totales"]

        uebd = (horas_efectivas / horas_operativas) / 100.0 if horas_operativas else 0.0
        disponibilidad = (
            (horas_operativas / horas_totales) / 100.0 if horas_totales else 0.0
        )

        row = {column: key[idx] for idx, column in enumerate(args.group_by)}
        row.update(
            {
                "Horas Efectivas": format_number(horas_efectivas, args.hours_decimals),
                "Horas Operativas": format_number(horas_operativas, args.hours_decimals),
                "Horas Totales": format_number(horas_totales, args.hours_decimals),
                "UEBD": format_number(uebd, args.ratio_decimals),
                "Disponibilidad": format_number(disponibilidad, args.ratio_decimals),
            }
        )
        output_rows.append(row)

        totals["Horas Efectivas"] += horas_efectivas
        totals["Horas Operativas"] += horas_operativas
        totals["Horas Totales"] += horas_totales

    if not args.no_total_row:
        total_uebd = (
            (totals["Horas Efectivas"] / totals["Horas Operativas"]) / 100.0
            if totals["Horas Operativas"]
            else 0.0
        )
        total_disponibilidad = (
            (totals["Horas Operativas"] / totals["Horas Totales"]) / 100.0
            if totals["Horas Totales"]
            else 0.0
        )
        total_row = {column: ("TOTAL" if idx == 0 else "") for idx, column in enumerate(args.group_by)}
        total_row.update(
            {
                "Horas Efectivas": format_number(totals["Horas Efectivas"], args.hours_decimals),
                "Horas Operativas": format_number(totals["Horas Operativas"], args.hours_decimals),
                "Horas Totales": format_number(totals["Horas Totales"], args.hours_decimals),
                "UEBD": format_number(total_uebd, args.ratio_decimals),
                "Disponibilidad": format_number(total_disponibilidad, args.ratio_decimals),
            }
        )
        output_rows.append(total_row)

    output_columns = list(args.group_by) + [
        "Horas Efectivas",
        "Horas Operativas",
        "Horas Totales",
        "UEBD",
        "Disponibilidad",
    ]

    with args.output.open("w", encoding=args.encoding, newline="") as out_file:
        writer = csv.DictWriter(out_file, fieldnames=output_columns, delimiter=args.delimiter)
        writer.writeheader()
        writer.writerows(output_rows)

    print(f"Archivo generado: {args.output}")
    print(f"Filas generadas: {len(output_rows)}")


if __name__ == "__main__":
    main()
