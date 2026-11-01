"""CLI para generar el reporte desde un CSV."""

from __future__ import annotations

import argparse
import sys
from decimal import Decimal
from pathlib import Path

from . import __version__
from .datos import leer_ventas
from .excel import generar_reporte


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="reporte-ventas",
        description="Genera un reporte de ventas en Excel (por mes, vendedor y producto) a partir de un CSV.",
    )
    parser.add_argument("entrada", type=Path, help="CSV con columnas fecha,vendedor,producto,cantidad,precio_unitario")
    parser.add_argument("-o", "--salida", type=Path, default=Path("salida/reporte_ventas.xlsx"),
                        help="Archivo .xlsx de salida (por defecto salida/reporte_ventas.xlsx)")
    parser.add_argument("-t", "--titulo", default="Reporte de ventas", help="Título del reporte")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    try:
        ventas = leer_ventas(args.entrada)
        ruta = generar_reporte(ventas, args.salida, args.titulo)
    except FileNotFoundError:
        print(f"Error: no existe el archivo {args.entrada}", file=sys.stderr)
        return 1
    except ValueError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    total = sum((v.importe for v in ventas), Decimal("0"))
    print(f"{len(ventas)} ventas procesadas | Total: S/ {total:,.2f}")
    print(f"Reporte generado: {ruta}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
