"""Lectura y agregación de ventas desde CSV."""

from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

COLUMNAS_REQUERIDAS = ("fecha", "vendedor", "producto", "cantidad", "precio_unitario")
MESES_CORTOS = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic")


@dataclass(frozen=True)
class Venta:
    fecha: date
    vendedor: str
    producto: str
    cantidad: int
    precio_unitario: Decimal

    @property
    def importe(self) -> Decimal:
        return self.precio_unitario * self.cantidad

    @property
    def mes(self) -> str:
        return f"{self.fecha.year}-{self.fecha.month:02d}"


@dataclass
class Agregado:
    clave: str
    transacciones: int = 0
    unidades: int = 0
    total: Decimal = Decimal("0")

    @property
    def ticket_promedio(self) -> Decimal:
        return (self.total / self.transacciones).quantize(Decimal("0.01")) if self.transacciones else Decimal("0")


def etiqueta_mes(mes: str) -> str:
    """'2026-03' -> 'Mar 2026'."""
    anio, numero = mes.split("-")
    return f"{MESES_CORTOS[int(numero) - 1]} {anio}"


def leer_ventas(ruta: str | Path) -> list[Venta]:
    """Lee el CSV y valida cada fila; lanza ValueError indicando la línea con problemas."""
    with open(ruta, encoding="utf-8-sig", newline="") as archivo:
        lector = csv.DictReader(archivo)
        faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in (lector.fieldnames or [])]
        if faltantes:
            raise ValueError(f"Faltan columnas en el CSV: {', '.join(faltantes)}")

        ventas = []
        for numero_linea, fila in enumerate(lector, start=2):
            try:
                venta = Venta(
                    fecha=date.fromisoformat(fila["fecha"].strip()),
                    vendedor=fila["vendedor"].strip(),
                    producto=fila["producto"].strip(),
                    cantidad=int(fila["cantidad"]),
                    precio_unitario=Decimal(fila["precio_unitario"].strip()),
                )
            except (ValueError, InvalidOperation, AttributeError) as error:
                raise ValueError(f"Línea {numero_linea}: dato inválido ({error})") from error
            if venta.cantidad <= 0 or venta.precio_unitario < 0 or not venta.vendedor or not venta.producto:
                raise ValueError(f"Línea {numero_linea}: cantidad, precio, vendedor o producto inválidos")
            ventas.append(venta)
    return ventas


def agregar(ventas: list[Venta], campo: str) -> list[Agregado]:
    """Agrupa por 'mes', 'vendedor' o 'producto'. Meses en orden cronológico; el resto por total descendente."""
    grupos: dict[str, Agregado] = defaultdict(lambda: Agregado(""))
    for venta in ventas:
        clave = getattr(venta, campo)
        grupo = grupos[clave]
        grupo.clave = clave
        grupo.transacciones += 1
        grupo.unidades += venta.cantidad
        grupo.total += venta.importe
    if campo == "mes":
        return sorted(grupos.values(), key=lambda g: g.clave)
    return sorted(grupos.values(), key=lambda g: (-g.total, g.clave))
