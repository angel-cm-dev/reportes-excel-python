"""Generador de reportes de ventas en Excel."""

from .datos import Agregado, Venta, agregar, leer_ventas
from .excel import generar_reporte

__all__ = ["Agregado", "Venta", "agregar", "generar_reporte", "leer_ventas"]
__version__ = "1.0.0"
