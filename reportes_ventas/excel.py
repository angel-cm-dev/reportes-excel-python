"""Construcción del libro Excel con openpyxl."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .datos import Agregado, Venta, agregar, etiqueta_mes

FORMATO_SOLES = '"S/" #,##0.00'
FORMATO_ENTERO = "#,##0"
FORMATO_PORCENTAJE = "0.0%"
FORMATO_FECHA = "dd/mm/yyyy"

COLOR_PRIMARIO = "1F4E78"
RELLENO_ENCABEZADO = PatternFill("solid", fgColor=COLOR_PRIMARIO)
RELLENO_TOTAL = PatternFill("solid", fgColor="DDEBF7")
FUENTE_ENCABEZADO = Font(bold=True, color="FFFFFF")
BORDE_SUPERIOR = Border(top=Side(style="thin", color=COLOR_PRIMARIO))


def _monto(valor: Decimal) -> float:
    return float(valor.quantize(Decimal("0.01")))


def _escribir_tabla(
    hoja: Worksheet,
    fila_inicio: int,
    encabezados: list[str],
    filas: list[list],
    formatos: list[str | None],
    columnas_suma: tuple[int, ...] = (),
) -> tuple[int, int]:
    """Escribe encabezado, datos y fila de totales (fórmulas SUM). Devuelve (primera, última) fila de datos."""
    for col, texto in enumerate(encabezados, start=1):
        celda = hoja.cell(row=fila_inicio, column=col, value=texto)
        celda.fill = RELLENO_ENCABEZADO
        celda.font = FUENTE_ENCABEZADO
        celda.alignment = Alignment(horizontal="center", vertical="center")

    primera = fila_inicio + 1
    for desplazamiento, valores in enumerate(filas):
        for col, (valor, formato) in enumerate(zip(valores, formatos), start=1):
            celda = hoja.cell(row=primera + desplazamiento, column=col, value=valor)
            if formato:
                celda.number_format = formato
    ultima = primera + len(filas) - 1

    if columnas_suma:
        fila_total = ultima + 1
        hoja.cell(row=fila_total, column=1, value="Total")
        for col in columnas_suma:
            letra = get_column_letter(col)
            celda = hoja.cell(row=fila_total, column=col, value=f"=SUM({letra}{primera}:{letra}{ultima})")
            celda.number_format = formatos[col - 1] or "General"
        for col in range(1, len(encabezados) + 1):
            celda = hoja.cell(row=fila_total, column=col)
            celda.font = Font(bold=True)
            celda.fill = RELLENO_TOTAL
            celda.border = BORDE_SUPERIOR
    return primera, ultima


def _ajustar_anchos(hoja: Worksheet, minimo: int = 10, maximo: int = 40) -> None:
    for columna in hoja.columns:
        largo = max((len(str(c.value)) for c in columna if c.value is not None and not str(c.value).startswith("=")), default=0)
        hoja.column_dimensions[columna[0].column_letter].width = max(minimo, min(maximo, largo + 3))


def _grafico_barras(hoja: Worksheet, titulo: str, col_categoria: int, col_valor: int,
                    fila_encabezado: int, ultima: int, ancla: str) -> None:
    grafico = BarChart()
    grafico.type = "col"
    grafico.title = titulo
    grafico.y_axis.title = "Soles (S/)"
    grafico.y_axis.numFmt = FORMATO_SOLES
    grafico.legend = None
    grafico.height, grafico.width = 8, 16
    datos = Reference(hoja, min_col=col_valor, min_row=fila_encabezado, max_row=ultima)
    categorias = Reference(hoja, min_col=col_categoria, min_row=fila_encabezado + 1, max_row=ultima)
    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    hoja.add_chart(grafico, ancla)


def _hoja_agregada(libro: Workbook, nombre: str, etiqueta: str, grupos: list[Agregado], total: Decimal) -> None:
    hoja = libro.create_sheet(nombre)
    encabezados = [etiqueta, "Transacciones", "Unidades", "Total vendido", "Ticket promedio", "% del total"]
    filas = [
        [g.clave, g.transacciones, g.unidades, _monto(g.total), _monto(g.ticket_promedio),
         float(g.total / total) if total else 0.0]
        for g in grupos
    ]
    formatos = [None, FORMATO_ENTERO, FORMATO_ENTERO, FORMATO_SOLES, FORMATO_SOLES, FORMATO_PORCENTAJE]
    _, ultima = _escribir_tabla(hoja, 1, encabezados, filas, formatos, columnas_suma=(2, 3, 4, 6))
    hoja.freeze_panes = "A2"
    _ajustar_anchos(hoja)
    _grafico_barras(hoja, f"Ventas por {etiqueta.lower()}", 1, 4, 1, ultima, "H2")


def generar_reporte(ventas: list[Venta], ruta: str | Path, titulo: str = "Reporte de ventas") -> Path:
    if not ventas:
        raise ValueError("No hay ventas para generar el reporte")
    ruta = Path(ruta)
    total = sum((v.importe for v in ventas), Decimal("0"))
    por_mes = agregar(ventas, "mes")
    por_vendedor = agregar(ventas, "vendedor")
    por_producto = agregar(ventas, "producto")

    libro = Workbook()
    resumen = libro.active
    resumen.title = "Resumen"

    # Encabezado y KPIs
    resumen["A1"] = titulo
    resumen["A1"].font = Font(bold=True, size=16, color=COLOR_PRIMARIO)
    inicio, fin = min(v.fecha for v in ventas), max(v.fecha for v in ventas)
    resumen["A2"] = f"Periodo: {inicio:%d/%m/%Y} al {fin:%d/%m/%Y}"
    resumen["A2"].font = Font(italic=True, color="595959")

    kpis = [
        ("Total vendido", _monto(total), FORMATO_SOLES),
        ("Transacciones", len(ventas), FORMATO_ENTERO),
        ("Unidades vendidas", sum(v.cantidad for v in ventas), FORMATO_ENTERO),
        ("Ticket promedio", _monto(total / len(ventas)), FORMATO_SOLES),
        ("Mejor vendedor", por_vendedor[0].clave, None),
        ("Producto top", por_producto[0].clave, None),
    ]
    for fila, (nombre, valor, formato) in enumerate(kpis, start=4):
        resumen.cell(row=fila, column=1, value=nombre).font = Font(bold=True)
        celda = resumen.cell(row=fila, column=2, value=valor)
        if formato:
            celda.number_format = formato

    # Tabla mensual con gráfico
    fila_tabla = 4 + len(kpis) + 2
    resumen.cell(row=fila_tabla - 1, column=1, value="Ventas por mes").font = Font(bold=True, size=12)
    filas_mes = [[etiqueta_mes(g.clave), g.transacciones, g.unidades, _monto(g.total)] for g in por_mes]
    _, ultima = _escribir_tabla(
        resumen, fila_tabla, ["Mes", "Transacciones", "Unidades", "Total vendido"], filas_mes,
        [None, FORMATO_ENTERO, FORMATO_ENTERO, FORMATO_SOLES], columnas_suma=(2, 3, 4),
    )
    _ajustar_anchos(resumen)
    resumen.column_dimensions["A"].width = 22
    _grafico_barras(resumen, "Ventas por mes", 1, 4, fila_tabla, ultima, "F4")

    _hoja_agregada(libro, "Por vendedor", "Vendedor", por_vendedor, total)
    _hoja_agregada(libro, "Por producto", "Producto", por_producto, total)

    # Detalle con importe como fórmula para que el usuario pueda editar cantidades
    detalle = libro.create_sheet("Detalle")
    filas = [
        [v.fecha, v.vendedor, v.producto, v.cantidad, _monto(v.precio_unitario), f"=D{i}*E{i}"]
        for i, v in enumerate(ventas, start=2)
    ]
    _, ultima = _escribir_tabla(
        detalle, 1, ["Fecha", "Vendedor", "Producto", "Cantidad", "Precio unitario", "Importe"], filas,
        [FORMATO_FECHA, None, None, FORMATO_ENTERO, FORMATO_SOLES, FORMATO_SOLES], columnas_suma=(4, 6),
    )
    detalle.auto_filter.ref = f"A1:F{ultima}"
    detalle.freeze_panes = "A2"
    _ajustar_anchos(detalle)

    ruta.parent.mkdir(parents=True, exist_ok=True)
    libro.save(ruta)
    return ruta
