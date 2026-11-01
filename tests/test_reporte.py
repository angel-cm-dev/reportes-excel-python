from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from reportes_ventas import Venta, agregar, generar_reporte, leer_ventas
from reportes_ventas.cli import main
from reportes_ventas.datos import etiqueta_mes
from reportes_ventas.excel import FORMATO_SOLES

EJEMPLO = Path(__file__).resolve().parent.parent / "ejemplos" / "ventas.csv"


def _venta(fecha, vendedor, producto, cantidad, precio):
    return Venta(date.fromisoformat(fecha), vendedor, producto, cantidad, Decimal(precio))


@pytest.fixture
def ventas():
    return [
        _venta("2026-01-05", "Ana", "Laptop", 1, "3000.00"),
        _venta("2026-01-20", "Luis", "Mouse", 3, "50.00"),
        _venta("2026-02-02", "Ana", "Mouse", 2, "50.00"),
        _venta("2026-02-15", "Luis", "Monitor", 1, "700.50"),
    ]


@pytest.fixture
def libro(ventas, tmp_path):
    ruta = generar_reporte(ventas, tmp_path / "reporte.xlsx", titulo="Ventas de prueba")
    return load_workbook(ruta)


def test_leer_csv_de_ejemplo():
    ventas = leer_ventas(EJEMPLO)
    assert len(ventas) == 72
    assert all(v.cantidad > 0 for v in ventas)


def test_leer_csv_rechaza_columnas_faltantes(tmp_path):
    ruta = tmp_path / "malo.csv"
    ruta.write_text("fecha,vendedor\n2026-01-01,Ana\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Faltan columnas"):
        leer_ventas(ruta)


def test_leer_csv_indica_linea_invalida(tmp_path):
    ruta = tmp_path / "malo.csv"
    ruta.write_text(
        "fecha,vendedor,producto,cantidad,precio_unitario\n"
        "2026-01-01,Ana,Mouse,1,50\n"
        "2026-13-01,Ana,Mouse,1,50\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="Línea 3"):
        leer_ventas(ruta)


def test_agregar_por_mes_vendedor_y_producto(ventas):
    meses = agregar(ventas, "mes")
    assert [(m.clave, m.total) for m in meses] == [("2026-01", Decimal("3150.00")), ("2026-02", Decimal("800.50"))]

    vendedores = agregar(ventas, "vendedor")
    assert vendedores[0].clave == "Ana" and vendedores[0].total == Decimal("3100.00")
    assert vendedores[1].ticket_promedio == Decimal("425.25")

    productos = agregar(ventas, "producto")
    mouse = next(p for p in productos if p.clave == "Mouse")
    assert (mouse.transacciones, mouse.unidades, mouse.total) == (2, 5, Decimal("250.00"))


def test_etiqueta_mes():
    assert etiqueta_mes("2026-09") == "Set 2026"


def test_libro_tiene_hojas_esperadas(libro):
    assert libro.sheetnames == ["Resumen", "Por vendedor", "Por producto", "Detalle"]


def test_resumen_kpis_y_formato_moneda(libro):
    hoja = libro["Resumen"]
    assert hoja["A1"].value == "Ventas de prueba"
    assert hoja["A2"].value == "Periodo: 05/01/2026 al 15/02/2026"
    assert hoja["B4"].value == pytest.approx(3950.50)
    assert hoja["B4"].number_format == FORMATO_SOLES
    assert hoja["B5"].value == 4
    assert hoja["B8"].value == "Ana"       # mejor vendedor
    assert hoja["B9"].value == "Laptop"    # producto top


def test_resumen_tabla_mensual_con_totales_y_grafico(libro):
    hoja = libro["Resumen"]
    filas = list(hoja.iter_rows(min_row=12, max_row=15, max_col=4, values_only=True))
    assert filas[0] == ("Mes", "Transacciones", "Unidades", "Total vendido")
    assert filas[1] == ("Ene 2026", 2, 4, 3150.0)
    assert filas[2] == ("Feb 2026", 2, 3, 800.5)
    assert filas[3] == ("Total", "=SUM(B13:B14)", "=SUM(C13:C14)", "=SUM(D13:D14)")
    assert hoja["D13"].number_format == FORMATO_SOLES
    assert len(hoja._charts) == 1
    assert hoja._charts[0].type == "col"


def test_hojas_agregadas(libro):
    vendedores = libro["Por vendedor"]
    assert vendedores["A2"].value == "Ana"
    assert vendedores["D2"].value == pytest.approx(3100.0)
    assert vendedores["F2"].value == pytest.approx(3100 / 3950.5)
    assert vendedores["F2"].number_format == "0.0%"
    assert vendedores["A4"].value == "Total"
    assert vendedores.freeze_panes == "A2"
    assert len(vendedores._charts) == 1
    assert len(libro["Por producto"]._charts) == 1


def test_detalle_con_formulas_y_filtro(libro):
    hoja = libro["Detalle"]
    assert hoja.max_row == 6  # encabezado + 4 ventas + total
    assert hoja["F2"].value == "=D2*E2"
    assert hoja["F6"].value == "=SUM(F2:F5)"
    assert hoja["A2"].number_format == "dd/mm/yyyy"
    assert hoja.auto_filter.ref == "A1:F5"
    assert hoja["A1"].font.bold


def test_generar_sin_ventas_falla(tmp_path):
    with pytest.raises(ValueError):
        generar_reporte([], tmp_path / "vacio.xlsx")


def test_cli_genera_archivo(tmp_path, capsys):
    salida = tmp_path / "sub" / "reporte.xlsx"
    assert main([str(EJEMPLO), "-o", str(salida)]) == 0
    assert salida.exists()
    assert "72 ventas procesadas" in capsys.readouterr().out
    libro = load_workbook(salida)
    assert libro["Detalle"].max_row == 74


def test_cli_archivo_inexistente(tmp_path, capsys):
    assert main([str(tmp_path / "no.csv")]) == 1
    assert "no existe" in capsys.readouterr().err
