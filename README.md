# reportes-excel-python

Generador de reportes de ventas en Excel con openpyxl: agrega un CSV por mes, vendedor y producto, con formato en soles y gráficos de barras.

![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)
![Licencia MIT](https://img.shields.io/badge/licencia-MIT-green)
![CI](https://github.com/angel-cm-dev/reportes-excel-python/actions/workflows/ci.yml/badge.svg)

## Características

- Lee un CSV de ventas y **valida cada fila** (fecha ISO, cantidad > 0, precio ≥ 0), indicando la línea con error.
- Cálculos con `Decimal` para evitar errores de redondeo.
- Libro con 4 hojas:
  - **Resumen**: KPIs (total vendido, transacciones, unidades, ticket promedio, mejor vendedor, producto top), tabla mensual y **gráfico de barras**.
  - **Por vendedor** y **Por producto**: transacciones, unidades, total, ticket promedio, % del total y gráfico.
  - **Detalle**: todas las ventas con el importe como fórmula (`=D2*E2`), autofiltro y paneles inmovilizados.
- Formato moneda `S/ #,##0.00`, porcentajes, fechas `dd/mm/aaaa`, encabezados con color y filas de **totales con `SUM`**.
- Anchos de columna automáticos.

## Requisitos

- Python 3.10 o superior
- `openpyxl` 3.1+ (y `pytest` para las pruebas)

## Instalación y uso

```bash
git clone https://github.com/angel-cm-dev/reportes-excel-python.git
cd reportes-excel-python
python -m venv .venv
# Windows: .venv\Scripts\activate   |   Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt

python -m reportes_ventas ejemplos/ventas.csv -o salida/reporte_ventas.xlsx -t "Ventas 1er semestre 2026"
```

Formato del CSV de entrada:

```csv
fecha,vendedor,producto,cantidad,precio_unitario
2026-01-03,Carmen Villanueva,Audífonos BT,3,159.00
2026-01-10,Jorge Ríos,Laptop 14 pulgadas,1,2899.00
```

Uso como librería:

```python
from reportes_ventas import leer_ventas, generar_reporte

ventas = leer_ventas("ejemplos/ventas.csv")
generar_reporte(ventas, "salida/reporte.xlsx", titulo="Ventas Q1")
```

## Ejemplo de salida

```text
$ python -m reportes_ventas ejemplos/ventas.csv
72 ventas procesadas | Total: S/ 120,962.30
Reporte generado: salida/reporte_ventas.xlsx
```

Hoja **Resumen** (extracto):

| Mes | Transacciones | Unidades | Total vendido |
| --- | ---: | ---: | ---: |
| Ene 2026 | 12 | 29 | S/ 13,312.70 |
| Feb 2026 | 12 | 39 | S/ 25,698.10 |
| Mar 2026 | 12 | 33 | S/ 33,371.30 |
| … | | | |
| **Total** | **72** | **206** | **S/ 120,962.30** |

> Los datos de `ejemplos/ventas.csv` son ficticios.

## Estructura del proyecto

```text
reportes-excel-python/
├── reportes_ventas/
│   ├── __init__.py
│   ├── __main__.py     # permite `python -m reportes_ventas`
│   ├── cli.py          # argumentos y mensajes
│   ├── datos.py        # lectura/validación del CSV y agregaciones
│   └── excel.py        # hojas, estilos, totales y gráficos
├── ejemplos/ventas.csv
├── tests/
├── .github/workflows/ci.yml
├── pyproject.toml
└── requirements.txt
```

## Pruebas

```bash
python -m pytest -v
```

Las pruebas generan el `.xlsx` en un directorio temporal y lo vuelven a abrir con openpyxl para verificar hojas, valores, fórmulas, formatos y gráficos.

## Roadmap

- [ ] Filtros por rango de fechas y vendedor desde la CLI.
- [ ] Comparativo contra metas mensuales con formato condicional.
- [ ] Envío automático del reporte por correo desde un flujo n8n.

## Licencia y autor

Distribuido bajo licencia [MIT](LICENSE).

Desarrollado por **Angel Castañeda Meza** — Ingeniero de Sistemas, Lima, Perú · [github.com/angel-cm-dev](https://github.com/angel-cm-dev)
