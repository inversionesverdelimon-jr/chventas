import datetime as dt

import pandas as pd
import pytest

from chventas.lector import (Configuracion, a_fecha, a_numero, cargar, mapear_columnas,
                             resolver_archivos)


@pytest.mark.parametrize("entrada,esperado", [
    ("1.234.567", 1234567), ("1.234.567,89", 1234567.89), ("1,234,567.89", 1234567.89),
    ("$ -294.078.645", -294078645), ("(1.000)", -1000), (12.5, 12.5), (None, 0), ("", 0),
    ("3,5", 3.5),
])
def test_a_numero(entrada, esperado):
    assert a_numero(entrada) == pytest.approx(esperado)


@pytest.mark.parametrize("entrada,esperado", [
    ("23/11/2026", dt.date(2026, 11, 23)), ("2026-11-03", dt.date(2026, 11, 3)),
    (dt.datetime(2025, 1, 2, 10), dt.date(2025, 1, 2)), (46000, dt.date(2025, 12, 9)),
    ("Total cuenta", None), (None, None),
])
def test_a_fecha(entrada, esperado):
    assert a_fecha(entrada) == esperado


def test_mapear_columnas_auxiliar():
    m = mapear_columnas(["Cuenta", "Nombre cuenta", "Fecha", "NIT", "Nombre tercero", "Débito", "Crédito"])
    assert m == {"cuenta": 0, "fecha": 2, "cliente_id": 3, "cliente_nombre": 4, "debito": 5, "credito": 6}


def _guardar(tmp_path, filas, nombre="base.xlsx"):
    ruta = tmp_path / nombre
    pd.DataFrame(filas).to_excel(ruta, header=False, index=False)
    return ruta


def test_cargar_contable(tmp_path):
    _guardar(tmp_path, [
        ["EMPRESA DEMO", None, None, None, None, None],
        ["Cuenta", "Fecha", "NIT", "Nombre tercero", "Débito", "Crédito"],
        ["41352001", "05/01/2026", 900, "CLIENTE A", 0, "1.000.000"],
        ["41750101", "06/01/2026", 900, "CLIENTE A", "200.000", 0],
        ["41352001", "07/02/2025", 901, "CLIENTE B", 0, 500000],
        ["42100501", "07/02/2025", 901, "CLIENTE B", 0, 999],  # no operacional: se ignora
        ["4135", None, None, "Total", 0, 1500000],              # subtotal sin fecha
    ])
    # Ruta sin extensión, como en config.json
    p = cargar(Configuracion(ruta_excel=str(tmp_path / "base")))
    assert p["fecha_min"] == 20250207 and p["fecha_max"] == 20260106
    assert p["clientes"] == ["900 - CLIENTE A", "901 - CLIENTE B"]
    regs = {(r[0], r[1]): (r[2], r[3]) for r in p["registros"]}
    assert regs[(20260105, 0)] == (1000000, 0)
    assert regs[(20260106, 0)] == (0, 200000)
    assert regs[(20250207, 1)] == (500000, 0)
    assert p["fuente"]["hojas"][0]["modo"] == "contable"


def test_cargar_valor_con_tipo_y_carpeta(tmp_path):
    _guardar(tmp_path, [
        ["Fecha", "Cliente", "Tipo", "Valor"],
        [dt.datetime(2026, 3, 1), "811044814 - GIRDLE SAS", "Factura", 1000],
        [dt.datetime(2026, 3, 2), "811044814 - GIRDLE SAS", "Devolución", 300],
    ])
    p = cargar(Configuracion(ruta_excel=str(tmp_path)))  # carpeta
    assert p["clientes"] == ["811044814 - GIRDLE SAS"]
    assert sorted(r[2:] for r in p["registros"]) == [[0, 300], [1000, 0]]


def test_ruta_inexistente(tmp_path):
    with pytest.raises(FileNotFoundError):
        resolver_archivos(str(tmp_path / "no existe"))
