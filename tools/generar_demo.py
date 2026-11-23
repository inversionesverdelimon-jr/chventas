"""Genera un Excel de demostración con formato de auxiliar contable (clase 4).

    python -m tools.generar_demo [ruta_salida]
"""

from __future__ import annotations

import datetime as dt
import os
import random
import sys

from openpyxl import Workbook
from openpyxl.styles import Font

NOMBRES = [
    "GIRDLE & LINGERIE S.A.S.", "LINGERIE LE SENSUEL SAS", "CACHARRERIA BOMBAY S.A.S",
    "MUJER LATINA DE COLOMBIA S.A.S", "RINPLAY S, DE R.L DE C.V.", "LINEA DIRECTA S.A.S",
    "C.I. FAJAS MYD COLOMBIA S.A.S", "CRUVERT IMPORT & EXPORT S.A.C", "ROPA INTIMA CAROL SAS",
    "PROINTIMO S.A.S", "SANTIAGO URIBE GRACIANO", "TEXTILES SWANTEX S.A.S",
    "GRUPO CARVACARO S.A.S", "AURATEX SAS", "CONFECCIONES CALAR S.A.S", "ODEA S.A.S",
    "MARKETING PERSONAL SA", "RACIONAL DE INSUMOS S.A.S.", "ISAVELA ENTERPRISES, INC",
    "TIENDAS DE ROPA INTIMA S.A", "WIH. WISSNER GMBH & CO. KG", "C.I. CREACIONES SEGAR S.A.S",
]


def generar(ruta: str, hasta: dt.date = dt.date(2026, 11, 23), semilla: int = 7) -> str:
    rnd = random.Random(semilla)
    clientes = []
    for i in range(260):
        nombre = NOMBRES[i] if i < len(NOMBRES) else f"CLIENTE DEMO {i + 1:03d} S.A.S"
        nit = str(800000000 + rnd.randint(0, 199999999))
        peso = 1 / (i + 1) ** 1.15
        inicio = dt.date(2021, 1, 1) if rnd.random() < 0.7 else dt.date(rnd.randint(2021, 2026), rnd.randint(1, 12), 1)
        fin = hasta if rnd.random() < 0.85 else dt.date(rnd.randint(2022, 2026), rnd.randint(1, 12), 1)
        if i == 18:  # cliente nuevo en 2026
            inicio = dt.date(2026, 2, 1)
        clientes.append((nit, nombre, peso, inicio, fin))

    wb = Workbook()
    ws = wb.active
    ws.title = "Auxiliar"
    ws.append(["C.I. DEMO S.A.S"])
    ws.append(["Libro auxiliar por cuenta - Clase 4 Ingresos"])
    ws.append([f"Del 01/01/2021 al {hasta:%d/%m/%Y}"])
    ws.append([])
    ws.append(["Cuenta", "Nombre cuenta", "Fecha", "Comprobante", "NIT", "Nombre tercero",
               "Descripción", "Débito", "Crédito"])
    for c in ws[5]:
        c.font = Font(bold=True)

    consecutivo = 1000
    fecha = dt.date(2021, 1, 1)
    while fecha <= hasta:
        anio, mes = fecha.year, fecha.month
        estacional = 1 + 0.25 * (mes in (9, 10, 11)) - 0.3 * (mes == 6)
        tendencia = 1 - 0.035 * (anio - 2021)
        objetivo = 1_350_000_000 * estacional * tendencia * rnd.uniform(0.85, 1.15)
        activos = [c for c in clientes if c[3] <= fecha <= c[4]]
        total_peso = sum(c[2] for c in activos)
        ultimo_dia = (dt.date(anio + (mes == 12), mes % 12 + 1, 1) - dt.timedelta(days=1)).day
        for nit, nombre, peso, _, _ in activos:
            monto_cliente = objetivo * peso / total_peso * rnd.uniform(0.5, 1.5)
            facturas = max(1, int(peso * 40) + rnd.randint(0, 2))
            if rnd.random() < 0.25 and peso < 0.02:
                continue
            for _ in range(facturas):
                dia = rnd.randint(1, ultimo_dia)
                f = dt.date(anio, mes, dia)
                if f > hasta:
                    continue
                valor = round(monto_cliente / facturas * rnd.uniform(0.7, 1.3))
                consecutivo += 1
                ws.append(["41352001", "Venta de prendas de vestir", f, f"FV-{consecutivo}",
                           int(nit), nombre, "Factura de venta", 0, valor])
                if rnd.random() < 0.03:
                    ws.append(["41750101", "Devoluciones en ventas", f, f"NC-{consecutivo}",
                               int(nit), nombre, "Nota crédito devolución", round(valor * rnd.uniform(0.1, 0.5)), 0])
        ws.append(["4135", "Total comercio al por mayor", None, None, None, None, None, None, None])
        fecha = dt.date(anio + (mes == 12), mes % 12 + 1, 1)

    for fila in ws.iter_rows(min_row=6, min_col=3, max_col=3):
        fila[0].number_format = "DD/MM/YYYY"
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    wb.save(ruta)
    return ruta


if __name__ == "__main__":
    salida = sys.argv[1] if len(sys.argv) > 1 else os.path.join("demo", "chala aux 4.xlsx")
    print("Generado:", generar(salida))
