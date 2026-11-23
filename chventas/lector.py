"""Lectura y normalización de la base de ventas en Excel.

Soporta dos formatos de archivo:

* **Contable** (auxiliar de cuentas): columnas tipo Cuenta, Fecha, NIT,
  Tercero, Débito, Crédito. Ventas = créditos - débitos de las cuentas de
  ingreso (por defecto prefijo ``41``); devoluciones = débitos - créditos de
  las cuentas de devoluciones (por defecto prefijo ``4175``).
* **Valor**: una columna de valor por fila (venta), con devoluciones en una
  columna aparte, identificadas por una columna "tipo" o como valores
  negativos.

La ruta configurada puede ser un archivo (con o sin extensión) o una carpeta;
en el caso de una carpeta se leen y combinan todos los archivos de Excel/CSV
que contenga.
"""

from __future__ import annotations

import datetime as dt
import math
import os
import re
import shutil
import tempfile
import time
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

EXTENSIONES = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")

# Encabezados reconocidos (normalizados: minúsculas, sin tildes ni signos).
ALIAS = {
    "fecha": [
        "fecha", "fecha documento", "fecha doc", "fecha factura", "fecha elaboracion",
        "fecha comprobante", "fecha movimiento", "fecha contable", "fec", "date",
    ],
    "cuenta": [
        "cuenta", "codigo cuenta", "cuenta contable", "cod cuenta", "codigo contable",
        "codigo", "cta", "auxiliar",
    ],
    "cliente_id": [
        "nit", "identificacion", "nit tercero", "cedula", "documento tercero",
        "id tercero", "codigo tercero", "nit cliente", "codigo cliente", "id cliente",
        "nit cedula", "nro identificacion", "numero identificacion", "doc tercero",
    ],
    "cliente_nombre": [
        "nombre tercero", "tercero", "cliente", "razon social", "nombre cliente",
        "nombre del tercero", "nombre", "nombre razon social", "nombre o razon social",
    ],
    "debito": ["debito", "debitos", "debe", "valor debito", "db"],
    "credito": ["credito", "creditos", "haber", "valor credito", "cr"],
    "valor": [
        "valor", "total", "valor neto", "venta neta", "ventas netas", "subtotal",
        "valor total", "total venta", "venta", "ventas", "importe", "monto",
        "valor venta", "neto",
    ],
    "tipo": [
        "tipo", "tipo documento", "tipo doc", "tipo comprobante", "clase",
        "tipo de documento", "tipo movimiento", "concepto",
    ],
    "devolucion": ["devolucion", "devoluciones", "valor devolucion", "devol"],
}

PALABRAS_DEVOLUCION = ["devol", "nota credito", "nc ", "notas credito", "anulacion"]


def normalizar(texto) -> str:
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return s.strip()


# ---------------------------------------------------------------------------
# Conversión de valores
# ---------------------------------------------------------------------------

_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def a_fecha(v):
    """Convierte un valor de celda a ``datetime.date`` o ``None``."""
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.date()
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return None
        if 20000 < v < 80000:  # número de serie de Excel
            return (_EXCEL_EPOCH + pd.Timedelta(days=float(v))).date()
        return None
    if isinstance(v, str):
        s = v.strip()
        if not s or not re.search(r"\d", s):
            return None
        m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
        try:
            if m:
                return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", s)
            if m:
                a = int(m.group(3))
                a = a + 2000 if a < 100 else a
                return dt.date(a, int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
        try:
            ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        except Exception:  # noqa: BLE001
            return None
        return None if pd.isna(ts) else ts.date()
    return None


def a_numero(v) -> float:
    """Convierte un valor de celda a número (acepta formato 1.234.567,89)."""
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return 0.0 if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo, s = True, s[1:-1]
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.startswith("-"):
        negativo, s = (not negativo), s[1:]
    s = s.replace("-", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif "." in s:
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if negativo else n


def limpiar_id(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


# ---------------------------------------------------------------------------
# Localización de archivos
# ---------------------------------------------------------------------------

def resolver_archivos(ruta: str) -> list[str]:
    """Devuelve la lista de archivos a leer a partir de la ruta configurada."""
    ruta = os.path.expandvars(os.path.expanduser(str(ruta).strip().strip('"')))
    if os.path.isdir(ruta):
        archivos = [
            os.path.join(ruta, f)
            for f in sorted(os.listdir(ruta))
            if f.lower().endswith(EXTENSIONES) and not f.startswith("~$")
            and os.path.isfile(os.path.join(ruta, f))
        ]
        if not archivos:
            raise FileNotFoundError(
                f"La carpeta existe pero no contiene archivos de Excel: {ruta}"
            )
        return archivos
    if os.path.isfile(ruta):
        return [ruta]
    for ext in EXTENSIONES:
        if os.path.isfile(ruta + ext):
            return [ruta + ext]
    # Coincidencia aproximada: archivo que empiece por el nombre indicado.
    carpeta, base = os.path.split(ruta)
    if carpeta and os.path.isdir(carpeta):
        base_n = normalizar(base)
        candidatos = [
            os.path.join(carpeta, f)
            for f in sorted(os.listdir(carpeta))
            if f.lower().endswith(EXTENSIONES) and not f.startswith("~$")
            and normalizar(os.path.splitext(f)[0]).startswith(base_n)
        ]
        if candidatos:
            return candidatos[:1]
    raise FileNotFoundError(
        "No se encontró la base de datos en la ruta configurada: "
        f"{ruta}. Revise 'ruta_excel' en config.json (puede ser un archivo "
        ".xlsx o una carpeta con archivos de Excel)."
    )


def firma_archivos(archivos: list[str]) -> str:
    partes = []
    for a in archivos:
        try:
            st = os.stat(a)
            partes.append(f"{a}|{st.st_mtime_ns}|{st.st_size}")
        except OSError:
            partes.append(f"{a}|?")
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------

def _leer_crudo(ruta: str, hoja) -> dict[str, pd.DataFrame]:
    if ruta.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(ruta, sep=None, engine="python", header=None,
                                 dtype=object, encoding=enc)
                return {"csv": df}
            except UnicodeDecodeError:
                continue
        raise ValueError(f"No se pudo leer el CSV {ruta}")
    hojas = pd.read_excel(ruta, sheet_name=hoja if hoja not in ("", None) else None,
                          header=None, dtype=object)
    if isinstance(hojas, pd.DataFrame):
        hojas = {str(hoja): hojas}
    return hojas


def leer_archivo(ruta: str, hoja=None) -> dict[str, pd.DataFrame]:
    """Lee un archivo; si está bloqueado (Excel abierto / OneDrive) lee una copia."""
    try:
        return _leer_crudo(ruta, hoja)
    except PermissionError:
        tmp_dir = tempfile.mkdtemp(prefix="chventas_")
        copia = os.path.join(tmp_dir, os.path.basename(ruta))
        for intento in range(3):
            try:
                shutil.copy2(ruta, copia)
                break
            except PermissionError:
                if intento == 2:
                    raise
                time.sleep(1)
        try:
            return _leer_crudo(copia, hoja)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _puntaje_fila(valores) -> int:
    todos = {a for lista in ALIAS.values() for a in lista}
    return sum(1 for v in valores if normalizar(v) in todos)


def detectar_encabezado(df: pd.DataFrame, max_filas: int = 40) -> int | None:
    mejor, mejor_puntaje = None, 1
    for i in range(min(max_filas, len(df))):
        p = _puntaje_fila(df.iloc[i].tolist())
        if p > mejor_puntaje:
            mejor, mejor_puntaje = i, p
    return mejor


def mapear_columnas(encabezados: list, forzadas: dict | None = None) -> dict[str, int]:
    """Asigna cada campo lógico a un índice de columna."""
    forzadas = {k: v for k, v in (forzadas or {}).items() if v not in (None, "")}
    norm = [normalizar(h) for h in encabezados]
    mapa: dict[str, int] = {}
    usadas: set[int] = set()

    for campo, nombre in forzadas.items():
        n = normalizar(nombre)
        if n in norm:
            mapa[campo] = norm.index(n)
            usadas.add(mapa[campo])

    # 1) coincidencia exacta, en orden de prioridad de alias
    for campo, alias in ALIAS.items():
        if campo in mapa:
            continue
        for a in alias:
            idx = next((i for i, h in enumerate(norm) if h == a and i not in usadas), None)
            if idx is not None:
                mapa[campo] = idx
                usadas.add(idx)
                break
    # 2) coincidencia parcial (el encabezado empieza con el alias)
    for campo, alias in ALIAS.items():
        if campo in mapa:
            continue
        for a in alias:
            if len(a) < 4:
                continue
            idx = next((i for i, h in enumerate(norm)
                        if h.startswith(a + " ") and i not in usadas), None)
            if idx is not None:
                mapa[campo] = idx
                usadas.add(idx)
                break
    return mapa


@dataclass
class Configuracion:
    ruta_excel: str
    hoja: str | None = None
    columnas: dict = field(default_factory=dict)
    prefijos_ventas: list = field(default_factory=lambda: ["41"])
    prefijos_devoluciones: list = field(default_factory=lambda: ["4175"])
    palabras_devolucion: list = field(default_factory=lambda: list(PALABRAS_DEVOLUCION))

    @classmethod
    def desde_dict(cls, d: dict) -> "Configuracion":
        campos = {k: d[k] for k in cls.__dataclass_fields__ if k in d and d[k] is not None}
        return cls(**campos)


def _normalizar_hoja(df: pd.DataFrame, cfg: Configuracion, nombre: str):
    """Devuelve (DataFrame normalizado, info) o (None, motivo)."""
    fila = detectar_encabezado(df)
    if fila is None:
        return None, f"{nombre}: no se encontró fila de encabezados"
    encabezados = df.iloc[fila].tolist()
    mapa = mapear_columnas(encabezados, cfg.columnas)
    datos = df.iloc[fila + 1:].reset_index(drop=True)

    if "fecha" not in mapa:
        return None, f"{nombre}: sin columna de fecha"
    contable = "debito" in mapa and "credito" in mapa
    if not contable and "valor" not in mapa:
        return None, f"{nombre}: sin columnas de valor (Valor o Débito/Crédito)"
    if "cliente_id" not in mapa and "cliente_nombre" not in mapa:
        return None, f"{nombre}: sin columna de cliente/tercero"

    col = lambda campo: datos.iloc[:, mapa[campo]] if campo in mapa else None  # noqa: E731

    fechas = col("fecha").map(a_fecha)
    validas = fechas.notna()
    datos = datos[validas]
    fechas = fechas[validas]

    ids = col("cliente_id")
    nombres = col("cliente_nombre")
    ids = ids[validas].map(limpiar_id) if ids is not None else None
    nombres = (nombres[validas].map(lambda v: "" if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip())
               if nombres is not None else None)
    if ids is None:
        # "811044814 - NOMBRE" en una sola columna
        partes = nombres.str.extract(r"^\s*([0-9A-Za-z.\-]+?)\s+-\s+(.+)$")
        ids = partes[0].fillna("").str.replace(".", "", regex=False)
        nombres = partes[1].fillna(nombres)
        ids = ids.where(ids != "", nombres)
    if nombres is None:
        nombres = ids

    if contable:
        deb = datos.iloc[:, mapa["debito"]].map(a_numero)
        cre = datos.iloc[:, mapa["credito"]].map(a_numero)
        if "cuenta" in mapa:
            cuentas = datos.iloc[:, mapa["cuenta"]].map(limpiar_id).str.replace(r"\D", "", regex=True)
            es_dev = cuentas.map(lambda c: any(c.startswith(p) for p in cfg.prefijos_devoluciones))
            es_vta = cuentas.map(lambda c: any(c.startswith(p) for p in cfg.prefijos_ventas)) & ~es_dev
            if cfg.prefijos_ventas == [] or not cfg.prefijos_ventas:
                es_vta = ~es_dev
        else:
            es_dev = pd.Series(False, index=datos.index)
            es_vta = pd.Series(True, index=datos.index)
        venta = (cre - deb).where(es_vta, 0.0)
        devol = (deb - cre).where(es_dev, 0.0)
        incluir = es_vta | es_dev
        modo = "contable"
    else:
        valor = datos.iloc[:, mapa["valor"]].map(a_numero)
        if "devolucion" in mapa:
            devol = datos.iloc[:, mapa["devolucion"]].map(a_numero).abs()
            venta = valor
        elif "tipo" in mapa:
            palabras = [normalizar(p) for p in cfg.palabras_devolucion]
            tipos = datos.iloc[:, mapa["tipo"]].map(lambda t: " " + normalizar(t) + " ")
            es_dev = tipos.map(lambda t: any(p and p in t for p in palabras))
            devol = valor.abs().where(es_dev, 0.0)
            venta = valor.where(~es_dev, 0.0)
        else:
            devol = (-valor).where(valor < 0, 0.0)
            venta = valor.where(valor >= 0, 0.0)
        incluir = pd.Series(True, index=datos.index)
        modo = "valor"

    salida = pd.DataFrame({
        "fecha": fechas.values,
        "cliente_id": ids.values,
        "cliente_nombre": nombres.values,
        "venta": venta.values,
        "devol": devol.values,
    })[incluir.values]
    salida = salida[(salida["venta"] != 0) | (salida["devol"] != 0)]
    info = {
        "hoja": nombre,
        "fila_encabezado": fila + 1,
        "modo": modo,
        "columnas": {k: str(encabezados[v]) for k, v in mapa.items()},
        "filas": int(len(salida)),
    }
    return salida, info


def cargar(cfg: Configuracion) -> dict:
    """Lee la(s) fuente(s) y devuelve el paquete agregado para el navegador."""
    t0 = time.perf_counter()
    archivos = resolver_archivos(cfg.ruta_excel)
    partes, detalles, avisos = [], [], []
    for archivo in archivos:
        try:
            hojas = leer_archivo(archivo, cfg.hoja)
        except Exception as e:  # noqa: BLE001
            avisos.append(f"{os.path.basename(archivo)}: {e}")
            continue
        for nombre, df in hojas.items():
            etiqueta = f"{os.path.basename(archivo)} › {nombre}"
            norm, info = _normalizar_hoja(df, cfg, etiqueta)
            if norm is None:
                avisos.append(info)
                continue
            partes.append(norm)
            detalles.append(info)

    if not partes:
        raise ValueError(
            "No se encontraron datos de ventas reconocibles. "
            + (" | ".join(avisos) if avisos else "")
            + " — Ajuste la sección 'columnas' de config.json con los nombres "
            "exactos de los encabezados."
        )

    datos = pd.concat(partes, ignore_index=True)
    datos["clave"] = datos["cliente_id"].where(datos["cliente_id"] != "", datos["cliente_nombre"])
    datos = datos[datos["clave"] != ""]
    datos["dia"] = datos["fecha"].map(lambda d: d.year * 10000 + d.month * 100 + d.day)

    # Nombre más reciente por cliente
    ultimos = datos.sort_values("dia").groupby("clave")["cliente_nombre"].last()
    claves = sorted(ultimos.index)
    indice = {c: i for i, c in enumerate(claves)}
    clientes = []
    for c in claves:
        nombre = str(ultimos[c] or "").strip()
        etiqueta = c if not nombre or nombre == c else f"{c} - {nombre}"
        clientes.append(etiqueta)

    agg = datos.groupby(["dia", "clave"], sort=True)[["venta", "devol"]].sum().reset_index()
    registros = [
        [int(r.dia), indice[r.clave], round(float(r.venta), 2), round(float(r.devol), 2)]
        for r in agg.itertuples(index=False)
    ]
    return {
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "ms_lectura": round((time.perf_counter() - t0) * 1000),
        "fuente": {
            "ruta": cfg.ruta_excel,
            "archivos": archivos,
            "hojas": detalles,
            "avisos": avisos,
            "filas": int(len(datos)),
        },
        "fecha_min": int(datos["dia"].min()),
        "fecha_max": int(datos["dia"].max()),
        "clientes": clientes,
        "registros": registros,
    }
