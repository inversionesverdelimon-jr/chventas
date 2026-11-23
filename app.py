"""Servidor local del tablero CH Ventas.

Uso:
    python app.py                 # usa config.json
    python app.py --demo          # usa datos de demostración
    python app.py --puerto 8080 --sin-navegador

Abre http://localhost:8050 en el navegador. Los datos se vuelven a leer
automáticamente cada vez que el archivo de Excel cambia.
"""

from __future__ import annotations

import argparse
import gzip
import json
import os
import sys
import threading
import traceback
import webbrowser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from chventas import __version__
from chventas.lector import Configuracion, cargar, firma_archivos, resolver_archivos

BASE = os.path.dirname(os.path.abspath(__file__))
ESTATICOS = os.path.join(BASE, "static")


def leer_config(ruta: str) -> dict:
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


class Fuente:
    """Mantiene los datos en memoria y los recarga cuando cambia el Excel."""

    def __init__(self, ruta_config: str, ruta_excel_forzada: str | None = None):
        self.ruta_config = ruta_config
        self.ruta_excel_forzada = ruta_excel_forzada
        self._lock = threading.Lock()
        self._firma = None
        self._paquete: bytes | None = None
        self._error: str | None = None

    def _config(self) -> Configuracion:
        d = leer_config(self.ruta_config) if os.path.isfile(self.ruta_config) else {}
        if self.ruta_excel_forzada:
            d["ruta_excel"] = self.ruta_excel_forzada
        if not d.get("ruta_excel"):
            raise ValueError("Falta 'ruta_excel' en config.json")
        return Configuracion.desde_dict(d)

    def firma_actual(self) -> str:
        try:
            cfg = self._config()
            archivos = resolver_archivos(cfg.ruta_excel)
            extra = os.stat(self.ruta_config).st_mtime_ns if os.path.isfile(self.ruta_config) else 0
            return f"{firma_archivos(archivos)}\ncfg|{extra}"
        except Exception as e:  # noqa: BLE001
            return f"error|{e}"

    def obtener(self) -> tuple[bytes | None, str | None, str]:
        with self._lock:
            firma = self.firma_actual()
            if firma != self._firma:
                try:
                    paquete = cargar(self._config())
                    paquete["firma"] = str(abs(hash(firma)))
                    self._paquete = json.dumps(paquete, ensure_ascii=False,
                                               separators=(",", ":")).encode("utf-8")
                    self._error = None
                except Exception as e:  # noqa: BLE001
                    traceback.print_exc()
                    self._paquete = None
                    self._error = str(e)
                self._firma = firma
            return self._paquete, self._error, str(abs(hash(firma)))


class Manejador(SimpleHTTPRequestHandler):
    fuente: Fuente = None  # se asigna al iniciar

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ESTATICOS, **kwargs)

    def log_message(self, fmt, *args):  # silenciar peticiones de archivos estáticos
        if "/api/" in (self.path or ""):
            sys.stderr.write("%s - %s\n" % (self.log_date_time_string(), fmt % args))

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _json(self, cuerpo: bytes, estado=HTTPStatus.OK):
        usar_gzip = "gzip" in (self.headers.get("Accept-Encoding") or "") and len(cuerpo) > 2048
        if usar_gzip:
            cuerpo = gzip.compress(cuerpo, compresslevel=5)
        self.send_response(estado)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        if usar_gzip:
            self.send_header("Content-Encoding", "gzip")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def do_GET(self):
        ruta = urlparse(self.path).path
        if ruta == "/api/datos":
            paquete, error, _ = self.fuente.obtener()
            if error:
                self._json(json.dumps({"error": error}, ensure_ascii=False).encode("utf-8"),
                           HTTPStatus.INTERNAL_SERVER_ERROR)
            else:
                self._json(paquete)
            return
        if ruta == "/api/estado":
            firma = str(abs(hash(self.fuente.firma_actual())))
            self._json(json.dumps({"firma": firma, "version": __version__}).encode("utf-8"))
            return
        super().do_GET()


def main(argv=None):
    p = argparse.ArgumentParser(description="Tablero de ventas CH Ventas")
    p.add_argument("--config", default=os.environ.get("CHVENTAS_CONFIG",
                                                      os.path.join(BASE, "config.json")))
    p.add_argument("--excel", help="Ruta del Excel o carpeta (reemplaza la de config.json)")
    p.add_argument("--demo", action="store_true", help="Usar datos de demostración")
    p.add_argument("--puerto", type=int, default=None)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--sin-navegador", action="store_true")
    args = p.parse_args(argv)

    cfg = leer_config(args.config) if os.path.isfile(args.config) else {}
    ruta_excel = args.excel
    if args.demo:
        ruta_excel = os.path.join(BASE, "demo", "chala aux 4.xlsx")
        if not os.path.isfile(ruta_excel):
            from tools.generar_demo import generar
            generar(ruta_excel)
    puerto = args.puerto or int(cfg.get("puerto", 8050))

    Manejador.fuente = Fuente(args.config, ruta_excel)
    servidor = ThreadingHTTPServer((args.host, puerto), Manejador)
    url = f"http://localhost:{puerto}/"
    print(f"CH Ventas {__version__} en {url}")
    print(f"Base de datos: {ruta_excel or cfg.get('ruta_excel')}")
    print("Presione Ctrl+C para detener.")

    # Primera lectura en segundo plano para que el navegador abra rápido.
    threading.Thread(target=Manejador.fuente.obtener, daemon=True).start()
    if not args.sin_navegador:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nDetenido.")
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
