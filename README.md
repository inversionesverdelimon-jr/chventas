# chventas

Tablero de **ventas netas** (ventas − devoluciones) que se conecta automáticamente
a la base de datos en Excel y la vuelve a leer cada vez que el archivo cambia
(por ejemplo, cuando OneDrive sincroniza una nueva versión).

## Qué muestra

- **Periodo actual / Mes anterior / Segundo mes anterior**, cada uno comparado con el
  mismo periodo del año anterior (el mes en curso se compara con el mismo rango de días).
- Selector de **1 a 4 años** a comparar y **hasta qué mes**.
- Indicadores: ventas netas de cada año, variación y devoluciones.
- **Gráfica de ventas netas por mes** (el mes en curso se ve en color más claro).
- **Detalle mensual** con variación, variación % y variación % acumulada.
- **Top 20 clientes** en cuatro periodos: acumulado del año, periodo actual, mes anterior
  y segundo mes anterior. Incluye concentración, clientes con compras, clientes nuevos y
  clientes sin compras.
- Botones **Copiar CSV**: copian la tabla separada por `;` para pegarla en Excel.

## Instalación en Windows

1. Instale Python 3.10 o superior desde <https://www.python.org/downloads/>. En el
   instalador, marque **"Add Python to PATH"**.
2. Descargue este repositorio (botón *Code → Download ZIP*) y descomprímalo, o clónelo
   con Git.
3. Haga doble clic en **`iniciar.bat`**. La primera vez crea el entorno e instala las
   dependencias (toma un par de minutos). Después abre el navegador en
   <http://localhost:8050>.

Deje abierta la ventana negra mientras use el tablero; para detenerlo, ciérrela o
presione `Ctrl+C`.

## Conexión a la base de datos

La ruta está en `config.json`:

```json
"ruta_excel": "C:\\Users\\contabilidad\\OneDrive - CI CHALA SAS\\Office Lens\\2026_Automatización app\\chala aux 4"
```

- Puede ser un **archivo** (con o sin extensión: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`)
  o una **carpeta**; en ese caso se leen y combinan todos los Excel que contenga.
- Escriba las barras invertidas dobles (`\\`), como en el ejemplo.
- Si el archivo está abierto en Excel o OneDrive lo está sincronizando, la app lee una copia.
- La página revisa cada 30 segundos si el archivo cambió y se actualiza sola. El botón
  **Generar** fuerza una lectura inmediata.

### Formatos de Excel reconocidos

Las columnas se detectan solas por el nombre del encabezado, aunque el reporte tenga
filas de título arriba.

| Formato | Columnas | Cálculo |
|---|---|---|
| **Auxiliar contable** | Cuenta, Fecha, NIT, Nombre tercero, Débito, Crédito | Ventas = crédito − débito en cuentas `41…`; devoluciones = débito − crédito en `4175…` |
| **Lista de ventas** | Fecha, Cliente/NIT, Valor (y opcionalmente Tipo o Devolución) | Devoluciones según la columna Tipo (p. ej. "Devolución", "Nota crédito"), la columna Devolución o los valores negativos |

Si una columna no se reconoce, escriba el nombre exacto del encabezado en la sección
`columnas` de `config.json`, por ejemplo `"cliente_nombre": "Razón social"`. Los
prefijos de cuentas se ajustan en `prefijos_ventas` y `prefijos_devoluciones`.

Al pie de la página se ven el archivo leído, el número de movimientos y las columnas
detectadas. Sirve para comprobar que la lectura sea correcta.

## Otras opciones

```bash
python app.py --demo                 # datos de demostración
python app.py --excel "D:\otra\ruta.xlsx"
python app.py --puerto 8080 --sin-navegador
```

## Desarrollo

```bash
pip install -r requirements.txt pytest
python -m pytest -q
```

Estructura:

- `app.py`: servidor local (biblioteca estándar de Python) y API `/api/datos` y `/api/estado`.
- `chventas/lector.py`: lectura del Excel, detección de columnas y normalización.
- `static/`: interfaz web (HTML, CSS, JS y Chart.js incluido, funciona sin internet).
- `tools/generar_demo.py`: genera un Excel de ejemplo con formato de auxiliar contable.
