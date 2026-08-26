"""
Descarga automatica de reportes de Convertia/inConcert, segun lo que este
definido en reportes_config.json (cuentas, sitios y reportes por cuenta).

Requisitos (una sola vez):
    pip install playwright python-dotenv
    playwright install chrome

Antes de correr este script, crea el archivo convertia.env junto a este
script (C:\\Users\\dy2059\\Downloads\\convertia.env) con estas dos lineas:

    CONVERTIA_USER=tu_usuario
    CONVERTIA_PASS=tu_contrasena

El script abre su PROPIA ventana de Chrome (perfil temporal, independiente
de tu Chrome normal) e inicia sesion desde cero con esas credenciales -- no
hace falta cerrar tu Chrome ni tus pestañas antes de correrlo.

El reporte puede tardar entre 20 segundos y 5 minutos en generarse; el
script espera hasta ese margen antes de fallar por timeout.

Uso:
    python descargar_reportes_convertia.py
        Corre con el rango de fechas por defecto (2 meses atras + domingo
        anterior a hoy) para todas las cuentas del config.

    python descargar_reportes_convertia.py --fecha-inicio 2026-06-01 --fecha-fin 2026-08-16
        Fuerza un rango de fechas para todas las cuentas.

    python descargar_reportes_convertia.py --fecha-inicio 2026-06-01 --fecha-fin 2026-08-16 \\
        --fechas-cuenta "{\\"Claro_Cr\\": [\\"2026-07-01\\", \\"2026-08-16\\"]}"
        Igual que arriba, pero Claro_Cr usa su propio rango distinto.

    python descargar_reportes_convertia.py --modo-prueba
        Rango del mes presente y solo la primera cuenta de cada sitio, para
        iterar rapido mientras se prueban cambios al script.

Si algun reporte ya se descargo hoy (existe el .xlsx de esa cuenta/alias con
la fecha de hoy), se omite automaticamente -- asi, si una corrida fallo a
mitad de camino, correr el mismo comando de nuevo solo baja lo que falto.
Usa --forzar para descargar todo de nuevo aunque ya exista.
"""

import argparse
import json
import os
import re
import time
from datetime import date, timedelta
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

SCRIPT_DIR = Path(__file__).resolve().parent  # Codigo/ -- config y credenciales
RAIZ = SCRIPT_DIR.parent  # carpeta del agente -- donde viven los datos (crudos, Procesados)
load_dotenv(SCRIPT_DIR / "convertia.env")

CONVERTIA_USER = os.environ.get("CONVERTIA_USER")
CONVERTIA_PASS = os.environ.get("CONVERTIA_PASS")
if not CONVERTIA_USER or not CONVERTIA_PASS:
    raise SystemExit(
        f"Faltan credenciales: crea {SCRIPT_DIR / 'convertia.env'} con las lineas "
        "CONVERTIA_USER=... y CONVERTIA_PASS=..."
    )

CONFIG_PATH = SCRIPT_DIR / "reportes_config.json"
DOWNLOAD_DIR = RAIZ
HEADLESS = False  # pon False si quieres ver el navegador mientras corre

REPORT_LOAD_TIMEOUT_MS = 300_000  # 5 minutos (los reportes "brutas" con el rango completo de 2 meses a veces tardan mas de 2.5 min)
LOGIN_TIMEOUT_MS = 20_000
MAX_REINTENTOS_EJECUTAR = 15
ESPERA_REINTENTO_MS = 15_000


def cargar_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def ruta_destino(alias: str, cuenta: str, fecha_inicio: date, fecha_fin: date) -> Path:
    """El nombre incluye el RANGO consultado (no la fecha de descarga): dos
    descargas el mismo dia con rangos distintos deben dar archivos distintos.
    Antes solo llevaba la fecha de hoy, y una segunda descarga del mismo dia
    con otro rango pisaba silenciosamente al primero (bug real, 25/08/2026)."""
    return DOWNLOAD_DIR / f"{alias.upper()}_{cuenta}_{fecha_inicio.isoformat()}_{fecha_fin.isoformat()}.xlsx"


def fecha_inicio_dos_meses_atras(hoy: date) -> date:
    """Primer dia del mes, retrocediendo 2 meses (p.ej. 18-ago -> 01-jun)."""
    mes = hoy.month - 2
    anio = hoy.year
    while mes <= 0:
        mes += 12
        anio -= 1
    return date(anio, mes, 1)


def domingo_anterior(hoy: date) -> date:
    """Domingo inmediatamente anterior a hoy (nunca hoy mismo)."""
    dias_desde_domingo = (hoy.weekday() + 1) % 7  # lunes=0 ... domingo=6
    if dias_desde_domingo == 0:
        dias_desde_domingo = 7
    return hoy - timedelta(days=dias_desde_domingo)


def esperar_carga_o_aviso_ocupado(page: Page, aviso_ocupado, boton_carga, timeout_ms: int) -> str:
    """Sondea manualmente (en vez de usar .or_()) porque cuando el aviso de
    "otro reporte ejecutandose" y el boton de columnas (oculto en el DOM
    antes de cargar) coinciden, Playwright tira un error de "multiples
    elementos" al combinarlos con .or_(). Revisando cada uno por separado
    con is_visible() se evita ese conflicto."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if boton_carga.is_visible():
            return "cargado"
        if aviso_ocupado.is_visible():
            return "ocupado"
        page.wait_for_timeout(500)
    raise TimeoutError("El reporte no cargo ni aparecio aviso de 'otro reporte ejecutandose'.")


def seleccionar_todas_las_columnas(page: Page, seleccionar_columnas_btn) -> None:
    """A veces el panel de filtros (cuenta/periodo/fechas) queda superpuesto
    sobre el menu de columnas y bloquea el clic en "Seleccionar todo" (nos
    paso con Mundo_Pacifico: "Timeout 30000ms exceeded ... intercepts pointer
    events"). Se reintenta cerrando cualquier overlay con Escape antes de
    cada intento."""
    seleccionar_todo = page.get_by_text("Seleccionar todo", exact=True)
    intentos = 3
    for intento in range(1, intentos + 1):
        try:
            seleccionar_columnas_btn.click()
            seleccionar_todo.click(timeout=10_000)
            page.keyboard.press("Escape")
            return
        except PlaywrightTimeoutError:
            if intento == intentos:
                raise
            print(
                f"No se pudo abrir/seleccionar columnas (intento {intento}/{intentos}), reintentando...",
                flush=True,
            )
            page.keyboard.press("Escape")
            page.wait_for_timeout(1_000)


def guardar_captura_diagnostico(page: Page, nombre_archivo: str) -> None:
    debug_path = DOWNLOAD_DIR / nombre_archivo
    try:
        page.screenshot(path=str(debug_path))
        print(f"Captura de diagnostico guardada: {debug_path}", flush=True)
    except Exception as exc:
        print(f"No se pudo guardar la captura de diagnostico: {exc}", flush=True)


def iniciar_sesion(page: Page, base_url: str) -> None:
    print(f"Iniciando sesion en {base_url} ...", flush=True)
    page.goto(f"{base_url}/login", wait_until="networkidle")
    page.get_by_placeholder("Usuario o correo electrónico").fill(CONVERTIA_USER)
    page.get_by_placeholder("Ingresa tu contraseña").fill(CONVERTIA_PASS)
    # El texto del boton varia segun el sitio: "Ingresar." en uno,
    # "Ingresar" (sin punto) en otro. exact=False cubre ambos casos.
    page.get_by_role("button", name="Ingresar", exact=False).click()
    page.wait_for_url(lambda url: "/login" not in url, timeout=LOGIN_TIMEOUT_MS)
    print("Sesion iniciada correctamente.", flush=True)


def descargar_reporte(
    page: Page,
    base_url: str,
    nombre_reporte: str,
    alias: str,
    cuenta: str,
    fecha_inicio: date,
    fecha_fin: date,
) -> None:
    print(f"--- Descargando '{alias}' para {cuenta} ({nombre_reporte}) ---", flush=True)
    report_url = f"{base_url}/mas/analytics/report/view/{nombre_reporte}"
    page.goto(report_url, wait_until="networkidle")

    # El texto de los placeholders varia segun el sitio (p.ej. "Seleccione
    # la cuenta..." vs "Seleccionar cuenta"), asi que se busca por
    # coincidencia parcial (regex) de la palabra clave, no el texto exacto.
    cuenta_input = page.get_by_placeholder(re.compile("cuenta", re.IGNORECASE))
    periodo_input = page.get_by_placeholder(re.compile("periodo", re.IGNORECASE))

    # Cuenta (se usa role=link porque algunos nombres de cuenta son substring
    # de otro, p.ej. "wowperu" / "wowperu_in", y get_by_text ambiguo falla)
    cuenta_input.click()
    page.get_by_role("link", name=cuenta, exact=True).click()

    # Periodo: forzar "Personalizado" explicitamente. En algunos reportes/
    # sitios el valor por defecto es otro (p.ej. "Hoy" en el sitio de Claro),
    # y mientras no diga "Personalizado" los campos de fecha quedan
    # bloqueados sin poder editarse.
    periodo_input.click()
    page.get_by_text("Personalizado", exact=True).click()

    # Fechas (yyyy-mm-dd)
    fechas = page.get_by_placeholder("yyyy-mm-dd")
    fechas.nth(0).fill(fecha_inicio.isoformat())
    page.keyboard.press("Escape")
    fechas.nth(1).fill(fecha_fin.isoformat())
    page.keyboard.press("Escape")

    # Estos botones no tienen texto visible ni aria-label: su nombre
    # accesible sale del atributo title, y get_by_role(name=...) no los
    # encontraba de forma confiable. Se usa selector CSS por title en su
    # lugar, que es directo y no depende del calculo de nombre accesible.
    seleccionar_columnas_btn = page.locator('button[title="Seleccionar columnas"]')

    # Ejecutar, con reintentos: si otro usuario esta corriendo un reporte al
    # mismo tiempo, aparece un aviso "Existe otro reporte ejecutandose..." y
    # el boton Ejecutar vuelve a su estado normal sin cargar nada -- hay que
    # esperar unos segundos y volver a darle Ejecutar.
    ejecutar_btn = page.get_by_role("button", name="Ejecutar", exact=True)
    aviso_ocupado = page.get_by_text("Existe otro reporte ejecutándose", exact=False)

    for intento in range(1, MAX_REINTENTOS_EJECUTAR + 1):
        ejecutar_btn.click()
        resultado = esperar_carga_o_aviso_ocupado(
            page, aviso_ocupado, seleccionar_columnas_btn, REPORT_LOAD_TIMEOUT_MS
        )

        if resultado == "cargado":
            break

        print(
            f"Otro reporte esta corriendo, reintentando en unos segundos "
            f"(intento {intento}/{MAX_REINTENTOS_EJECUTAR})...",
            flush=True,
        )
        page.wait_for_timeout(ESPERA_REINTENTO_MS)
    else:
        raise RuntimeError(
            "Se agotaron los reintentos: el reporte de otro usuario nunca libero el turno."
        )

    # Seleccionar todas las columnas -- UN SOLO CLIC.
    # Ojo: si se hace doble clic aqui, se desmarcan todas las columnas y el
    # Excel sale vacio/incompleto (nos paso durante las pruebas).
    seleccionar_todas_las_columnas(page, seleccionar_columnas_btn)

    exportar_btn = page.locator('button[title="Exportar datos"]')

    with page.expect_download() as download_info:
        exportar_btn.click()
        page.get_by_text("Exportar Excel", exact=True).click()
    download = download_info.value

    destino = ruta_destino(alias, cuenta, fecha_inicio, fecha_fin)
    download.save_as(destino)
    print(f"Guardado: {destino}", flush=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fecha-inicio", type=str, default=None, help="yyyy-mm-dd, aplica a todas las cuentas salvo que --fechas-cuenta la sobreescriba")
    parser.add_argument("--fecha-fin", type=str, default=None, help="yyyy-mm-dd, aplica a todas las cuentas salvo que --fechas-cuenta la sobreescriba")
    parser.add_argument(
        "--fechas-cuenta",
        type=str,
        default=None,
        help='JSON con overrides por cuenta, ej: {"Claro_Cr": ["2026-07-01", "2026-08-16"]}',
    )
    parser.add_argument(
        "--modo-prueba",
        action="store_true",
        help="Rango del mes presente y solo la primera cuenta de cada sitio, para iterar rapido",
    )
    parser.add_argument(
        "--forzar",
        action="store_true",
        help=(
            "Volver a descargar aunque ya exista el archivo de hoy para esa cuenta/reporte. "
            "Sin esta opcion, si ya existe se omite (asi una corrida que fallo a mitad de "
            "camino se puede repetir y solo baja lo que falto)."
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = cargar_config()
    sitios = config["sitios"]
    cuentas = config["cuentas"]

    hoy = date.today()
    if args.modo_prueba:
        fecha_inicio_default = hoy.replace(day=1)
        fecha_fin_default = hoy
        # solo la primera cuenta de cada sitio, para iterar rapido
        vistos = set()
        cuentas_a_correr = {}
        for cuenta, datos in cuentas.items():
            if datos["sitio"] not in vistos:
                cuentas_a_correr[cuenta] = datos
                vistos.add(datos["sitio"])
    else:
        fecha_inicio_default = (
            date.fromisoformat(args.fecha_inicio) if args.fecha_inicio else fecha_inicio_dos_meses_atras(hoy)
        )
        fecha_fin_default = date.fromisoformat(args.fecha_fin) if args.fecha_fin else domingo_anterior(hoy)
        cuentas_a_correr = cuentas

    # Overrides de fecha por cuenta (p.ej. Claro_Cr con un rango distinto)
    fechas_por_cuenta: dict[str, tuple[date, date]] = {}
    if args.fechas_cuenta:
        overrides = json.loads(args.fechas_cuenta)
        for cta, (ini, fin) in overrides.items():
            fechas_por_cuenta[cta] = (date.fromisoformat(ini), date.fromisoformat(fin))

    print(
        f"Rango de fechas por defecto: {fecha_inicio_default} a {fecha_fin_default} | "
        f"cuentas: {list(cuentas_a_correr.keys())} | "
        f"excepciones de fecha: {fechas_por_cuenta}",
        flush=True,
    )

    # Agrupar por sitio para no repetir el login de mas
    por_sitio: dict[str, list[tuple[str, list[dict]]]] = {}
    for cuenta, datos in cuentas_a_correr.items():
        por_sitio.setdefault(datos["sitio"], []).append((cuenta, datos["reportes"]))

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=HEADLESS)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1366, "height": 800})
        page = context.new_page()

        for sitio_id, cuentas_del_sitio in por_sitio.items():
            base_url = sitios[sitio_id]
            try:
                iniciar_sesion(page, base_url)
            except Exception as exc:
                print(f"ERROR iniciando sesion en {base_url}: {exc}", flush=True)
                guardar_captura_diagnostico(page, f"_debug_error_login_{sitio_id}.png")
                print(f"Se omiten todas las cuentas del sitio '{sitio_id}' en esta corrida.", flush=True)
                continue

            for cuenta, reportes in cuentas_del_sitio:
                fecha_inicio, fecha_fin = fechas_por_cuenta.get(cuenta, (fecha_inicio_default, fecha_fin_default))
                for rep in reportes:
                    destino = ruta_destino(rep["alias"], cuenta, fecha_inicio, fecha_fin)
                    if destino.exists() and not args.forzar:
                        print(
                            f"Ya existe {destino.name}, se omite (usa --forzar para volver a descargarlo).",
                            flush=True,
                        )
                        continue
                    try:
                        descargar_reporte(
                            page,
                            base_url,
                            rep["nombre_reporte"],
                            rep["alias"],
                            cuenta,
                            fecha_inicio,
                            fecha_fin,
                        )
                    except Exception as exc:
                        print(f"ERROR descargando {rep['alias']} de {cuenta}: {exc}", flush=True)
                        guardar_captura_diagnostico(page, f"_debug_error_{cuenta}_{rep['alias']}.png")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
