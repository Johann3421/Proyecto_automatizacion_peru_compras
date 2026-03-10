# -*- coding: utf-8 -*-
"""
peru_compras_bot.py
Automatiza la actualización de stock en catalogos.perucompras.gob.pe

Requisitos:
    pip install selenium pandas openpyxl

Uso:
    python peru_compras_bot.py
"""

import csv
import logging
import os
import re
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    NoSuchElementException,
    TimeoutException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------------------------------------------------------
# CONFIGURACION - Ajusta estas variables segun tu entorno
# ---------------------------------------------------------------------------
EXCEL_PATH = Path(__file__).parent / "productos.xlsx"  # Ruta al archivo Excel
BASE_URL = "https://www.catalogos.perucompras.gob.pe"
LOGIN_URL = f"{BASE_URL}/AccesoGeneral"
MEJORA_URL = f"{BASE_URL}/MejoraBasica"

# Textos visibles de los selectores de filtro
ACUERDO_TEXTO = "EXT-CE-2022-5 COMPUTADORAS DE ESCRITORIO"
CATALOGO_TEXTO = "COMPUTADORAS DE ESCRITORIO"
CATEGORIA_TEXTO = "MONITOR"

# Tiempos de espera (segundos)
WAIT_NORMAL = 15
WAIT_LARGO = 30
WAIT_CORTO = 5
PAUSA_ENTRE_PRODUCTOS = 2  # Pausa entre iteraciones para no sobrecargar

# Estado para modo GUI
MODO_GUI = False
EVENTO_LOGIN = None
GUI_NOTIFICAR_LOGIN = None

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            Path(__file__).parent / "bot.log", encoding="utf-8"
        ),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# REPORTE EXCEL
# ---------------------------------------------------------------------------
REPORTE_PATH = Path(__file__).parent / f"reporte_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
RESULTADOS: list = []


def nueva_ruta_reporte():
    return Path(__file__).parent / f"reporte_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


def clasificar_error(mensaje: str) -> str:
    """Convierte un error técnico en una categoría comprensible para el usuario."""
    msg = str(mensaje).lower()
    if "no se encontraron resultados" in msg:
        return "Producto no encontrado en la tabla"
    if "n_stock" in msg or "campo de stock" in msg:
        return "Modal de stock no se abrió"
    if "timeout" in msg or "timed out" in msg:
        return "Tiempo de espera agotado"
    if "no such element" in msg:
        return "Elemento de página no encontrado"
    if "stale" in msg:
        return "Página cambió durante la operación"
    return "Error inesperado"


def registrar_resultado(parte: str, stock, estado: str, detalle: str = "", duracion: float = 0.0):
    """Registra el resultado de un producto en la lista interna."""
    RESULTADOS.append({
        "Parte": parte,
        "Stock": stock,
        "Estado": estado,
        "Tipo de Fallo": clasificar_error(detalle) if estado == "FALLO" else "",
        "Descripción": detalle,
        "Duración (seg)": round(duracion, 1),
    })
    simbolo = "OK" if estado == "EXITO" else "FALLO"
    log.info(f"  [{simbolo}] {parte} -> {detalle}")


def generar_reporte_excel(acuerdo_texto: str = "", catalogo_texto: str = "", categoria_texto: str = ""):
    """Genera un reporte Excel profesional con resumen, detalle por producto y gráficos."""
    from collections import Counter
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.chart import PieChart, BarChart, Reference
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    total = len(RESULTADOS)
    exitos = sum(1 for r in RESULTADOS if r["Estado"] == "EXITO")
    fallos = total - exitos

    # ── Helpers de estilo ────────────────────────────────────────────────
    def fill(color):
        return PatternFill("solid", fgColor=color)

    def font(bold=False, size=10, color="000000", italic=False):
        return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

    borde = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    ac = Alignment(horizontal="center", vertical="center", wrap_text=False)
    al = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def celda(ws, row, col, value, fll=None, fnt=None, aln=None):
        c = ws.cell(row=row, column=col, value=value)
        if fll:
            c.fill = fll
        if fnt:
            c.font = fnt
        c.border = borde
        c.alignment = aln or ac
        return c

    # =====================================================================
    # HOJA 1: RESUMEN
    # =====================================================================
    ws_res = wb.active
    ws_res.title = "Resumen"
    ws_res.sheet_view.showGridLines = False
    ws_res.column_dimensions["A"].width = 30
    ws_res.column_dimensions["B"].width = 14
    ws_res.column_dimensions["C"].width = 14

    # Título
    ws_res.merge_cells("A1:C1")
    c = ws_res["A1"]
    c.value = "PERU COMPRAS BOT — Reporte de Actualización de Stock"
    c.font = font(bold=True, size=14, color="FFFFFF")
    c.fill = fill("00205B")
    c.alignment = ac
    ws_res.row_dimensions[1].height = 26

    # Fecha
    ws_res.merge_cells("A2:C2")
    c = ws_res["A2"]
    c.value = f"Generado el {datetime.now():%d/%m/%Y a las %H:%M:%S}"
    c.font = font(italic=True, size=9, color="595959")
    c.alignment = ac
    ws_res.row_dimensions[2].height = 14

    # Configuración
    ws_res.merge_cells("A3:C3")
    c = ws_res["A3"]
    c.value = f"Acuerdo: {acuerdo_texto}  |  Catálogo: {catalogo_texto}  |  Categoría: {categoria_texto}"
    c.font = font(italic=True, size=9, color="595959")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_res.row_dimensions[3].height = 14
    ws_res.row_dimensions[4].height = 8

    # Encabezado de tabla resumen
    for col, txt in enumerate(["Resultado", "Cantidad", "Porcentaje"], start=1):
        celda(ws_res, 5, col, txt, fill("1F4E79"), font(bold=True, size=11, color="FFFFFF"))
    ws_res.row_dimensions[5].height = 20

    pct_e = f"{exitos / total * 100:.1f}%" if total > 0 else "0.0%"
    pct_f = f"{fallos / total * 100:.1f}%" if total > 0 else "0.0%"

    for col, val in enumerate(["Exitosos ✔", exitos, pct_e], start=1):
        celda(ws_res, 6, col, val, fill("C6EFCE"), font(bold=True, size=10, color="375623"))
    ws_res.row_dimensions[6].height = 18

    for col, val in enumerate(["Fallidos ✘", fallos, pct_f], start=1):
        celda(ws_res, 7, col, val, fill("FFC7CE"), font(bold=True, size=10, color="9C0006"))
    ws_res.row_dimensions[7].height = 18

    for col, val in enumerate(["Total", total, "100%"], start=1):
        celda(ws_res, 8, col, val, fill("D9E1F2"), font(bold=True, size=10))
    ws_res.row_dimensions[8].height = 18

    # Gráfico de torta: Éxitos vs Fallos
    if total > 0:
        pie = PieChart()
        pie.title = "Resultados de Actualización"
        pie.style = 10
        pie.width = 13
        pie.height = 9
        data_ref = Reference(ws_res, min_col=2, min_row=5, max_row=7)
        cats_ref = Reference(ws_res, min_col=1, min_row=6, max_row=7)
        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(cats_ref)
        ws_res.add_chart(pie, "E1")

    # Desglose de tipos de fallo
    tipos = Counter(r["Tipo de Fallo"] for r in RESULTADOS if r["Estado"] == "FALLO")
    if tipos:
        ws_res.row_dimensions[10].height = 8
        for col, txt in enumerate(["Tipo de Fallo", "Cantidad"], start=1):
            celda(ws_res, 11, col, txt, fill("1F4E79"), font(bold=True, size=11, color="FFFFFF"))
        ws_res.row_dimensions[11].height = 20
        for i, (tipo, cnt) in enumerate(tipos.most_common(), start=12):
            celda(ws_res, i, 1, tipo, fill("FFC7CE"), font(size=10, color="9C0006"), aln=al)
            celda(ws_res, i, 2, cnt,  fill("FFC7CE"), font(bold=True, size=10, color="9C0006"))
            ws_res.row_dimensions[i].height = 16

        last_tipo_row = 11 + len(tipos)
        bar_tipos = BarChart()
        bar_tipos.type = "col"
        bar_tipos.title = "Detalle de Fallos por Categoría"
        bar_tipos.style = 10
        bar_tipos.y_axis.title = "Productos"
        bar_tipos.width = 13
        bar_tipos.height = 9
        data_b = Reference(ws_res, min_col=2, min_row=11, max_row=last_tipo_row)
        cats_b = Reference(ws_res, min_col=1, min_row=12, max_row=last_tipo_row)
        bar_tipos.add_data(data_b, titles_from_data=True)
        bar_tipos.set_categories(cats_b)
        ws_res.add_chart(bar_tipos, "E18")

    # =====================================================================
    # HOJA 2: DETALLE POR PRODUCTO
    # =====================================================================
    ws_det = wb.create_sheet("Detalle por Producto")
    ws_det.sheet_view.showGridLines = False

    ws_det.merge_cells("A1:G1")
    c = ws_det["A1"]
    c.value = "Detalle de Todos los Productos"
    c.font = font(bold=True, size=13, color="FFFFFF")
    c.fill = fill("1F4E79")
    c.alignment = ac
    ws_det.row_dimensions[1].height = 24

    det_headers = ["#", "Parte", "Stock", "Estado", "Tipo de Fallo", "Descripción del Error", "Duración (seg)"]
    for col, h in enumerate(det_headers, start=1):
        celda(ws_det, 2, col, h, fill("2E75B6"), font(bold=True, size=10, color="FFFFFF"))
    ws_det.row_dimensions[2].height = 18

    for i, r in enumerate(RESULTADOS, start=1):
        is_ok = r["Estado"] == "EXITO"
        fll = fill("C6EFCE") if is_ok else fill("FFC7CE")
        fnt_row = font(bold=True, size=9, color="375623" if is_ok else "9C0006")
        row = i + 2
        vals = [i, r["Parte"], r["Stock"], r["Estado"],
                r["Tipo de Fallo"] or "—", r["Descripción"], r["Duración (seg)"]]
        for col, val in enumerate(vals, start=1):
            celda(ws_det, row, col, val, fll, fnt_row, aln=al if col in (2, 5, 6) else ac)
        ws_det.row_dimensions[row].height = 16

    for col_letter, width in {"A": 5, "B": 22, "C": 8, "D": 10, "E": 30, "F": 42, "G": 14}.items():
        ws_det.column_dimensions[col_letter].width = width

    # Gráfico de barras: duración por producto
    if RESULTADOS:
        last_det = 2 + len(RESULTADOS)
        bar_dur = BarChart()
        bar_dur.type = "col"
        bar_dur.title = "Tiempo de Procesamiento por Producto (segundos)"
        bar_dur.style = 10
        bar_dur.y_axis.title = "Segundos"
        bar_dur.width = min(len(RESULTADOS) * 1.5 + 4, 28)
        bar_dur.height = 12
        data_dur = Reference(ws_det, min_col=7, min_row=2, max_row=last_det)
        cats_dur = Reference(ws_det, min_col=2, min_row=3, max_row=last_det)
        bar_dur.add_data(data_dur, titles_from_data=True)
        bar_dur.set_categories(cats_dur)
        ws_det.add_chart(bar_dur, f"A{last_det + 3}")

    # =====================================================================
    # HOJA 3: SOLO FALLIDOS
    # =====================================================================
    fallidos = [r for r in RESULTADOS if r["Estado"] == "FALLO"]
    ws_fail = wb.create_sheet("Solo Fallidos")
    ws_fail.sheet_view.showGridLines = False

    ws_fail.merge_cells("A1:F1")
    c = ws_fail["A1"]
    c.value = f"Productos Fallidos — {len(fallidos)} de {total}"
    c.font = font(bold=True, size=13, color="FFFFFF")
    c.fill = fill("9C0006")
    c.alignment = ac
    ws_fail.row_dimensions[1].height = 24

    if fallidos:
        fail_headers = ["#", "Parte", "Stock Intentado", "Tipo de Fallo", "Descripción del Error", "Duración (seg)"]
        for col, h in enumerate(fail_headers, start=1):
            celda(ws_fail, 2, col, h, fill("2E75B6"), font(bold=True, size=10, color="FFFFFF"))
        ws_fail.row_dimensions[2].height = 18

        for i, r in enumerate(fallidos, start=1):
            row = i + 2
            fll_alt = fill("FFE4E4") if i % 2 == 0 else fill("FFC7CE")
            vals = [i, r["Parte"], r["Stock"], r["Tipo de Fallo"], r["Descripción"], r["Duración (seg)"]]
            for col, val in enumerate(vals, start=1):
                celda(ws_fail, row, col, val, fll_alt, font(size=10, color="9C0006"),
                      aln=al if col in (2, 4, 5) else ac)
            ws_fail.row_dimensions[row].height = 18

        for col_letter, width in {"A": 5, "B": 22, "C": 14, "D": 32, "E": 48, "F": 14}.items():
            ws_fail.column_dimensions[col_letter].width = width
    else:
        ws_fail.merge_cells("A2:F2")
        c = ws_fail["A2"]
        c.value = "¡Sin fallos! Todos los productos se actualizaron correctamente."
        c.font = font(bold=True, size=11, color="375623")
        c.fill = fill("C6EFCE")
        c.alignment = ac
        ws_fail.row_dimensions[2].height = 22

    wb.save(REPORTE_PATH)
    log.info(f"Reporte Excel generado: {REPORTE_PATH}")
    return REPORTE_PATH


# ---------------------------------------------------------------------------
# HELPERS DE SELENIUM
# ---------------------------------------------------------------------------
def esperar_elemento(driver, by, valor, timeout=WAIT_NORMAL):
    """Espera a que un elemento sea visible y lo devuelve."""
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, valor))
    )


def esperar_clickeable(driver, by, valor, timeout=WAIT_NORMAL):
    """Espera a que un elemento sea clickeable y lo devuelve."""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, valor))
    )


def esperar_opciones_select(driver, select_id, timeout=WAIT_NORMAL):
    """Espera a que un <select> tenga mas de 1 opcion (la primera suele ser placeholder)."""
    def tiene_opciones(drv):
        sel = drv.find_element(By.ID, select_id)
        opciones = sel.find_elements(By.TAG_NAME, "option")
        return len(opciones) > 1

    WebDriverWait(driver, timeout).until(tiene_opciones)
    return Select(driver.find_element(By.ID, select_id))


def seleccionar_por_texto_parcial(select_obj: Select, texto_parcial: str):
    """Selecciona una opcion cuyo texto visible contenga el texto parcial dado."""
    for opcion in select_obj.options:
        if texto_parcial.upper() in opcion.text.upper():
            select_obj.select_by_visible_text(opcion.text)
            log.info(f"    Seleccionado: {opcion.text}")
            return True
    raise NoSuchElementException(
        f"No se encontro opcion que contenga '{texto_parcial}'"
    )


def aceptar_alerta(driver, timeout=WAIT_NORMAL):
    """Espera y acepta un alert/confirm del navegador."""
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        alerta = driver.switch_to.alert
        texto = alerta.text
        alerta.accept()
        log.info(f"    Alerta aceptada: {texto}")
        return texto
    except TimeoutException:
        log.warning("    No aparecio alerta en el tiempo esperado.")
        return None


def manejar_confirmacion_sweetalert(driver, timeout=WAIT_NORMAL):
    """
    Maneja dialogos tipo SweetAlert (muy comunes en este portal).
    Busca el boton de confirmacion "Si" / "Aceptar" / "Cerrar".
    """
    selectores_boton = [
        "//button[contains(@class,'swal2-confirm')]",
        "//button[contains(@class,'confirm')]",
        "//button[contains(text(),'Si')]",
        "//button[contains(text(),'Sí')]",
        "//button[contains(text(),'OK')]",
        "//button[contains(text(),'Aceptar')]",
        "//button[contains(text(),'Cerrar')]",
    ]
    for xpath in selectores_boton:
        try:
            btn = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            btn.click()
            log.info(f"    SweetAlert confirmado con: {xpath}")
            return True
        except TimeoutException:
            continue
    return False


def leer_opciones_select(driver, select_id, timeout=WAIT_NORMAL):
    """Lee las opciones de un <select> excluyendo el placeholder (value=0 o vacío)."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, select_id))
        )
        el = driver.find_element(By.ID, select_id)
        return [
            opt.text.strip()
            for opt in el.find_elements(By.TAG_NAME, "option")
            if opt.get_attribute("value") not in ("", "0") and opt.text.strip()
        ]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# PASO 1: LOGIN Y VERIFICACION
# ---------------------------------------------------------------------------
def paso1_login(driver):
    """Carga la pagina de login y espera la verificacion de 2 pasos."""
    log.info("=" * 60)
    log.info("PASO 1: Login y verificacion")
    log.info("=" * 60)

    driver.get(LOGIN_URL)
    log.info(f"Pagina cargada: {LOGIN_URL}")

    # --- Esperar a que el usuario llene credenciales manualmente ---
    if MODO_GUI and EVENTO_LOGIN is not None:
        log.info("Esperando confirmación de login desde la interfaz...")
        if GUI_NOTIFICAR_LOGIN:
            GUI_NOTIFICAR_LOGIN()
        EVENTO_LOGIN.clear()
        EVENTO_LOGIN.wait()
    else:
        print("\n" + "=" * 60)
        print("ACCION MANUAL REQUERIDA")
        print("=" * 60)
        print("1. Ingresa tu RUC, usuario y contrasena en el navegador.")
        print("2. Haz clic en 'Ingresar'.")
        print("3. Si aparece la ventana de verificacion de codigo, solo presiona ENTER.")
        input(">>> Presiona ENTER aqui cuando hayas completado el login... ")

    # --- Saltar verificacion haciendo click atras ---
    log.info("Saltando verificacion de codigo con navegacion hacia atras...")
    driver.back()
    time.sleep(3)
    log.info(f"Navegacion hacia atras completada. URL actual: {driver.current_url}")

    # Esperar a que cargue el dashboard
    time.sleep(2)
    log.info(f"Login completado. URL actual: {driver.current_url}")


# ---------------------------------------------------------------------------
# PASO 2: NAVEGACION Y TRUCO DE RETROCESO
# ---------------------------------------------------------------------------
def paso2_navegacion(driver):
    """Navega a la seccion de Mejora Basica desbloqueando el menu."""
    log.info("=" * 60)
    log.info("PASO 2: Navegacion a Mejora de Precio y Existencias")
    log.info("=" * 60)

    # Truco de retroceso: volver atras para desbloquear el menu
    driver.back()
    time.sleep(2)
    driver.get(BASE_URL)
    time.sleep(2)
    log.info("Retroceso ejecutado para desbloquear menu.")

    # Intentar navegar por URL directa
    driver.get(MEJORA_URL)
    time.sleep(3)

    # Verificar que estamos en la pagina correcta
    if "MejoraBasica" in driver.current_url:
        log.info(f"Navegacion exitosa a: {driver.current_url}")
        return

    # Fallback: intentar por menu lateral
    log.info("Intentando navegacion por menu lateral...")
    try:
        menu_mejora = esperar_clickeable(
            driver,
            By.XPATH,
            "//a[contains(text(),'Mejora de ofertas')]"
            "|//span[contains(text(),'Mejora de ofertas')]"
            "|//li[contains(text(),'Mejora de ofertas')]",
        )
        menu_mejora.click()
        time.sleep(1)

        submenu = esperar_clickeable(
            driver,
            By.XPATH,
            "//a[contains(@href,'MejoraBasica')]"
            "|//a[contains(text(),'Precio y Existencias')]"
            "|//a[contains(text(),'Precio y existencias')]",
        )
        submenu.click()
        time.sleep(2)
        log.info(f"Navegacion por menu completada. URL: {driver.current_url}")
    except TimeoutException:
        log.warning(
            "No se pudo navegar por menu. Intentando URL directa nuevamente..."
        )
        driver.get(MEJORA_URL)
        time.sleep(3)


# ---------------------------------------------------------------------------
# PASO 3: CONFIGURACION DE FILTROS
# ---------------------------------------------------------------------------
def paso3_filtros(driver):
    """Selecciona Acuerdo Marco, Catalogo y Categoria."""
    log.info("=" * 60)
    log.info("PASO 3: Configuracion de filtros")
    log.info("=" * 60)

    # Acuerdo Marco
    log.info("  Seleccionando Acuerdo Marco...")
    select_acuerdo = esperar_opciones_select(driver, "ajaxAcuerdo", WAIT_LARGO)
    seleccionar_por_texto_parcial(select_acuerdo, ACUERDO_TEXTO)
    time.sleep(2)  # Esperar carga del siguiente select

    # Catalogo Electronico
    log.info("  Seleccionando Catalogo Electronico...")
    select_catalogo = esperar_opciones_select(driver, "ajaxCatalogo", WAIT_LARGO)
    seleccionar_por_texto_parcial(select_catalogo, CATALOGO_TEXTO)
    time.sleep(2)

    # Categoria
    log.info("  Seleccionando Categoria...")
    select_categoria = esperar_opciones_select(driver, "ajaxCategoria", WAIT_LARGO)
    seleccionar_por_texto_parcial(select_categoria, CATEGORIA_TEXTO)
    time.sleep(1)

    log.info("Filtros configurados correctamente.")


# ---------------------------------------------------------------------------
# PASO 4: BUCLE DE ACTUALIZACION DE STOCK
# ---------------------------------------------------------------------------
def paso4_actualizar_stock(driver, df: pd.DataFrame):
    """Itera sobre cada producto del DataFrame y actualiza el stock."""
    log.info("=" * 60)
    log.info(f"PASO 4: Actualizacion de stock ({len(df)} productos)")
    log.info("=" * 60)

    total = len(df)
    exitos = 0
    fallos = 0

    for idx, fila in df.iterrows():
        parte = str(fila["Parte"]).strip()
        stock = str(int(fila["Stock"]))
        log.info(f"\n--- Producto {idx + 1}/{total}: Parte={parte}, Stock={stock} ---")

        t_inicio = time.time()
        try:
            actualizar_producto(driver, parte, stock)
            registrar_resultado(parte, stock, "EXITO", "Stock actualizado correctamente",
                                duracion=time.time() - t_inicio)
            exitos += 1
        except Exception as e:
            log.error(f"  Error procesando {parte}: {e}")
            registrar_resultado(parte, stock, "FALLO", str(e),
                                duracion=time.time() - t_inicio)
            fallos += 1
            # Intentar recuperar el estado limpio para el siguiente producto
            try:
                recuperar_estado(driver)
            except Exception:
                log.warning("  No se pudo recuperar estado. Recargando pagina...")
                driver.get(MEJORA_URL)
                time.sleep(3)
                paso3_filtros(driver)

        time.sleep(PAUSA_ENTRE_PRODUCTOS)

    log.info("=" * 60)
    log.info(f"RESULTADO FINAL: {exitos} exitosos, {fallos} fallidos de {total} total")
    log.info(f"Reporte guardado en: {REPORTE_PATH}")
    log.info("=" * 60)


def actualizar_producto(driver, parte: str, stock: str):
    """Busca un producto por numero de parte y actualiza su stock."""

    # 1. Escribir numero de parte en el campo de busqueda
    campo_busqueda = esperar_elemento(driver, By.ID, "C_Descripcion")
    campo_busqueda.clear()
    campo_busqueda.send_keys(parte)

    # 2. Click en buscar
    btn_buscar = esperar_clickeable(driver, By.ID, "btnBuscar")
    btn_buscar.click()
    log.info(f"  Busqueda lanzada para: {parte}")

    # 3. Esperar a que cargue la tabla de resultados
    time.sleep(3)  # Pausa inicial para carga AJAX

    try:
        WebDriverWait(driver, WAIT_NORMAL).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//table//tbody//tr[contains(@class,'row')]"
                "|//table//tbody//tr[td]"
                "|//div[contains(@class,'dataTables')]//tbody//tr"
            ))
        )
    except TimeoutException:
        raise TimeoutException(f"No se encontraron resultados para '{parte}'")

    # 4. Localizar el boton "Existencias" DENTRO DE LA TABLA de resultados
    # Ser muy específico: buscar en la tabla, el enlace con fnModificarStock
    xpath_existencias = (
        f"//table//tbody//tr//a[contains(@onclick,'fnModificarStock')]"
        f"|//table//tr//a[contains(@onclick,'fnModificarStock')]"
        f"|//div[contains(@class,'dataTables')]//a[contains(@onclick,'fnModificarStock')]"
    )

    btn_existencias = esperar_clickeable(
        driver, By.XPATH, xpath_existencias, WAIT_NORMAL
    )
    
    # Scroll al elemento antes de hacer click
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_existencias)
    time.sleep(0.5)
    
    # DEBUGGING: Capturar info del botón encontrado
    try:
        tag_name = btn_existencias.tag_name
        onclick_value = btn_existencias.get_attribute("onclick")
        text = btn_existencias.text
        html = btn_existencias.get_attribute("outerHTML")
        
        log.info(f"  Botón encontrado (VERIFICAR):")
        log.info(f"    Tag: {tag_name}")
        log.info(f"    Texto: {text}")
        log.info(f"    onclick: {onclick_value[:100] if onclick_value else 'None'}")
        log.info(f"    HTML: {html[:200] if len(html) > 200 else html}")
    except Exception as e:
        log.error(f"  Error obteniendo info del botón: {e}")
    
    # NUEVA ESTRATEGIA: Ejecutar directamente fnModificarStock() en lugar de hacer click
    # Primero, intentar obtener el onclick attribute
    onclick_value = btn_existencias.get_attribute("onclick")
    
    # Extraer el ID del producto del onclick (ej: fnModificarStock(57072189))
    id_producto = None
    if onclick_value:
        import re
        match = re.search(r'fnModificarStock\((\d+)\)', onclick_value)
        if match:
            id_producto = match.group(1)
            log.info(f"  ✓ ID del producto extraído: {id_producto}")
        else:
            log.warning(f"  ✗ No se pudo extraer ID desde onclick: {onclick_value}")
    else:
        log.warning(f"  ✗ El atributo onclick es None o no existe")
    
    # Si se extrajo el ID, ejecutar directamente fnModificarStock
    if id_producto:
        log.info(f"  → Ejecutando fnModificarStock({id_producto}) directamente...")
        try:
            # Intentar ejecutar la función
            resultado = driver.execute_script(f"return fnModificarStock({id_producto});")
            log.info(f"  ✓ Función fnModificarStock({id_producto}) ejecutada correctamente")
        except Exception as e:
            log.error(f"  ✗ Error ejecutando fnModificarStock: {e}")
    else:
        # Si no se extrajo el ID, hacer click normal
        log.info("  → No se pudo extraer ID, haciendo click normal...")
        try:
            btn_existencias.click()
            log.info("  ✓ Click realizado")
        except Exception as e:
            log.error(f"  ✗ Error en click: {e}")
    
    log.info("  → Esperando que aparezca el modal...")
    time.sleep(3)

    # 5. Esperar a que el formulario del modal aparezca de múltiples formas
    form_encontrado = False
    
    # Intentar 1: Buscar directamente por ID formStock
    try:
        WebDriverWait(driver, WAIT_CORTO).until(
            EC.presence_of_element_located((By.ID, "formStock"))
        )
        log.info("  Formulario formStock encontrado en DOM")
        form_encontrado = True
    except TimeoutException:
        log.warning("  Primer intento: formStock no encontrado")
    
    # Intentar 2: Buscar usando presencia de atributo action="/MejoraBasica/ModificarStock"
    if not form_encontrado:
        try:
            formulario = WebDriverWait(driver, WAIT_CORTO).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    "//form[contains(@action,'ModificarStock')]"
                ))
            )
            log.info("  Formulario encontrado por action=/MejoraBasica/ModificarStock")
            form_encontrado = True
        except TimeoutException:
            log.warning("  Segundo intento: formulario con action no encontrado")
    
    # Intentar 3: Buscar cualquier form que contenga un campo N_Stock
    if not form_encontrado:
        try:
            formulario = WebDriverWait(driver, WAIT_CORTO).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    "//form[.//input[@name='N_Stock']]"
                ))
            )
            log.info("  Formulario encontrado por input N_Stock")
            form_encontrado = True
        except TimeoutException:
            log.warning("  Tercer intento: formulario con input N_Stock no encontrado")
    
    # Capturar el HTML actual para debugging
    html_actual = driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML")
    if "formStock" in html_actual:
        log.info("  INFO: 'formStock' EXISTE en el HTML actual")
    else:
        log.warning("  INFO: 'formStock' NO existe en el HTML actual")
    
    if "N_Stock" in html_actual:
        log.info("  INFO: 'N_Stock' EXISTE en el HTML actual")
        # Mostrar contexto del N_Stock
        import re
        match = re.search(r'.{0,150}N_Stock.{0,150}', html_actual)
        if match:
            log.info(f"  Contexto N_Stock: ...{match.group(0)}...")
    else:
        log.warning("  INFO: 'N_Stock' NO existe en el HTML actual")
    
    time.sleep(1)

    # 6. Buscar el campo N_Stock 
    campo_stock = None
    try:
        campo_stock = driver.find_element(By.ID, "N_Stock")
        log.info("  Campo N_Stock encontrado por ID")
    except NoSuchElementException:
        log.warning("  Campo N_Stock no encontrado por ID, buscando en el formulario...")
        try:
            campo_stock = driver.find_element(By.CSS_SELECTOR, "#formStock #N_Stock")
            log.info("  Campo N_Stock encontrado en formulario")
        except NoSuchElementException:
            pass
    
    if campo_stock is None:
        # Último intento: buscar visible en el modal
        try:
            modal = driver.find_element(By.CSS_SELECTOR, ".modal.show, .modal.in, [style*='display: block']")
            inputs = modal.find_elements(By.CSS_SELECTOR, "input[id='N_Stock']")
            if inputs:
                campo_stock = inputs[0]
                log.info("  Campo N_Stock encontrado en modal visible")
        except:
            pass
    
    if campo_stock is None:
        # Mostrar el HTML del body para debugging
        try:
            body_html = driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML")
            if "N_Stock" in body_html:
                log.warning("  El campo N_Stock EXISTE en el HTML pero no se puede acceder")
            else:
                log.warning("  El campo N_Stock NO existe en el HTML actual")
        except:
            pass
        raise TimeoutException("No se encontro el campo de stock (N_Stock)")

    # 7. Limpiar el campo y escribir el nuevo stock
    log.info(f"  Limpiando campo y escribiendo stock: {stock}")
    campo_stock.clear()
    time.sleep(0.5)
    campo_stock.send_keys(stock)
    log.info(f"  Stock ingresado: {stock}")

    time.sleep(1)

    # 8. Buscar y hacer click en el boton de guardar
    btn_guardar = None
    selectores_guardar = [
        (By.ID, "btn_guardar"),
        (By.XPATH, "//button[@type='submit']"),
        (By.XPATH, "//button[contains(text(),'Guardar')]"),
    ]
    
    for selector_type, selector in selectores_guardar:
        try:
            btn_guardar = WebDriverWait(driver, WAIT_CORTO).until(
                EC.element_to_be_clickable((selector_type, selector))
            )
            log.info(f"  Boton guardar encontrado, haciendo click...")
            btn_guardar.click()
            log.info(f"  Click en guardar realizado")
            break
        except TimeoutException:
            continue

    if btn_guardar is None:
        log.warning("  No se encontro boton guardar. Intentando presionar Enter...")
        campo_stock.send_keys("\n")

    time.sleep(2)

    # 9. Manejar confirmacion "¿Esta seguro de registrar los cambios realizados?"
    confirmado = False

    # Intentar alert nativo primero
    try:
        alerta = WebDriverWait(driver, WAIT_CORTO).until(EC.alert_is_present())
        log.info(f"  Alerta detectada, aceptando...")
        alerta.accept()
        confirmado = True
    except TimeoutException:
        pass

    # Si no fue alert nativo, buscar el boton "Sí" del modal personalizado
    if not confirmado:
        log.info("  Buscando boton de confirmacion en modal...")
        selectores_si = [
            (By.XPATH, "//div[contains(@class,'_wModal_btn_ok')]"),
            (By.XPATH, "//div[contains(@class,'_wModal_btn') and contains(text(),'Sí')]"),
            (By.XPATH, "//button[contains(@class,'_wModal_btn_ok')]"),
            (By.XPATH, "//button[contains(text(),'Sí')]"),
        ]
        
        for selector_type, selector in selectores_si:
            try:
                btn_si = WebDriverWait(driver, WAIT_CORTO).until(
                    EC.element_to_be_clickable((selector_type, selector))
                )
                log.info(f"  Boton Sí encontrado, haciendo click...")
                btn_si.click()
                log.info(f"  Confirmacion aceptada")
                confirmado = True
                break
            except TimeoutException:
                continue

    if not confirmado:
        log.warning("  No se detecto dialogo de confirmacion")

    time.sleep(2)

    # 10. Cerrar cualquier modal abierto
    try:
        btn_cerrar = WebDriverWait(driver, WAIT_CORTO).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[contains(@class,'close')]"
                "|//button[@data-dismiss='modal']"
                "|//div[contains(@class,'_wModal_close')]"
            ))
        )
        btn_cerrar.click()
        log.info("  Modal cerrado")
    except TimeoutException:
        log.info("  No hay modal para cerrar")

    time.sleep(2)
    log.info(f"  Producto {parte} actualizado exitosamente")


def recuperar_estado(driver):
    """Intenta cerrar modales abiertos y limpiar la busqueda."""
    # Cerrar cualquier modal abierto
    try:
        modales_cerrar = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'modal')]//button[contains(@class,'close')]"
            "|//button[@data-dismiss='modal']",
        )
        for btn in modales_cerrar:
            try:
                btn.click()
                time.sleep(0.5)
            except Exception:
                pass
    except Exception:
        pass

    # Cerrar alertas pendientes
    try:
        driver.switch_to.alert.accept()
    except NoAlertPresentException:
        pass

    # Limpiar campo de busqueda
    try:
        campo = driver.find_element(By.ID, "C_Descripcion")
        campo.clear()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# EJECUCION REUTILIZABLE (CLI + GUI)
# ---------------------------------------------------------------------------
def ejecutar_bot(
    excel_path: Path,
    acuerdo_texto: str,
    catalogo_texto: str,
    categoria_texto: str,
    pausa_entre_productos: int = 2,
):
    """Ejecuta el flujo completo del bot con configuración dinámica."""
    global EXCEL_PATH, ACUERDO_TEXTO, CATALOGO_TEXTO, CATEGORIA_TEXTO
    global PAUSA_ENTRE_PRODUCTOS, REPORTE_PATH, RESULTADOS
    RESULTADOS = []

    EXCEL_PATH = Path(excel_path)
    ACUERDO_TEXTO = acuerdo_texto
    CATALOGO_TEXTO = catalogo_texto
    CATEGORIA_TEXTO = categoria_texto
    PAUSA_ENTRE_PRODUCTOS = pausa_entre_productos
    REPORTE_PATH = nueva_ruta_reporte()

    log.info("Iniciando Peru Compras Bot...")
    log.info(f"Archivo Excel seleccionado: {EXCEL_PATH}")
    log.info(f"Filtro acuerdo: {ACUERDO_TEXTO}")
    log.info(f"Filtro catálogo: {CATALOGO_TEXTO}")
    log.info(f"Filtro categoría: {CATEGORIA_TEXTO}")

    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontro el archivo Excel: {EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH)
    if "Parte" not in df.columns or "Stock" not in df.columns:
        raise ValueError("El Excel debe tener columnas 'Parte' y 'Stock'.")

    df = df.dropna(subset=["Parte", "Stock"]).reset_index(drop=True)
    log.info(f"Productos cargados: {len(df)}")
    if len(df) == 0:
        raise ValueError("No hay productos para procesar en el Excel.")

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        paso1_login(driver)
        paso2_navegacion(driver)
        paso3_filtros(driver)
        paso4_actualizar_stock(driver, df)
        generar_reporte_excel(acuerdo_texto, catalogo_texto, categoria_texto)
        return REPORTE_PATH
    finally:
        driver.quit()
        log.info("Navegador cerrado. Fin del script.")


def main_cli():
    """Modo clásico por consola."""
    global MODO_GUI, EVENTO_LOGIN, GUI_NOTIFICAR_LOGIN
    MODO_GUI = False
    EVENTO_LOGIN = None
    GUI_NOTIFICAR_LOGIN = None

    try:
        reporte = ejecutar_bot(
            excel_path=EXCEL_PATH,
            acuerdo_texto=ACUERDO_TEXTO,
            catalogo_texto=CATALOGO_TEXTO,
            categoria_texto=CATEGORIA_TEXTO,
            pausa_entre_productos=PAUSA_ENTRE_PRODUCTOS,
        )
        print("\n" + "=" * 60)
        print(f"Reporte guardado en: {reporte}")
        input("Presiona ENTER para cerrar...")
    except KeyboardInterrupt:
        log.info("\nEjecucion interrumpida por el usuario.")
    except Exception as e:
        log.error(f"Error fatal: {e}", exc_info=True)


class TextQueueLogHandler(logging.Handler):
    def __init__(self, cola: Queue):
        super().__init__()
        self.cola = cola

    def emit(self, record):
        try:
            msg = self.format(record)
            self.cola.put(msg)
        except Exception:
            pass


class PeruComprasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Peru Compras Bot - Interfaz")
        self.root.geometry("980x700")

        self.log_queue = Queue()
        self.login_event = threading.Event()
        self.worker = None
        self.reporte_generado = None

        self.excel_var = tk.StringVar(value=str(Path(__file__).parent / "productos.xlsx"))
        self.acuerdo_var = tk.StringVar(value=ACUERDO_TEXTO)
        self.catalogo_var = tk.StringVar(value=CATALOGO_TEXTO)
        self.categoria_var = tk.StringVar(value=CATEGORIA_TEXTO)
        self.pausa_var = tk.StringVar(value=str(PAUSA_ENTRE_PRODUCTOS))
        self.estado_var = tk.StringVar(value="Listo para iniciar")

        self._build_ui()
        self._configurar_logging_gui()
        self._tick_logs()

    def _build_ui(self):
        contenedor = ttk.Frame(self.root, padding=12)
        contenedor.pack(fill="both", expand=True)

        titulo = ttk.Label(
            contenedor,
            text="Peru Compras Bot (uso por clicks)",
            font=("Segoe UI", 14, "bold"),
        )
        titulo.pack(anchor="w", pady=(0, 8))

        estado = ttk.Label(contenedor, textvariable=self.estado_var)
        estado.pack(anchor="w", pady=(0, 10))

        frame_excel = ttk.LabelFrame(contenedor, text="Archivo Excel", padding=10)
        frame_excel.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_excel, text="Ruta del Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_excel, textvariable=self.excel_var, width=85).grid(
            row=1, column=0, padx=(0, 8), pady=(4, 0), sticky="we"
        )
        ttk.Button(frame_excel, text="Seleccionar Excel", command=self._seleccionar_excel).grid(
            row=1, column=1, pady=(4, 0)
        )
        frame_excel.columnconfigure(0, weight=1)

        frame_filtros = ttk.LabelFrame(contenedor, text="Filtros", padding=10)
        frame_filtros.pack(fill="x", pady=(0, 10))

        ttk.Label(frame_filtros, text="Acuerdo Marco:").grid(row=0, column=0, sticky="w")
        self.combo_acuerdo = ttk.Combobox(frame_filtros, textvariable=self.acuerdo_var, width=78)
        self.combo_acuerdo.grid(row=0, column=1, sticky="we", padx=(8, 0), pady=(0, 6))
        self.combo_acuerdo.bind("<<ComboboxSelected>>", self._on_acuerdo_changed)

        ttk.Label(frame_filtros, text="Catálogo:").grid(row=1, column=0, sticky="w")
        self.combo_catalogo = ttk.Combobox(frame_filtros, textvariable=self.catalogo_var, width=78)
        self.combo_catalogo.grid(row=1, column=1, sticky="we", padx=(8, 0), pady=(0, 6))
        self.combo_catalogo.bind("<<ComboboxSelected>>", self._on_catalogo_changed)

        ttk.Label(frame_filtros, text="Categoría:").grid(row=2, column=0, sticky="w")
        self.combo_categoria = ttk.Combobox(frame_filtros, textvariable=self.categoria_var, width=78)
        self.combo_categoria.grid(row=2, column=1, sticky="we", padx=(8, 0), pady=(0, 6))

        ttk.Label(frame_filtros, text="Pausa entre productos (seg):").grid(
            row=3, column=0, sticky="w"
        )
        ttk.Entry(frame_filtros, textvariable=self.pausa_var, width=10).grid(
            row=3, column=1, sticky="w", padx=(8, 0)
        )

        self.btn_cargar_opts = ttk.Button(
            frame_filtros, text="Cargar opciones del portal", command=self._cargar_opciones
        )
        self.btn_cargar_opts.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="w")
        ttk.Label(
            frame_filtros,
            text="Inicia sesión manualmente cuando Chrome se abra, igual que al iniciar la automatización.",
            foreground="#666666",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 0))

        frame_filtros.columnconfigure(1, weight=1)

        frame_acciones = ttk.Frame(contenedor)
        frame_acciones.pack(fill="x", pady=(0, 10))

        self.btn_iniciar = ttk.Button(frame_acciones, text="Iniciar automatización", command=self._iniciar)
        self.btn_iniciar.pack(side="left", padx=(0, 8))

        self.btn_login = ttk.Button(
            frame_acciones,
            text="Ya inicié sesión (continuar)",
            command=self._continuar_login,
            state="disabled",
        )
        self.btn_login.pack(side="left", padx=(0, 8))

        ttk.Button(frame_acciones, text="Abrir carpeta del proyecto", command=self._abrir_carpeta).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(frame_acciones, text="Abrir último reporte", command=self._abrir_reporte).pack(side="left")

        frame_log = ttk.LabelFrame(contenedor, text="Ejecución en tiempo real", padding=10)
        frame_log.pack(fill="both", expand=True)

        self.txt_log = scrolledtext.ScrolledText(frame_log, height=20, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True)

    def _configurar_logging_gui(self):
        self.gui_handler = TextQueueLogHandler(self.log_queue)
        self.gui_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        log.addHandler(self.gui_handler)

    def _tick_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.txt_log.configure(state="normal")
                self.txt_log.insert("end", msg + "\n")
                self.txt_log.see("end")
                self.txt_log.configure(state="disabled")
                if "PASO" in msg:
                    self.estado_var.set(msg.split("] ", 1)[-1])
        except Empty:
            pass
        self.root.after(200, self._tick_logs)

    def _seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls")],
            initialdir=str(Path(__file__).parent),
        )
        if ruta:
            self.excel_var.set(ruta)

    def _abrir_carpeta(self):
        os.startfile(str(Path(__file__).parent))

    def _abrir_reporte(self):
        if self.reporte_generado and Path(self.reporte_generado).exists():
            os.startfile(str(self.reporte_generado))
            return
        messagebox.showinfo("Reporte", "Aún no hay reporte generado en esta sesión.")

    def _continuar_login(self):
        self.login_event.set()
        self.btn_login.configure(state="disabled")
        self.estado_var.set("Login confirmado. Continuando proceso...")

    def _notificar_login_ui(self):
        def _update():
            self.btn_login.configure(state="normal")
            self.estado_var.set("Completa login en Chrome y luego haz click en 'Ya inicié sesión'")
            messagebox.showinfo(
                "Acción requerida",
                "1) Inicia sesión en la web en la ventana de Chrome.\n"
                "2) Haz clic en 'Ingresar'.\n"
                "3) Regresa aquí y haz clic en 'Ya inicié sesión (continuar)'.",
            )

        self.root.after(0, _update)

    # ------------------------------------------------------------------
    # Cascada de filtros
    # ------------------------------------------------------------------
    def _on_acuerdo_changed(self, event=None):
        """Al cambiar el Acuerdo, los catálogos/categorías ya no son válidos."""
        self.combo_catalogo["values"] = []
        self.catalogo_var.set("")
        self.combo_categoria["values"] = []
        self.categoria_var.set("")
        self.estado_var.set("Acuerdo cambiado — haz clic en 'Cargar opciones del portal' para actualizar los filtros")

    def _on_catalogo_changed(self, event=None):
        """Al cambiar el Catálogo, la categoría ya no es válida."""
        self.combo_categoria["values"] = []
        self.categoria_var.set("")
        self.estado_var.set("Catálogo cambiado — haz clic en 'Cargar opciones del portal' para actualizar categorías")

    # ------------------------------------------------------------------
    # Cargar opciones desde el portal
    # ------------------------------------------------------------------
    def _cargar_opciones(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Ocupado", "Espera a que termine el proceso en curso.")
            return
        self.login_event.clear()
        self.btn_cargar_opts.configure(state="disabled")
        self.btn_iniciar.configure(state="disabled")
        self.btn_login.configure(state="disabled")
        self.estado_var.set("Conectando al portal para cargar opciones...")
        self.worker = threading.Thread(target=self._cargar_opciones_worker, daemon=True)
        self.worker.start()

    def _cargar_opciones_worker(self):
        global MODO_GUI, EVENTO_LOGIN, GUI_NOTIFICAR_LOGIN
        MODO_GUI = True
        EVENTO_LOGIN = self.login_event
        GUI_NOTIFICAR_LOGIN = self._notificar_login_ui
        driver = None
        try:
            chrome_opts = Options()
            chrome_opts.add_argument("--start-maximized")
            chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
            chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_opts.add_experimental_option("useAutomationExtension", False)
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_opts)

            paso1_login(driver)
            paso2_navegacion(driver)

            # --- Acuerdo Marco ---
            acuerdo_opts = leer_opciones_select(driver, "ajaxAcuerdo")
            log.info(f"Opciones Acuerdo ({len(acuerdo_opts)}): {acuerdo_opts}")

            # Seleccionar el Acuerdo actual (o el primero disponible)
            catalogo_opts = []
            categoria_opts = []
            acuerdo_actual = self.acuerdo_var.get().strip()
            if acuerdo_opts:
                try:
                    sel_a = esperar_opciones_select(driver, "ajaxAcuerdo", WAIT_LARGO)
                    texto_a = acuerdo_actual if acuerdo_actual else acuerdo_opts[0]
                    seleccionar_por_texto_parcial(sel_a, texto_a)
                    time.sleep(2)
                    catalogo_opts = leer_opciones_select(driver, "ajaxCatalogo")
                    log.info(f"Opciones Catálogo ({len(catalogo_opts)}): {catalogo_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar catálogos: {e}")

            # --- Catálogo → Categoría ---
            catalogo_actual = self.catalogo_var.get().strip()
            if catalogo_opts:
                try:
                    sel_c = esperar_opciones_select(driver, "ajaxCatalogo", WAIT_LARGO)
                    texto_c = catalogo_actual if catalogo_actual else catalogo_opts[0]
                    seleccionar_por_texto_parcial(sel_c, texto_c)
                    time.sleep(2)
                    categoria_opts = leer_opciones_select(driver, "ajaxCategoria")
                    log.info(f"Opciones Categoría ({len(categoria_opts)}): {categoria_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar categorías: {e}")

            self.root.after(0, lambda: self._actualizar_combos(acuerdo_opts, catalogo_opts, categoria_opts))

        except Exception as e:
            err = str(e)
            log.error(f"Error cargando opciones del portal: {e}", exc_info=True)
            self.root.after(0, lambda: messagebox.showerror(
                "Error", f"No se pudieron cargar las opciones del portal:\n{err}"
            ))
            self.root.after(0, lambda: self.estado_var.set("Error al cargar opciones"))
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.root.after(0, lambda: self.btn_cargar_opts.configure(state="normal"))
            self.root.after(0, lambda: self.btn_iniciar.configure(state="normal"))
            self.root.after(0, lambda: self.btn_login.configure(state="disabled"))

    def _actualizar_combos(self, acuerdos, catalogos, categorias):
        self.combo_acuerdo["values"] = acuerdos
        self.combo_catalogo["values"] = catalogos
        self.combo_categoria["values"] = categorias

        # Asignar primer valor si el campo estaba vacío
        if acuerdos and not self.acuerdo_var.get():
            self.acuerdo_var.set(acuerdos[0])
        if catalogos and not self.catalogo_var.get():
            self.catalogo_var.set(catalogos[0])
        if categorias and not self.categoria_var.get():
            self.categoria_var.set(categorias[0])

        self.estado_var.set(
            f"Listo — {len(acuerdos)} acuerdos, {len(catalogos)} catálogos, {len(categorias)} categorías cargados"
        )
        messagebox.showinfo(
            "Opciones cargadas",
            f"Se cargaron correctamente desde el portal:\n"
            f"  • {len(acuerdos)} Acuerdo(s) Marco\n"
            f"  • {len(catalogos)} Catálogo(s)\n"
            f"  • {len(categorias)} Categoría(s)\n\n"
            f"Ahora puedes seleccionar cada filtro desde el desplegable.",
        )

    def _iniciar(self):
        excel = Path(self.excel_var.get().strip())
        acuerdo = self.acuerdo_var.get().strip()
        catalogo = self.catalogo_var.get().strip()
        categoria = self.categoria_var.get().strip()
        pausa_txt = self.pausa_var.get().strip()

        if not excel.exists():
            messagebox.showerror("Validación", "El archivo Excel no existe. Selecciona un archivo válido.")
            return
        if not acuerdo or not catalogo or not categoria:
            messagebox.showerror("Validación", "Completa Acuerdo, Catálogo y Categoría.")
            return
        try:
            pausa = int(pausa_txt)
            if pausa < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Validación", "La pausa debe ser un número entero >= 0.")
            return

        self.btn_iniciar.configure(state="disabled")
        self.btn_login.configure(state="disabled")
        self.estado_var.set("Iniciando automatización...")

        self.worker = threading.Thread(
            target=self._worker_run,
            args=(excel, acuerdo, catalogo, categoria, pausa),
            daemon=True,
        )
        self.worker.start()

    def _worker_run(self, excel, acuerdo, catalogo, categoria, pausa):
        global MODO_GUI, EVENTO_LOGIN, GUI_NOTIFICAR_LOGIN
        MODO_GUI = True
        EVENTO_LOGIN = self.login_event
        GUI_NOTIFICAR_LOGIN = self._notificar_login_ui

        try:
            reporte = ejecutar_bot(
                excel_path=excel,
                acuerdo_texto=acuerdo,
                catalogo_texto=catalogo,
                categoria_texto=categoria,
                pausa_entre_productos=pausa,
            )
            self.reporte_generado = reporte
            self.root.after(0, lambda: self.estado_var.set(f"Finalizado. Reporte: {reporte}"))
            self.root.after(0, lambda: messagebox.showinfo("Proceso completado", f"Proceso finalizado.\nReporte: {reporte}"))
        except Exception as e:
            detalle = f"{e}\n\n{traceback.format_exc()}"
            log.error(f"Error fatal: {e}", exc_info=True)
            self.root.after(0, lambda: self.estado_var.set("Error en la ejecución"))
            self.root.after(0, lambda: messagebox.showerror("Error", detalle))
        finally:
            self.root.after(0, lambda: self.btn_iniciar.configure(state="normal"))
            self.root.after(0, lambda: self.btn_login.configure(state="disabled"))


def iniciar_interfaz():
    root = tk.Tk()
    PeruComprasGUI(root)
    root.mainloop()


if __name__ == "__main__":
    if "--cli" in sys.argv:
        main_cli()
    else:
        iniciar_interfaz()
