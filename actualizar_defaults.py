"""
actualizar_defaults.py
======================
Conecta al portal de Peru Compras, lee las opciones reales de todos los
selectores (Acuerdo → Catálogo → Categoría) y sobreescribe PORTAL_DEFAULTS
en automation.py con los valores actuales.

USO:
    python actualizar_defaults.py

Pasos automáticos:
  1. Abre Chrome
  2. Espera que hagas login manual
  3. Navega a MejoraBasica y lee acuerdos / catálogos / categorías
  4. Reescribe PORTAL_DEFAULTS en automation.py
  5. Cierra Chrome

Después de ejecutarlo, reconstruye el instalador con:
    build_exe.bat
    build_installer.bat
"""

import re
import sys
import time
import textwrap
from pathlib import Path

# Asegurarse de importar desde la carpeta del proyecto
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import peru_compras_bot_app.automation as bot

# ---------------------------------------------------------------------------
# Constantes del portal
# ---------------------------------------------------------------------------
WAIT = 45   # segundos de espera para selectores lentos

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _leer(driver, select_id: str) -> list[str]:
    """Lee opciones de un <select> ignorando placeholders."""
    try:
        WebDriverWait(driver, WAIT).until(
            EC.presence_of_element_located((By.ID, select_id))
        )
        el = driver.find_element(By.ID, select_id)
        return [
            opt.text.strip()
            for opt in el.find_elements(By.TAG_NAME, "option")
            if opt.get_attribute("value") not in ("", "0") and opt.text.strip()
        ]
    except Exception as exc:
        print(f"  [WARN] No se pudo leer #{select_id}: {exc}")
        return []


def _seleccionar(driver, select_id: str, texto: str, timeout: int = WAIT) -> bool:
    """Selecciona la opción cuyo texto contenga 'texto'."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len([
                o for o in d.find_element(By.ID, select_id)
                               .find_elements(By.TAG_NAME, "option")
                if o.get_attribute("value") not in ("", "0")
            ]) > 0
        )
        sel = Select(driver.find_element(By.ID, select_id))
        for opt in sel.options:
            if texto.upper() in opt.text.upper():
                sel.select_by_visible_text(opt.text)
                return True
        # fallback exacto
        sel.select_by_visible_text(texto)
        return True
    except Exception as exc:
        print(f"  [WARN] No se pudo seleccionar '{texto}' en #{select_id}: {exc}")
        return False


def _esperar_cascade(driver, select_id: str, timeout: int = WAIT) -> list[str]:
    """Espera a que el select se recargue y devuelve sus opciones."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len([
                o for o in d.find_element(By.ID, select_id)
                               .find_elements(By.TAG_NAME, "option")
                if o.get_attribute("value") not in ("", "0")
            ]) > 0
        )
        return _leer(driver, select_id)
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Scraping principal
# ---------------------------------------------------------------------------

def scrape_defaults() -> dict:
    chrome_opts = Options()
    chrome_opts.add_argument("--start-maximized")
    chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
    chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_opts.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=chrome_opts)

    try:
        # ── LOGIN MANUAL ────────────────────────────────────────────────
        bot.MODO_GUI = False
        bot.step1_done = False
        bot.paso1_login(driver)          # mostrará prompt en consola

        # ── NAVEGACIÓN ──────────────────────────────────────────────────
        print("\n[1/4] Navegando a MejoraBasica...")
        bot.paso2_navegacion(driver)
        time.sleep(3)

        # ── ACUERDOS ────────────────────────────────────────────────────
        print("[2/4] Leyendo acuerdos...")
        acuerdos = _leer(driver, "ajaxAcuerdo")
        if not acuerdos:
            raise RuntimeError("No se encontraron acuerdos. Verifica que el login fue correcto.")
        print(f"  → {len(acuerdos)} acuerdo(s): {acuerdos}")

        # ── CATÁLOGOS Y CATEGORÍAS por acuerdo ──────────────────────────
        print("[3/4] Iterando catálogos y categorías...")
        catalogo_por_acuerdo: dict[str, list[str]] = {}
        categoria_por_catalogo: dict[str, list[str]] = {}

        for acuerdo in acuerdos:
            print(f"\n  Acuerdo: {acuerdo}")
            if not _seleccionar(driver, "ajaxAcuerdo", acuerdo):
                catalogo_por_acuerdo[acuerdo] = []
                continue
            time.sleep(2)

            catalogos = _esperar_cascade(driver, "ajaxCatalogo")
            print(f"    → {len(catalogos)} catálogo(s): {catalogos}")
            catalogo_por_acuerdo[acuerdo] = catalogos

            for cat in catalogos:
                if not _seleccionar(driver, "ajaxCatalogo", cat):
                    categoria_por_catalogo[cat] = []
                    continue
                time.sleep(2)
                categorias = _esperar_cascade(driver, "ajaxCategoria")
                print(f"      Catálogo '{cat}' → {len(categorias)} categoría(s): {categorias}")
                categoria_por_catalogo[cat] = categorias

        print("\n[4/4] Scraping completado.")
        return {
            "acuerdos": acuerdos,
            "catalogo_por_acuerdo": catalogo_por_acuerdo,
            "categoria_por_catalogo": categoria_por_catalogo,
        }

    finally:
        driver.quit()
        print("Navegador cerrado.")


# ---------------------------------------------------------------------------
# Escribir en automation.py
# ---------------------------------------------------------------------------

def _repr_dict_str(d: dict) -> str:
    """Convierte un dict a código Python con sangría de 4 espacios (nivel 1)."""
    lines = ["{"]
    for k, v in d.items():
        if isinstance(v, list):
            if not v:
                lines.append(f"        {k!r}: [],")
            else:
                items = ",\n            ".join(repr(x) for x in v)
                lines.append(f"        {k!r}: [\n            {items},\n        ],")
        elif isinstance(v, dict):
            lines.append(f"        {k!r}: " + _repr_dict_str(v) + ",")
        else:
            lines.append(f"        {k!r}: {v!r},")
    lines.append("    }")
    return "\n    ".join(lines)


def _formatear_portal_defaults(data: dict) -> str:
    acuerdos_repr = "[\n        " + ",\n        ".join(repr(a) for a in data["acuerdos"]) + ",\n    ]"

    def fmt_nested(d: dict[str, list]) -> str:
        inner = []
        for k, v in d.items():
            items = ",\n            ".join(repr(x) for x in v) if v else ""
            if items:
                inner.append(f"        {k!r}: [\n            {items},\n        ]")
            else:
                inner.append(f"        {k!r}: []")
        return "{\n" + ",\n".join(inner) + "\n    }"

    bloque = textwrap.dedent(f"""\
    PORTAL_DEFAULTS: dict = {{
        "acuerdos": {acuerdos_repr},
        "catalogo_por_acuerdo": {fmt_nested(data["catalogo_por_acuerdo"])},
        "categoria_por_catalogo": {fmt_nested(data["categoria_por_catalogo"])},
    }}""")
    return bloque


def actualizar_automation_py(data: dict):
    automation_path = ROOT / "peru_compras_bot_app" / "automation.py"
    source = automation_path.read_text(encoding="utf-8")

    nuevo_bloque = _formatear_portal_defaults(data)

    # Patrón que coincide con el bloque PORTAL_DEFAULTS existente (incluyendo cierre })
    patron = re.compile(
        r"PORTAL_DEFAULTS\s*:\s*dict\s*=\s*\{.*?\n\}",
        re.DOTALL,
    )

    if not patron.search(source):
        print("\n[ERROR] No se encontró PORTAL_DEFAULTS en automation.py")
        print("Imprimiendo el bloque para que lo pegues manualmente:\n")
        print(nuevo_bloque)
        return

    nuevo_source = patron.sub(nuevo_bloque, source)

    # Backup
    backup = automation_path.with_suffix(".py.bak")
    backup.write_text(source, encoding="utf-8")
    print(f"\nBackup guardado en: {backup.name}")

    automation_path.write_text(nuevo_source, encoding="utf-8")
    print(f"PORTAL_DEFAULTS actualizado en: {automation_path.name}")
    print("\nNuevo bloque escrito:")
    print("-" * 60)
    print(nuevo_bloque)
    print("-" * 60)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("  ACTUALIZADOR DE PORTAL_DEFAULTS — Peru Compras Bot")
    print("=" * 60)
    print()
    print("Este script abrirá Chrome, harás login manualmente y luego")
    print("leerá TODOS los acuerdos, catálogos y categorías del portal.")
    print("Al terminar reescribirá PORTAL_DEFAULTS en automation.py.")
    print()
    input("Presiona ENTER para continuar... ")

    try:
        datos = scrape_defaults()
        actualizar_automation_py(datos)
        print("\n✓ Listo. Ahora reconstruye el exe con: build_exe.bat  y luego build_installer.bat")
    except Exception as e:
        print(f"\n[ERROR FATAL] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
