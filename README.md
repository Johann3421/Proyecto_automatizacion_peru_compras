# Peru Compras Bot

Automatiza la actualización de stock en el portal [catalogos.perucompras.gob.pe](https://www.catalogos.perucompras.gob.pe) usando Selenium, con interfaz gráfica (Tkinter), validación previa del Excel y reporte Excel con gráficos.

---

## Requisitos del sistema

| Requisito | Versión mínima |
|-----------|---------------|
| Windows 10 / 11 (64-bit) | — |
| Python | 3.10 o superior |
| Google Chrome | (la versión que tengas; `webdriver-manager` descarga el driver automáticamente) |

---

## Opción A — Ejecutar desde el instalador (usuario final)

> No necesita Python ni instalar nada más.

1. Descarga `PeruComprasBot_Setup.exe` (en [Releases](../../releases)).
2. Ejecuta el instalador y sigue los pasos.
3. Abre **Peru Compras Bot** desde el escritorio o el menú Inicio.

---

## Opción B — Ejecutar desde el código fuente (desarrollador)

### 1. Clona el repositorio

```bash
git clone https://github.com/TU_USUARIO/TU_REPO.git
cd TU_REPO
```

### 2. Crea y activa un entorno virtual

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate
```

### 3. Instala las dependencias

```bash
pip install -r requirements.txt
```

### 4. Prepara el archivo de productos

Coloca un archivo Excel llamado `productos.xlsx` en la raíz del proyecto con **al menos estas dos columnas**:

| Parte | Stock |
|-------|-------|
| ABC-123 | 5 |
| XYZ-456 | 0 |

### 5. Ejecuta el bot

**Modo GUI (recomendado):**
```bash
python peru_compras_bot.py
```

**Modo consola (sin ventana gráfica):**
```bash
python peru_compras_bot.py --cli
```

---

## Opción C — Compilar el ejecutable y el instalador

### Paso 1: Compilar el `.exe`

```bash
build_exe.bat
```

Genera `dist\peru_compras_bot.exe`.

### Paso 2: Compilar el instalador de Windows

Requiere [Inno Setup 6](https://jrsoftware.org/isdl.php) instalado.

```bash
build_installer.bat
```

Genera `installer_output\PeruComprasBot_Setup.exe`.

---

## Uso de la interfaz gráfica

1. **Selecciona el archivo Excel** con los productos a actualizar o descarga la plantilla si todavía no tienes uno.
2. La aplicación hace una **validación previa** y te indica si hay columnas faltantes, stocks inválidos, filas vacías o advertencias por duplicados.
3. Ajusta los **filtros** (Acuerdo Marco, Catálogo, Categoría) según tu contrato, o importa las opciones desde el portal.
4. Haz clic en **"Iniciar actualización de stock"** solo cuando el panel indique que el archivo está listo.
5. Chrome se abrirá automáticamente. **Inicia sesión** con tu usuario y contraseña en el portal.
6. Regresa a la ventana del bot y haz clic en **"Ya inicié sesión, continuar"**.
7. El bot actualizará todos los productos. El progreso y el estado operativo se muestran en tiempo real.
8. Al finalizar, usa **"Abrir último reporte"** para revisar el Excel generado.

### Qué cambió en la interfaz

- La ventana principal ahora funciona como un **panel guiado**, con resumen de preparación, estado operativo y acciones rápidas.
- El Excel se valida **antes** de ejecutar Selenium para evitar errores evitables a mitad del proceso.
- El usuario ve un estado claro de si está **listo para ejecutar**, si necesita corregir el archivo o si debe iniciar sesión en Chrome.

---

## Reporte generado

Cada ejecución genera un archivo `reporte_YYYYMMDD_HHMMSS.xlsx` con tres hojas:

| Hoja | Contenido |
|------|-----------|
| **Resumen** | Totales de éxitos/fallos, gráfico de torta y desglose de tipos de error |
| **Detalle por Producto** | Tabla completa con estado, tipo de fallo, descripción y tiempo por producto; gráfico de tiempos |
| **Solo Fallidos** | Lista filtrada de productos fallidos con descripción del error legible |

---

## Variables de configuración

Edita las siguientes constantes en `peru_compras_bot.py` si necesitas cambiar los valores por defecto:

```python
ACUERDO_TEXTO  = "EXT-CE-2022-5 COMPUTADORAS DE ESCRITORIO"
CATALOGO_TEXTO = "COMPUTADORAS DE ESCRITORIO"
CATEGORIA_TEXTO = "MONITOR"
PAUSA_ENTRE_PRODUCTOS = 2  # segundos entre productos
```

También se pueden modificar directamente desde la interfaz gráfica sin tocar el código.

---

## Dependencias principales

```
selenium>=4.10.0
pandas>=2.0.0
openpyxl>=3.1.0
webdriver-manager>=4.0.0
pyinstaller>=5.13.0
```

---

## Estructura del proyecto

```
peru_compras_bot.py         # Entrypoint principal (GUI o CLI)
peru_compras_bot_app/
	automation.py            # Selenium, validación de Excel, reportes y ejecución CLI
	gui.py                   # Interfaz Tkinter y orquestación visual
	__init__.py
productos.xlsx             # Archivo de productos a actualizar (no rastrear en git si tiene datos sensibles)
requirements.txt           # Dependencias Python
build_exe.bat              # Compila el .exe usando la spec existente
build_installer.bat        # Compila el instalador con Inno Setup
peru_compras_bot.spec      # Configuración de PyInstaller
installer.iss              # Script de Inno Setup
.gitignore
README.md
```

---

## Notas de seguridad

- **No subas credenciales** al repositorio. El bot pide el login manualmente en cada ejecución; no almacena contraseñas.
- Si `productos.xlsx` contiene información confidencial, agrégalo a `.gitignore`.

---

## Licencia

MIT — libre para uso y modificación.
