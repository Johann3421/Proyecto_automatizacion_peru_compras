# Peru Compras Bot

Automatiza la actualización de stock, la carga de coberturas de atención y la mejora de plazo de entrega máximo en el portal [catalogos.perucompras.gob.pe](https://www.catalogos.perucompras.gob.pe) usando Selenium, con interfaz gráfica (Tkinter), validación previa del Excel, guardado de progreso y reporte Excel con gráficos.

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

1. **Selecciona el módulo**: `Precio y existencias`, `Cobertura de atención` o `Plazo de entrega máximo`.
2. Cada módulo se abre en su propia página operativa dentro de la misma ventana.
3. **Selecciona el archivo Excel** correspondiente o descarga la plantilla del módulo actual si todavía no tienes uno.
4. La aplicación hace una **validación previa** y te indica si hay columnas faltantes, valores inválidos, filas vacías o advertencias por duplicados.
5. Ajusta los **filtros** del portal. En stock se usan Acuerdo, Catálogo y Categoría. En cobertura solo se usa Acuerdo Marco. En plazo se usan Acuerdo, Catálogo, Categoría, Región y Provincia.
6. Si necesitas dejarlo para después, usa **Guardar progreso** y luego **Continuar desde progreso**.
7. Haz clic en el botón de inicio del módulo actual cuando el panel indique que todo está listo.
8. Chrome se abrirá automáticamente. **Inicia sesión** con tu usuario y contraseña en el portal.
9. Regresa a la ventana del bot y haz clic en **"Ya inicié sesión, continuar"**.
10. El bot procesará todos los registros o el bloque seleccionado. El progreso y el estado operativo se muestran en tiempo real.
11. Al finalizar, usa **"Abrir último reporte"** para revisar el Excel generado.

### Formatos de Excel por modo

- `Precio y existencias`: columnas `Parte` y `Stock`.
- `Cobertura de atención`: columnas `Region` y `Plazo`.
- `Plazo de entrega máximo` por artículos: columnas `Parte` y `Plazo`.
- `Plazo de entrega máximo` por bloque: no requiere Excel; usa filtros y un plazo general.
- En cobertura, `Plazo` debe ser un número entero entre `1` y `90` días calendarios.
- En cobertura, el portal solo permitirá seleccionar regiones que aún no estén registradas.

### Qué cambió en la interfaz

- La ventana principal ahora funciona como un **panel guiado**, con resumen de preparación, estado operativo y acciones rápidas.
- Cada script ahora tiene su **propia página operativa** para reducir confusión durante la carga.
- El Excel se valida **antes** de ejecutar Selenium para evitar errores evitables a mitad del proceso.
- El usuario ve un estado claro de si está **listo para ejecutar**, si necesita corregir el archivo o si debe iniciar sesión en Chrome.
- Ahora existe un **modo adicional de Cobertura de atención** para cargar regiones y plazos máximos desde un Excel separado.
- Ahora existe un **modo de Plazo de entrega máximo** con trabajo por bloque o por artículos.
- Se puede **guardar el progreso** del formulario y retomarlo más tarde.

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
