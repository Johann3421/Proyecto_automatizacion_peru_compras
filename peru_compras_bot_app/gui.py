import json
import logging
import os
import threading
import time
import traceback
from pathlib import Path
from queue import Empty, Queue

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from peru_compras_bot_app import automation as bot


log = bot.log


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


class _Tooltip:
    """Tooltip sencillo que aparece al pasar el mouse sobre un widget."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self._tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _show(self, _event=None):
        x, y, _, cy = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 24
        y += self.widget.winfo_rooty() + cy + 20
        self._tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(
            tw, text=self.text, justify="left",
            background="#FFFBE6", foreground="#333333",
            relief="flat", borderwidth=1,
            font=("Segoe UI", 8),
            wraplength=320, padx=6, pady=4,
        )
        lbl.pack()

    def _hide(self, _event=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


class PeruComprasGUI:
    MODO_STOCK = "stock"
    MODO_COBERTURA = "cobertura"
    MODO_PLAZO = "plazo"
    PLAZO_BLOQUE = "bloque"
    PLAZO_INDIVIDUAL = "individual"

    # ── Paleta de colores ──────────────────────────────────────────────
    C_FONDO       = "#F3F6F3"
    C_SUPERFICIE  = "#FFFFFF"
    C_SUPERFICIE_2 = "#F7FAF7"
    C_HEADER      = "#1F4E43"
    C_STEP_ACTIVE = "#217346"
    C_STEP_DONE   = "#217346"
    C_STEP_IDLE   = "#94A3B8"
    C_ACCION      = "#217346"
    C_ACCION_HOVER = "#1B5E20"
    C_PELIGRO     = "#B42318"
    C_ADVERTENCIA = "#B54708"
    C_TEXTO       = "#1F2937"
    C_TEXTO_SUAVE = "#52606D"
    C_BORDE       = "#D9E2EC"
    C_LOGIN_BG    = "#FFF7E8"
    C_LOGIN_BORDE = "#F79009"
    C_OK_BG       = "#ECFDF3"
    C_OK_FG       = "#166534"
    C_INFO_BG     = "#E0F2FE"
    C_INFO_FG     = "#0C4A6E"

    def __init__(self, root):
        self.root = root
        self.root.configure(bg=self.C_FONDO)

        self.log_queue = Queue()
        self.login_event = threading.Event()
        self.worker = None
        self.reporte_generado = None
        self._pausado = False
        self._total_productos = 0
        self._procesados = 0
        self.validation_summary = None
        self._portal_snapshot = {"acuerdos": 0, "catalogos": 0, "categorias": 0}
        self._progress_file = bot.BASE_DIR / "progreso_guardado.json"

        self.operation_var = tk.StringVar(value=self.MODO_STOCK)
        self.plazo_mode_var = tk.StringVar(value=self.PLAZO_BLOQUE)
        self.excel_var    = tk.StringVar(value=str(bot.BASE_DIR / "productos.xlsx"))
        self.acuerdo_var  = tk.StringVar(value=bot.ACUERDO_TEXTO)
        self.catalogo_var = tk.StringVar(value=bot.CATALOGO_TEXTO)
        self.categoria_var = tk.StringVar(value=bot.CATEGORIA_TEXTO)
        self.region_var = tk.StringVar(value="")
        self.provincia_var = tk.StringVar(value="")
        self.plazo_general_var = tk.StringVar(value="5")
        self.pausa_var    = tk.StringVar(value=str(bot.PAUSA_ENTRE_PRODUCTOS))
        self.estado_var   = tk.StringVar(value="")
        self.readiness_var = tk.StringVar(value="Pendiente de preparación")
        self.readiness_detail_var = tk.StringVar(
            value="Selecciona un Excel y valida el contenido antes de iniciar."
        )
        self.metric_archivo_var = tk.StringVar(value="Sin revisar")
        self.metric_productos_var = tk.StringVar(value="0")
        self.metric_alertas_var = tk.StringVar(value="0")
        self.metric_portal_var = tk.StringVar(value="Sin importar")
        self.metric_reporte_var = tk.StringVar(value="Sin reporte")
        self.selection_summary_var = tk.StringVar(
            value="Aún no hay una configuración lista para ejecutar."
        )
        self.quick_status_var = tk.StringVar(value="Esperando configuración")

        self._apply_theme()
        self._build_ui()
        self._configurar_logging_gui()
        self._tick_logs()
        self._analizar_excel_actual(silencioso=True)
        self._actualizar_resumen_seleccion()

    # ------------------------------------------------------------------
    # Tema ttk personalizado
    # ------------------------------------------------------------------
    def _apply_theme(self):
        style = ttk.Style(self.root)
        # Usar 'clam' como base (compatible con Windows/Linux)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Segoe UI", 10), background=self.C_FONDO, foreground=self.C_TEXTO)
        style.configure("TFrame",      background=self.C_FONDO)
        style.configure("TLabel",      background=self.C_FONDO, foreground=self.C_TEXTO)
        style.configure("TLabelframe", background=self.C_FONDO)
        style.configure("TLabelframe.Label", background=self.C_FONDO, foreground=self.C_HEADER,
                        font=("Segoe UI Semibold", 10))
        style.configure("TEntry", fieldbackground="#FFFFFF", bordercolor=self.C_BORDE,
                        lightcolor=self.C_BORDE, darkcolor=self.C_BORDE, padding=8)
        style.configure("TCombobox", fieldbackground="#FFFFFF", bordercolor=self.C_BORDE,
                        lightcolor=self.C_BORDE, darkcolor=self.C_BORDE, padding=6)
        style.configure("Excel.TNotebook", background=self.C_FONDO, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "Excel.TNotebook.Tab",
            background="#D9EAD3",
            foreground=self.C_HEADER,
            padding=(16, 8),
            font=("Segoe UI Semibold", 10),
            borderwidth=1,
        )
        style.map(
            "Excel.TNotebook.Tab",
            background=[("selected", "#FFFFFF"), ("active", "#EAF4E6")],
            foreground=[("selected", self.C_ACCION), ("active", self.C_HEADER)],
        )

        # Botón principal
        style.configure("Accion.TButton",
                        background=self.C_ACCION, foreground="#FFFFFF",
                        font=("Segoe UI Semibold", 10),
                        padding=(14, 10), relief="flat", borderwidth=0)
        style.map("Accion.TButton",
                  background=[("active", self.C_ACCION_HOVER), ("disabled", "#C7D2DA")],
                  foreground=[("disabled", "#FFFFFF")])

        # Botón secundario
        style.configure("Secundario.TButton",
                        background="#FFFFFF", foreground=self.C_TEXTO,
                        font=("Segoe UI", 9),
                        padding=(10, 8), relief="flat", borderwidth=1)
        style.map("Secundario.TButton",
                  background=[("active", "#F0F4F8")])

        style.configure("Peligro.TButton",
                        background="#FDECEA", foreground=self.C_PELIGRO,
                        font=("Segoe UI Semibold", 9),
                        padding=(10, 8), relief="flat", borderwidth=0)
        style.map("Peligro.TButton",
                  background=[("active", "#FAD2CF"), ("disabled", "#F5F5F5")],
                  foreground=[("disabled", "#9AA0A6")])

        style.configure("Pausa.TButton",
                        background="#FFF8E1", foreground=self.C_ADVERTENCIA,
                        font=("Segoe UI Semibold", 9),
                        padding=(10, 8), relief="flat", borderwidth=0)
        style.map("Pausa.TButton",
                  background=[("active", "#FFF3CD"), ("disabled", "#F5F5F5")],
                  foreground=[("disabled", "#9AA0A6")])

        style.configure("Login.TButton",
                        background=self.C_LOGIN_BG, foreground="#33691E",
                        font=("Segoe UI Semibold", 10),
                        padding=(14, 10), relief="flat", borderwidth=2)
        style.map("Login.TButton",
                  background=[("active", "#F9FBE7"), ("disabled", "#F5F5F5")],
                  foreground=[("disabled", "#9AA0A6")])

        style.configure("Verde.Horizontal.TProgressbar",
                        troughcolor=self.C_BORDE,
                        background=self.C_STEP_DONE,
                        thickness=12)

    # ------------------------------------------------------------------
    # Construcción de la UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        header = tk.Frame(self.root, bg=self.C_HEADER, padx=24, pady=18)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        brand = tk.Frame(header, bg=self.C_HEADER)
        brand.grid(row=0, column=0, sticky="w")
        tk.Label(
            brand,
            text="PERU COMPRAS BOT",
            bg=self.C_HEADER,
            fg="#F0B429",
            font=("Bahnschrift SemiBold", 10),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            brand,
            text="Panel de stock, cobertura y plazo",
            bg=self.C_HEADER,
            fg="#FFFFFF",
            font=("Bahnschrift SemiBold", 21),
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(
            brand,
            text="Valida archivos, prepara filtros del portal y ejecuta cada módulo desde su propia página, con opción para guardar y continuar después.",
            bg=self.C_HEADER,
            fg="#C5D3E0",
            font=("Segoe UI", 10),
            anchor="w",
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        self._readiness_pill = tk.Label(
            header,
            textvariable=self.readiness_var,
            bg="#DCEBFF",
            fg="#0C4A6E",
            font=("Segoe UI Semibold", 10),
            padx=14,
            pady=8,
        )
        self._readiness_pill.grid(row=0, column=1, sticky="e")

        self._banner_frame = tk.Frame(self.root, bg=self.C_INFO_BG, height=38)
        self._banner_frame.grid(row=1, column=0, sticky="ew")
        self._banner_frame.grid_propagate(False)
        self._banner_lbl = tk.Label(
            self._banner_frame,
            textvariable=self.estado_var,
            bg=self.C_INFO_BG,
            fg=self.C_INFO_FG,
            font=("Segoe UI", 10),
            anchor="w",
            padx=20,
        )
        self._banner_lbl.pack(fill="both", expand=True)

        content_host = tk.Frame(self.root, bg=self.C_FONDO)
        content_host.grid(row=2, column=0, sticky="nsew")
        content_host.grid_columnconfigure(0, weight=1)
        content_host.grid_rowconfigure(0, weight=1)

        self._main_canvas = tk.Canvas(
            content_host,
            bg=self.C_FONDO,
            highlightthickness=0,
            bd=0,
        )
        self._main_canvas.grid(row=0, column=0, sticky="nsew")

        body_scroll = ttk.Scrollbar(content_host, orient="vertical", command=self._main_canvas.yview)
        body_scroll.grid(row=0, column=1, sticky="ns")
        self._main_canvas.configure(yscrollcommand=body_scroll.set)

        body = tk.Frame(self._main_canvas, bg=self.C_FONDO, padx=20, pady=18)
        self._main_canvas_window = self._main_canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", self._sync_main_scroll_region)
        self._main_canvas.bind("<Configure>", self._sync_main_scroll_width)
        self.root.bind_all("<MouseWheel>", self._on_main_mousewheel, add="+")

        body.grid_columnconfigure(0, weight=5)
        body.grid_columnconfigure(1, weight=3)

        hero = tk.Frame(
            body,
            bg=self.C_HEADER,
            highlightbackground="#203B62",
            highlightthickness=1,
            padx=18,
            pady=18,
        )
        hero.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        hero.grid_columnconfigure(0, weight=1)
        hero.grid_columnconfigure(1, weight=1)

        hero_left = tk.Frame(hero, bg=self.C_HEADER)
        hero_left.grid(row=0, column=0, sticky="nw")
        tk.Label(
            hero_left,
            text="Centro de operación listo para uso",
            bg=self.C_HEADER,
            fg="#FFFFFF",
            font=("Bahnschrift SemiBold", 17),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            hero_left,
            text="La aplicación separa cada flujo de trabajo, te indica qué falta corregir antes de abrir Chrome y reduce errores comunes antes de ejecutar.",
            bg=self.C_HEADER,
            fg="#C5D3E0",
            font=("Segoe UI", 10),
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(anchor="w", pady=(8, 14))

        hero_actions = tk.Frame(hero_left, bg=self.C_HEADER)
        hero_actions.pack(anchor="w")
        ttk.Button(
            hero_actions,
            text="Seleccionar Excel",
            command=self._seleccionar_excel,
            style="Accion.TButton",
        ).pack(side="left")
        ttk.Button(
            hero_actions,
            text="Abrir carpeta del proyecto",
            command=self._abrir_carpeta,
            style="Secundario.TButton",
        ).pack(side="left", padx=(10, 0))

        hero_right = tk.Frame(hero, bg=self.C_HEADER)
        hero_right.grid(row=0, column=1, sticky="ne", padx=(18, 0))
        tk.Label(
            hero_right,
            text="Resumen rápido",
            bg=self.C_HEADER,
            fg="#F8FAFC",
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).pack(anchor="w")

        stats_row = tk.Frame(hero_right, bg=self.C_HEADER)
        stats_row.pack(anchor="e", pady=(10, 0))
        self._make_metric_card(stats_row, "Archivo", self.metric_archivo_var, 0, dark=True)
        self._make_metric_card(stats_row, "Productos listos", self.metric_productos_var, 1, dark=True)
        self._make_metric_card(stats_row, "Alertas", self.metric_alertas_var, 2, dark=True)
        self._make_metric_card(stats_row, "Portal", self.metric_portal_var, 3, dark=True)

        workflow_panel, workflow_body = self._make_panel(
            body,
            "Flujo de trabajo",
            "Elige el tipo de mejora, valida el Excel y ejecuta el flujo correspondiente.",
        )
        workflow_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))

        sidebar_panel, sidebar_body = self._make_panel(
            body,
            "Control de operación",
            "Estado de preparación, guía y acciones rápidas.",
        )
        sidebar_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))

        f0 = self._make_section(workflow_body, "0", "Pestañas de trabajo", "Cambia de módulo como si cambiaras de hoja en Excel: cada pestaña conserva su configuración visual dentro del mismo panel.")

        self._module_notebook = ttk.Notebook(f0, style="Excel.TNotebook")
        self._module_notebook.pack(fill="x")
        self._module_tabs = {}
        module_defs = [
            (self.MODO_STOCK, "Precio y existencias", "Actualiza stock por parte con una hoja tipo carga masiva."),
            (self.MODO_COBERTURA, "Cobertura de atención", "Trabaja regiones y plazo máximo con una vista orientada a cobertura."),
            (self.MODO_PLAZO, "Plazo de entrega máximo", "Usa bloque o artículos como si cambiaras de formato en una planilla."),
        ]
        for mode_key, label, description in module_defs:
            tab = tk.Frame(self._module_notebook, bg="#FFFFFF", padx=14, pady=12)
            tk.Label(
                tab,
                text=label,
                bg="#FFFFFF",
                fg=self.C_ACCION,
                font=("Segoe UI Semibold", 11),
                anchor="w",
            ).pack(anchor="w")
            tk.Label(
                tab,
                text=description,
                bg="#FFFFFF",
                fg=self.C_TEXTO_SUAVE,
                font=("Segoe UI", 9),
                anchor="w",
                justify="left",
                wraplength=620,
            ).pack(anchor="w", pady=(6, 0))
            self._module_notebook.add(tab, text=label)
            self._module_tabs[mode_key] = tab
        self._module_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._mode_help_lbl = tk.Label(
            f0,
            text="Usa 'Precio y existencias' para actualizar stock por número de parte.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._mode_help_lbl.pack(fill="x", pady=(10, 0))

        f1 = self._make_section(workflow_body, "1", "Archivo de carga", "Carga tu Excel y revisa la validación antes de correr el bot cuando el módulo lo requiera.")

        self._plazo_mode_frame = tk.Frame(f1, bg=self.C_SUPERFICIE)
        self._plazo_mode_title = tk.Label(
            self._plazo_mode_frame,
            text="Forma de trabajo para plazo",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        )
        self._plazo_mode_title.pack(anchor="w")
        plazo_mode_row = tk.Frame(self._plazo_mode_frame, bg=self.C_SUPERFICIE)
        plazo_mode_row.pack(fill="x", pady=(8, 0))
        tk.Radiobutton(
            plazo_mode_row,
            text="Por bloque",
            variable=self.plazo_mode_var,
            value=self.PLAZO_BLOQUE,
            command=self._on_plazo_mode_changed,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            activebackground=self.C_SUPERFICIE,
            font=("Segoe UI", 9),
        ).pack(side="left")
        tk.Radiobutton(
            plazo_mode_row,
            text="Por artículos desde Excel",
            variable=self.plazo_mode_var,
            value=self.PLAZO_INDIVIDUAL,
            command=self._on_plazo_mode_changed,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            activebackground=self.C_SUPERFICIE,
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(18, 0))
        self._plazo_mode_hint_lbl = tk.Label(
            self._plazo_mode_frame,
            text="Por bloque aplica el mismo plazo a todos los resultados visibles. Por artículos usa Excel con columnas Parte y Plazo.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._plazo_mode_hint_lbl.pack(fill="x", pady=(6, 0))
        self._plazo_mode_frame.pack(fill="x", pady=(0, 12))

        self._plazo_bloque_frame = tk.Frame(f1, bg=self.C_SUPERFICIE)
        plazo_bloque_row = tk.Frame(self._plazo_bloque_frame, bg=self.C_SUPERFICIE)
        plazo_bloque_row.pack(fill="x")
        tk.Label(
            plazo_bloque_row,
            text="Plazo general para el bloque (días)",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI", 9),
        ).pack(side="left")
        self.entry_plazo_general = ttk.Entry(plazo_bloque_row, textvariable=self.plazo_general_var, width=10)
        self.entry_plazo_general.pack(side="left", padx=(8, 0))
        _Tooltip(self.entry_plazo_general, "Este valor se aplicará a todos los resultados visibles del bloque de plazo.")
        self._plazo_bloque_hint_lbl = tk.Label(
            self._plazo_bloque_frame,
            text="En modo por bloque no se requiere Excel. El bot buscará por acuerdo, catálogo, categoría, región y provincia.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._plazo_bloque_hint_lbl.pack(fill="x", pady=(6, 0))
        self._plazo_bloque_frame.pack(fill="x", pady=(0, 12))

        fila_excel = ttk.Frame(f1)
        fila_excel.pack(fill="x", pady=(0, 8))
        fila_excel.columnconfigure(0, weight=1)
        self._excel_row_frame = fila_excel

        self.entry_excel = ttk.Entry(fila_excel, textvariable=self.excel_var)
        self.entry_excel.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        btn_sel = ttk.Button(
            fila_excel,
            text="Buscar archivo",
            command=self._seleccionar_excel,
            style="Secundario.TButton",
        )
        btn_sel.grid(row=0, column=1)
        _Tooltip(btn_sel, "Selecciona un Excel .xlsx o .xls. La estructura depende del módulo activo.")

        actions_file = tk.Frame(f1, bg=self.C_SUPERFICIE)
        actions_file.pack(fill="x")
        ttk.Button(
            actions_file,
            text="Descargar plantilla",
            command=self._descargar_plantilla,
            style="Secundario.TButton",
        ).pack(side="left")
        ttk.Button(
            actions_file,
            text="Validar archivo ahora",
            command=self._analizar_excel_actual,
            style="Secundario.TButton",
        ).pack(side="left", padx=(8, 0))

        tk.Label(
            f1,
            text="El sistema revisa columnas, filas vacías, valores inválidos y duplicados antes de permitir la ejecución cuando el módulo usa Excel.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(10, 0))

        self._validation_box = tk.Frame(
            f1,
            bg=self.C_SUPERFICIE_2,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        self._validation_box.pack(fill="x", pady=(12, 0))
        self._validation_title_lbl = tk.Label(
            self._validation_box,
            text="Sin análisis todavía",
            bg=self.C_SUPERFICIE_2,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 11),
            anchor="w",
        )
        self._validation_title_lbl.pack(anchor="w")
        self._validation_detail_lbl = tk.Label(
            self._validation_box,
            text="Selecciona un archivo para revisar qué tan listo está antes de ejecutar.",
            bg=self.C_SUPERFICIE_2,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._validation_detail_lbl.pack(anchor="w", pady=(6, 0))
        self._validation_examples_lbl = tk.Label(
            self._validation_box,
            text="",
            bg=self.C_SUPERFICIE_2,
            fg=self.C_ADVERTENCIA,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._validation_examples_lbl.pack(anchor="w", pady=(6, 0))

        f2 = self._make_section(workflow_body, "2", "Filtros del portal", "Trae las opciones desde Peru Compras o ajusta las predeterminadas según el modo seleccionado.")

        aviso_f = tk.Frame(
            f2,
            bg="#EFF6FF",
            highlightbackground="#BFDBFE",
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        aviso_f.pack(fill="x", pady=(0, 12))
        tk.Label(
            aviso_f,
            text="Importa las listas del portal si necesitas que los filtros coincidan exactamente con tu contrato actual.",
            bg="#EFF6FF",
            fg="#1D4ED8",
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=460,
        ).pack(side="left", fill="x", expand=True)
        self.btn_cargar_opts = ttk.Button(
            aviso_f,
            text="Importar opciones del portal",
            command=self._cargar_opciones,
            style="Secundario.TButton",
        )
        self.btn_cargar_opts.pack(side="right", padx=(12, 0))

        self._filter_mode_note_lbl = tk.Label(
            f2,
            text="En este modo se usan Acuerdo, Catálogo y Categoría.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        )
        self._filter_mode_note_lbl.pack(fill="x", pady=(0, 10))

        grid_f = tk.Frame(f2, bg=self.C_SUPERFICIE)
        grid_f.pack(fill="x")
        grid_f.columnconfigure(1, weight=1)

        labels_filtros = ["Acuerdo Marco", "Catálogo", "Categoría", "Región", "Provincia"]
        tips_filtros = [
            "Selecciona el Acuerdo Marco correcto para los productos que vas a actualizar.",
            "Selecciona el Catálogo Electrónico correspondiente.",
            "Selecciona la categoría específica dentro del catálogo.",
            "Selecciona la región que se usará en cobertura o plazo.",
            "Selecciona la provincia específica para el bloque de plazo.",
        ]
        self.combo_acuerdo = self._make_combo_row(grid_f, 0, labels_filtros[0], self.acuerdo_var, tips_filtros[0])
        self.combo_catalogo = self._make_combo_row(grid_f, 1, labels_filtros[1], self.catalogo_var, tips_filtros[1])
        self.combo_categoria = self._make_combo_row(grid_f, 2, labels_filtros[2], self.categoria_var, tips_filtros[2])
        self.combo_region = self._make_combo_row(grid_f, 3, labels_filtros[3], self.region_var, tips_filtros[3])
        self.combo_provincia = self._make_combo_row(grid_f, 4, labels_filtros[4], self.provincia_var, tips_filtros[4])
        self.combo_acuerdo.bind("<<ComboboxSelected>>", self._on_acuerdo_changed)
        self.combo_catalogo.bind("<<ComboboxSelected>>", self._on_catalogo_changed)
        self.combo_region.bind("<<ComboboxSelected>>", self._on_region_changed)

        self._avanzado_visible = tk.BooleanVar(value=False)
        self._btn_avanzado = ttk.Button(
            f2,
            text="Ver configuración avanzada",
            command=self._toggle_avanzado,
            style="Secundario.TButton",
        )
        self._btn_avanzado.pack(anchor="w", pady=(10, 0))

        self._frame_avanzado = tk.Frame(f2, bg=self.C_SUPERFICIE)
        av_lbl = ttk.Label(
            self._frame_avanzado,
            text="Pausa entre productos (segundos)",
            foreground=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
        )
        av_lbl.pack(side="left", pady=(4, 0))
        av_entry = ttk.Entry(self._frame_avanzado, textvariable=self.pausa_var, width=8)
        av_entry.pack(side="left", padx=8, pady=(4, 0))
        _Tooltip(av_entry, "Auméntalo si el portal responde lento o bloquea acciones seguidas.")

        f3 = self._make_section(workflow_body, "3", "Ejecución", "La app solo arrancará cuando el Excel esté listo y los filtros estén completos.")

        self._execution_state_lbl = tk.Label(
            f3,
            textvariable=self.quick_status_var,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
        )
        self._execution_state_lbl.pack(fill="x", pady=(0, 8))

        self.btn_iniciar = ttk.Button(
            f3,
            text="Iniciar actualización de stock",
            command=self._iniciar,
            style="Accion.TButton",
        )
        self.btn_iniciar.pack(fill="x", pady=(0, 12))

        self._panel_login = tk.Frame(
            f3,
            bg=self.C_LOGIN_BG,
            highlightbackground=self.C_LOGIN_BORDE,
            highlightthickness=2,
            padx=14,
            pady=12,
        )
        tk.Label(
            self._panel_login,
            text="Chrome está esperando que inicies sesión",
            bg=self.C_LOGIN_BG,
            fg="#7B5800",
            font=("Segoe UI Semibold", 11),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            self._panel_login,
            text="1. Ve a la ventana de Chrome que se abrió.\n2. Inicia sesión en Peru Compras.\n3. Vuelve aquí y confirma para continuar.",
            bg=self.C_LOGIN_BG,
            fg="#5D4037",
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
        ).pack(anchor="w", pady=(6, 10))
        self.btn_login = ttk.Button(
            self._panel_login,
            text="Ya inicié sesión, continuar",
            command=self._continuar_login,
            style="Login.TButton",
        )
        self.btn_login.pack(fill="x")

        self._panel_ctrl = tk.Frame(f3, bg=self.C_SUPERFICIE)
        prog_frame = tk.Frame(self._panel_ctrl, bg=self.C_SUPERFICIE)
        prog_frame.pack(fill="x", pady=(0, 8))
        self._lbl_progreso = tk.Label(
            prog_frame,
            text="Preparando ejecución...",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
        )
        self._lbl_progreso.pack(anchor="w")
        self.progress = ttk.Progressbar(
            prog_frame,
            orient="horizontal",
            mode="determinate",
            style="Verde.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(4, 0))

        ctrl_btns = tk.Frame(self._panel_ctrl, bg=self.C_SUPERFICIE)
        ctrl_btns.pack(fill="x")
        self.btn_pausar = ttk.Button(
            ctrl_btns,
            text="Pausar",
            command=self._pausar_reanudar,
            style="Pausa.TButton",
        )
        self.btn_pausar.pack(side="left", padx=(0, 8))
        self.btn_detener = ttk.Button(
            ctrl_btns,
            text="Detener y generar reporte",
            command=self._detener,
            style="Peligro.TButton",
        )
        self.btn_detener.pack(side="left")

        self._panel_resultado = tk.Frame(
            f3,
            bg=self.C_OK_BG,
            highlightbackground="#86EFAC",
            highlightthickness=2,
            padx=14,
            pady=12,
        )
        tk.Label(
            self._panel_resultado,
            text="Proceso completado",
            bg=self.C_OK_BG,
            fg=self.C_OK_FG,
            font=("Segoe UI Semibold", 11),
            anchor="w",
        ).pack(anchor="w")
        self._lbl_resultado_info = tk.Label(
            self._panel_resultado,
            text="",
            bg=self.C_OK_BG,
            fg=self.C_OK_FG,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
        )
        self._lbl_resultado_info.pack(anchor="w", pady=(6, 8))
        res_btns = tk.Frame(self._panel_resultado, bg=self.C_OK_BG)
        res_btns.pack(anchor="w")
        self.btn_abrir_reporte = ttk.Button(
            res_btns,
            text="Abrir reporte Excel",
            command=self._abrir_reporte,
            style="Accion.TButton",
        )
        self.btn_abrir_reporte.pack(side="left", padx=(0, 10))
        ttk.Button(
            res_btns,
            text="Abrir carpeta",
            command=self._abrir_carpeta,
            style="Secundario.TButton",
        ).pack(side="left")

        readiness_card = tk.Frame(
            sidebar_body,
            bg=self.C_SUPERFICIE_2,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        readiness_card.pack(fill="x")
        self._readiness_card = readiness_card
        tk.Label(
            readiness_card,
            text="Estado de preparación",
            bg=self.C_SUPERFICIE_2,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(anchor="w")
        self._readiness_title_lbl = tk.Label(
            readiness_card,
            textvariable=self.readiness_var,
            bg=self.C_SUPERFICIE_2,
            fg=self.C_TEXTO,
            font=("Bahnschrift SemiBold", 16),
            anchor="w",
        )
        self._readiness_title_lbl.pack(anchor="w", pady=(6, 4))
        self._readiness_detail_lbl = tk.Label(
            readiness_card,
            textvariable=self.readiness_detail_var,
            bg=self.C_SUPERFICIE_2,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=320,
        )
        self._readiness_detail_lbl.pack(anchor="w")

        selection_card = tk.Frame(
            sidebar_body,
            bg=self.C_SUPERFICIE,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        selection_card.pack(fill="x", pady=(12, 0))
        tk.Label(
            selection_card,
            text="Configuración actual",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            selection_card,
            textvariable=self.selection_summary_var,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=320,
        ).pack(anchor="w", pady=(6, 0))

        ops_card = tk.Frame(
            sidebar_body,
            bg=self.C_SUPERFICIE,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        ops_card.pack(fill="x", pady=(12, 0))
        tk.Label(
            ops_card,
            text="Acciones rápidas",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).pack(anchor="w")
        ops_metrics = tk.Frame(ops_card, bg=self.C_SUPERFICIE)
        ops_metrics.pack(fill="x", pady=(10, 2))
        self._make_metric_card(ops_metrics, "Último reporte", self.metric_reporte_var, 0)
        ttk.Button(
            ops_card,
            text="Abrir último reporte",
            command=self._abrir_reporte,
            style="Secundario.TButton",
        ).pack(fill="x", pady=(10, 6))
        ttk.Button(
            ops_card,
            text="Guardar progreso",
            command=self._guardar_progreso,
            style="Secundario.TButton",
        ).pack(fill="x", pady=(0, 6))
        ttk.Button(
            ops_card,
            text="Continuar desde progreso",
            command=self._cargar_progreso,
            style="Secundario.TButton",
        ).pack(fill="x", pady=(0, 6))
        ttk.Button(
            ops_card,
            text="Estadísticas de errores",
            command=self._ver_aprendizaje,
            style="Secundario.TButton",
        ).pack(fill="x")

        guide_card = tk.Frame(
            sidebar_body,
            bg=self.C_SUPERFICIE,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        guide_card.pack(fill="x", pady=(12, 0))
        tk.Label(
            guide_card,
            text="Guía rápida",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            guide_card,
            text="1. Elige el módulo correcto.\n2. Carga Excel si ese flujo lo necesita o define el plazo general.\n3. Revisa o importa los filtros del portal.\n4. Guarda progreso si necesitas continuar luego.\n5. Inicia sesión en Chrome cuando se solicite y revisa el reporte al finalizar.",
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        log_panel, log_body = self._make_panel(
            body,
            "Actividad y diagnóstico",
            "El log técnico queda disponible para soporte, pero no interfiere con el flujo principal.",
        )
        log_panel.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(16, 0))

        log_header = tk.Frame(log_body, bg=self.C_SUPERFICIE)
        log_header.pack(fill="x", pady=(0, 8))
        self._log_visible = tk.BooleanVar(value=False)
        self._btn_toggle_log = ttk.Button(
            log_header,
            text="Mostrar actividad detallada",
            command=self._toggle_log,
            style="Secundario.TButton",
        )
        self._btn_toggle_log.pack(side="left")

        self._frame_log = ttk.LabelFrame(log_body, text="Actividad en tiempo real", padding=8)
        self.txt_log = scrolledtext.ScrolledText(
            self._frame_log,
            height=16,
            wrap="word",
            state="disabled",
            font=("Consolas", 8),
            background="#0F172A",
            foreground="#DDE7F0",
            insertbackground="#FFFFFF",
        )
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.tag_configure("error", foreground="#FCA5A5")
        self.txt_log.tag_configure("warning", foreground="#FCD34D")
        self.txt_log.tag_configure("ok", foreground="#86EFAC")
        self.txt_log.tag_configure("paso", foreground="#93C5FD", font=("Consolas", 8, "bold"))

        tk.Label(
            body,
            text="Peru Compras Bot  |  interfaz guiada para operación interna",
            bg=self.C_FONDO,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 8),
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self._sync_main_scroll_region()
        self._aplicar_modo_operacion_ui()

    # ------------------------------------------------------------------
    # Helpers de construcción
    # ------------------------------------------------------------------
    def _make_panel(self, parent, title: str, subtitle: str = ""):
        panel = tk.Frame(
            parent,
            bg=self.C_SUPERFICIE,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=16,
            pady=16,
        )
        tk.Label(
            panel,
            text=title,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Bahnschrift SemiBold", 15),
            anchor="w",
        ).pack(anchor="w")
        if subtitle:
            tk.Label(
                panel,
                text=subtitle,
                bg=self.C_SUPERFICIE,
                fg=self.C_TEXTO_SUAVE,
                font=("Segoe UI", 9),
                anchor="w",
                justify="left",
                wraplength=760,
            ).pack(anchor="w", pady=(4, 12))
        body = tk.Frame(panel, bg=self.C_SUPERFICIE)
        body.pack(fill="both", expand=True)
        return panel, body

    def _sync_main_scroll_region(self, _event=None):
        if hasattr(self, "_main_canvas"):
            self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all"))

    def _sync_main_scroll_width(self, event):
        if hasattr(self, "_main_canvas_window"):
            self._main_canvas.itemconfigure(self._main_canvas_window, width=event.width)

    def _on_main_mousewheel(self, event):
        if not hasattr(self, "_main_canvas"):
            return

        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is None or not self._widget_is_inside(widget, self._main_canvas):
            return
        if hasattr(self, "txt_log") and self._widget_is_inside(widget, self.txt_log):
            return

        region = self._main_canvas.bbox("all")
        if not region or region[3] <= self._main_canvas.winfo_height():
            return

        delta = int(-event.delta / 120) if event.delta else 0
        if delta:
            self._main_canvas.yview_scroll(delta, "units")

    @staticmethod
    def _widget_is_inside(widget, ancestor):
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            current = current.nametowidget(parent_name)
        return False

    def _make_section(self, parent, step: str, title: str, subtitle: str) -> tk.Frame:
        section = tk.Frame(parent, bg=self.C_SUPERFICIE)
        section.pack(fill="x", pady=(0, 18))
        head = tk.Frame(section, bg=self.C_SUPERFICIE)
        head.pack(fill="x")
        tk.Label(
            head,
            text=step,
            bg=self.C_STEP_ACTIVE,
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 10),
            width=3,
            pady=4,
        ).pack(side="left")
        text_wrap = tk.Frame(head, bg=self.C_SUPERFICIE)
        text_wrap.pack(side="left", fill="x", expand=True, padx=(10, 0))
        tk.Label(
            text_wrap,
            text=title,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO,
            font=("Segoe UI Semibold", 11),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            text_wrap,
            text=subtitle,
            bg=self.C_SUPERFICIE,
            fg=self.C_TEXTO_SUAVE,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=620,
        ).pack(anchor="w", pady=(2, 0))
        content = tk.Frame(
            section,
            bg=self.C_SUPERFICIE,
            highlightbackground=self.C_BORDE,
            highlightthickness=1,
            padx=14,
            pady=14,
        )
        content.pack(fill="x", pady=(10, 0))
        return content

    def _make_metric_card(self, parent, title: str, variable: tk.StringVar, column: int, dark: bool = False):
        bg = "#17304F" if dark else self.C_SUPERFICIE_2
        fg_title = "#B6C6D8" if dark else self.C_TEXTO_SUAVE
        fg_value = "#FFFFFF" if dark else self.C_TEXTO
        card = tk.Frame(
            parent,
            bg=bg,
            highlightbackground="#27486E" if dark else self.C_BORDE,
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        card.grid(row=0, column=column, padx=(0, 8), sticky="nsew")
        tk.Label(
            card,
            text=title,
            bg=bg,
            fg=fg_title,
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card,
            textvariable=variable,
            bg=bg,
            fg=fg_value,
            font=("Bahnschrift SemiBold", 15),
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

    def _aplicar_estado_preparacion(self, titulo: str, detalle: str, tono: str = "info"):
        tonos = {
            "info": {"bg": "#DCEBFF", "fg": "#0C4A6E", "card": self.C_SUPERFICIE_2},
            "ok": {"bg": "#DCFCE7", "fg": self.C_OK_FG, "card": "#F0FDF4"},
            "warning": {"bg": "#FEF3C7", "fg": self.C_ADVERTENCIA, "card": "#FFF7E8"},
            "error": {"bg": "#FEE4E2", "fg": self.C_PELIGRO, "card": "#FFF5F4"},
        }
        estilo = tonos.get(tono, tonos["info"])
        self.readiness_var.set(titulo)
        self.readiness_detail_var.set(detalle)
        self._readiness_pill.configure(bg=estilo["bg"], fg=estilo["fg"])
        self._readiness_card.configure(bg=estilo["card"])
        self._readiness_title_lbl.configure(bg=estilo["card"], fg=estilo["fg"])
        self._readiness_detail_lbl.configure(bg=estilo["card"])

    def _es_modo_plazo(self) -> bool:
        return self.operation_var.get() == self.MODO_PLAZO

    def _es_modo_cobertura(self) -> bool:
        return self.operation_var.get() == self.MODO_COBERTURA

    def _es_plazo_individual(self) -> bool:
        return self._es_modo_plazo() and self.plazo_mode_var.get() == self.PLAZO_INDIVIDUAL

    def _requiere_excel(self) -> bool:
        return not self._es_modo_plazo() or self._es_plazo_individual()

    def _change_operation(self, modo: str):
        self.operation_var.set(modo)
        if hasattr(self, "_module_notebook") and modo in getattr(self, "_module_tabs", {}):
            tab_id = str(self._module_tabs[modo])
            if self._module_notebook.select() != tab_id:
                self._module_notebook.select(self._module_tabs[modo])
        self._aplicar_modo_operacion_ui()
        self._analizar_excel_actual(silencioso=True)
        self._actualizar_resumen_seleccion()

    def _on_tab_changed(self, _event=None):
        selected = self._module_notebook.select()
        for modo, tab in self._module_tabs.items():
            if str(tab) == selected and self.operation_var.get() != modo:
                self.operation_var.set(modo)
                self._aplicar_modo_operacion_ui()
                self._analizar_excel_actual(silencioso=True)
                self._actualizar_resumen_seleccion()
                break

    def _texto_operacion(self) -> str:
        if self._es_modo_cobertura():
            return "cobertura de atención"
        if self._es_modo_plazo():
            return "plazo de entrega máximo"
        return "precio y existencias"

    def _mostrar_combo(self, combo: ttk.Combobox, visible: bool):
        if visible:
            combo._label_widget.grid()
            combo.grid()
        else:
            combo._label_widget.grid_remove()
            combo.grid_remove()

    def _on_plazo_mode_changed(self):
        self._aplicar_modo_operacion_ui()
        self._analizar_excel_actual(silencioso=True)
        self._actualizar_resumen_seleccion()

    def _aplicar_modo_operacion_ui(self):
        self._plazo_mode_frame.pack_forget()
        self._plazo_bloque_frame.pack_forget()
        self._mostrar_combo(self.combo_catalogo, not self._es_modo_cobertura())
        self._mostrar_combo(self.combo_categoria, not self._es_modo_cobertura())
        self._mostrar_combo(self.combo_region, self._es_modo_plazo())
        self._mostrar_combo(self.combo_provincia, self._es_modo_plazo())

        if self._es_modo_cobertura():
            self._mode_help_lbl.configure(
                text="Usa 'Cobertura de atención' para agregar regiones y su plazo máximo de entrega por acuerdo marco."
            )
            self._filter_mode_note_lbl.configure(
                text="En cobertura solo se utiliza el Acuerdo Marco. Catálogo y Categoría se ignoran en la ejecución."
            )
            self.btn_iniciar.configure(text="Iniciar actualización de cobertura")
            self.quick_status_var.set("Modo cobertura seleccionado")
            self._set_banner(
                "Modo cobertura activo: el Excel debe tener columnas de región y plazo.",
                self.C_INFO_BG,
                self.C_INFO_FG,
            )
        elif self._es_modo_plazo():
            self._plazo_mode_frame.pack(fill="x", pady=(0, 12), before=self._excel_row_frame)
            if self._es_plazo_individual():
                self._mode_help_lbl.configure(
                    text="Usa 'Plazo de entrega máximo' por artículos para buscar por número de parte y aplicar el plazo de cada fila del Excel."
                )
                self._filter_mode_note_lbl.configure(
                    text="En plazo por artículos se usan Acuerdo, Catálogo, Categoría, Región y Provincia, además del Excel con Parte y Plazo."
                )
                self.btn_iniciar.configure(text="Iniciar carga individual de plazo")
                self.metric_archivo_var.set(Path(self.excel_var.get().strip()).name if self.excel_var.get().strip() else "Requerido")
            else:
                self._mode_help_lbl.configure(
                    text="Usa 'Plazo de entrega máximo' por bloque para aplicar un mismo plazo a todos los resultados visibles del filtro seleccionado."
                )
                self._filter_mode_note_lbl.configure(
                    text="En plazo por bloque se usan Acuerdo, Catálogo, Categoría, Región y Provincia. No se necesita Excel."
                )
                self._plazo_bloque_frame.pack(fill="x", pady=(0, 12), before=self._excel_row_frame)
                self.btn_iniciar.configure(text="Iniciar actualización de plazo por bloque")
            self.quick_status_var.set("Modo plazo seleccionado")
            self._set_banner(
                "Modo plazo activo: elige trabajo por bloque o por artículos antes de ejecutar.",
                self.C_INFO_BG,
                self.C_INFO_FG,
            )
        else:
            self._mode_help_lbl.configure(
                text="Usa 'Precio y existencias' para actualizar stock por número de parte."
            )
            self._filter_mode_note_lbl.configure(
                text="En este modo se usan Acuerdo, Catálogo y Categoría."
            )
            self.btn_iniciar.configure(text="Iniciar actualización de stock")
            self.quick_status_var.set("Modo precio y existencias seleccionado")
            self._set_banner(
                "Modo precio y existencias activo: el Excel debe tener columnas Parte y Stock.",
                self.C_INFO_BG,
                self.C_INFO_FG,
            )

    def _actualizar_resumen_seleccion(self):
        archivo = Path(self.excel_var.get().strip()).name if self.excel_var.get().strip() else "sin archivo"
        acuerdo = self.acuerdo_var.get().strip() or "pendiente"
        catalogo = self.catalogo_var.get().strip() or "pendiente"
        categoria = self.categoria_var.get().strip() or "pendiente"
        region = self.region_var.get().strip() or "pendiente"
        provincia = self.provincia_var.get().strip() or "pendiente"
        if self._es_modo_cobertura():
            self.selection_summary_var.set(
                f"Modo: Cobertura de atención\nArchivo: {archivo}\nAcuerdo: {acuerdo}\nPlazo por archivo: sí\nPausa: {self.pausa_var.get().strip() or '2'} s"
            )
        elif self._es_modo_plazo():
            if self._es_plazo_individual():
                self.selection_summary_var.set(
                    f"Modo: Plazo de entrega máximo\nSubmodo: Por artículos\nArchivo: {archivo}\nAcuerdo: {acuerdo}\nCatálogo: {catalogo}\nCategoría: {categoria}\nRegión: {region}\nProvincia: {provincia}\nPausa: {self.pausa_var.get().strip() or '2'} s"
                )
            else:
                self.selection_summary_var.set(
                    f"Modo: Plazo de entrega máximo\nSubmodo: Por bloque\nArchivo: no requerido\nAcuerdo: {acuerdo}\nCatálogo: {catalogo}\nCategoría: {categoria}\nRegión: {region}\nProvincia: {provincia}\nPlazo: {self.plazo_general_var.get().strip() or 'pendiente'} día(s)\nPausa: {self.pausa_var.get().strip() or '2'} s"
                )
        else:
            self.selection_summary_var.set(
                f"Modo: Precio y existencias\nArchivo: {archivo}\nAcuerdo: {acuerdo}\nCatálogo: {catalogo}\nCategoría: {categoria}\nPausa: {self.pausa_var.get().strip() or '2'} s"
            )

    def _actualizar_resumen_excel_ui(self, resumen):
        if self._es_modo_plazo() and not self._es_plazo_individual():
            self.metric_archivo_var.set("No requerido")
            self.metric_productos_var.set("Bloque")
            self.metric_alertas_var.set("0")
            self._validation_box.configure(bg="#F0F9FF", highlightbackground="#7DD3FC")
            self._validation_title_lbl.configure(bg="#F0F9FF", fg=self.C_INFO_FG, text="Modo por bloque listo")
            self._validation_detail_lbl.configure(
                bg="#F0F9FF",
                fg=self.C_INFO_FG,
                text="Este flujo no usa Excel. Solo debes completar los filtros y el plazo general antes de iniciar.",
            )
            self._validation_examples_lbl.configure(bg="#F0F9FF", fg=self.C_TEXTO_SUAVE, text="")
            self.quick_status_var.set("Completa filtros de plazo por bloque")
            self._aplicar_estado_preparacion(
                "Listo para preparar bloque",
                "El módulo por bloque no requiere archivo. Verifica acuerdo, catálogo, categoría, región, provincia y plazo general.",
                "info",
            )
            return

        if resumen is None:
            self.metric_archivo_var.set("Sin revisar")
            self.metric_productos_var.set("0")
            self.metric_alertas_var.set("0")
            self.quick_status_var.set("Selecciona un archivo para empezar")
            self._aplicar_estado_preparacion(
                "Pendiente de preparación",
                "Selecciona un Excel y revisa el resumen previo antes de iniciar.",
                "info",
            )
            return

        self.metric_archivo_var.set(resumen.file_path.name)
        self.metric_productos_var.set(str(resumen.valid_rows))
        self.metric_alertas_var.set(str(resumen.total_problem_rows + len(resumen.warnings)))
        entidad = "región(es)" if self._es_modo_cobertura() else "producto(s)"

        if resumen.is_ready:
            detalle = f"{resumen.valid_rows} {entidad} listo(s) para procesar."
            if resumen.warnings:
                detalle += " Hay advertencias no bloqueantes para revisar."
            self._validation_box.configure(bg="#F0FDF4", highlightbackground="#86EFAC")
            self._validation_title_lbl.configure(bg="#F0FDF4", fg=self.C_OK_FG, text="Archivo validado correctamente")
            self._validation_detail_lbl.configure(
                bg="#F0FDF4",
                fg=self.C_OK_FG,
                text=detalle,
            )
            self._validation_examples_lbl.configure(
                bg="#F0FDF4",
                fg=self.C_ADVERTENCIA if resumen.warnings else self.C_OK_FG,
                text="\n".join(resumen.warnings[:3]),
            )
            self.quick_status_var.set("Listo para iniciar la automatización")
            self._aplicar_estado_preparacion(
                "Listo para ejecutar",
                "El Excel pasó la validación. Solo confirma los filtros del portal y arranca el proceso.",
                "ok",
            )
        else:
            self._validation_box.configure(bg="#FFF7E8", highlightbackground="#FCD34D")
            self._validation_title_lbl.configure(bg="#FFF7E8", fg=self.C_ADVERTENCIA, text="El archivo necesita correcciones")
            self._validation_detail_lbl.configure(
                bg="#FFF7E8",
                fg=self.C_TEXTO,
                text="\n".join(resumen.blocking_issues) or "No hay productos válidos para ejecutar.",
            )
            self._validation_examples_lbl.configure(
                bg="#FFF7E8",
                fg=self.C_ADVERTENCIA,
                text="\n".join(resumen.issue_examples[:4]),
            )
            self.quick_status_var.set("Corrige el Excel antes de iniciar")
            self._aplicar_estado_preparacion(
                "Requiere correcciones",
                "Hay errores en el Excel que impedirían una ejecución estable. Corrígelos antes de abrir Chrome.",
                "warning",
            )

    def _make_combo_row(self, parent, row: int, label: str, variable: tk.StringVar, tip: str) -> ttk.Combobox:
        lbl = tk.Label(parent, text=label, bg=self.C_SUPERFICIE,
                 font=("Segoe UI", 9), fg=self.C_TEXTO)
        lbl.grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=4,
        )
        combo = ttk.Combobox(parent, textvariable=variable, state="normal")
        combo.grid(row=row, column=1, sticky="we", pady=4)
        combo._row_index = row
        combo._label_widget = lbl
        _Tooltip(combo, tip)
        return combo

    def _toggle_avanzado(self):
        if self._avanzado_visible.get():
            self._frame_avanzado.pack_forget()
            self._btn_avanzado.configure(text="Ver configuración avanzada")
            self._avanzado_visible.set(False)
        else:
            self._frame_avanzado.pack(anchor="w", pady=(4, 0))
            self._btn_avanzado.configure(text="Ocultar configuración avanzada")
            self._avanzado_visible.set(True)

    def _toggle_log(self):
        if self._log_visible.get():
            self._frame_log.pack_forget()
            self._btn_toggle_log.configure(text="Mostrar actividad detallada")
            self._log_visible.set(False)
        else:
            self._frame_log.pack(fill="both", expand=True, pady=(0, 10))
            self._btn_toggle_log.configure(text="Ocultar actividad detallada")
            self._log_visible.set(True)

    # ------------------------------------------------------------------
    # Actualización de estado visual
    # ------------------------------------------------------------------
    def _set_banner(self, msg: str, color_bg: str = "#E8F0FE", color_fg: str = None):
        self._banner_frame.configure(bg=color_bg)
        self._banner_lbl.configure(bg=color_bg, fg=color_fg or self.C_STEP_ACTIVE)
        self.estado_var.set(f"  {msg}")

    def _mostrar_panel_login(self, mostrar: bool):
        if mostrar:
            self._panel_login.pack(fill="x", pady=(0, 8))
            self.quick_status_var.set("Esperando confirmación de login en Chrome")
            self._aplicar_estado_preparacion(
                "Acción requerida",
                "La aplicación abrió Chrome. Inicia sesión y vuelve para continuar.",
                "warning",
            )
        else:
            self._panel_login.pack_forget()

    def _mostrar_panel_ctrl(self, mostrar: bool):
        if mostrar:
            self._panel_ctrl.pack(fill="x")
        else:
            self._panel_ctrl.pack_forget()

    def _mostrar_panel_resultado(self, mostrar: bool):
        if mostrar:
            self._panel_resultado.pack(fill="x", pady=(12, 0))
        else:
            self._panel_resultado.pack_forget()

    def _actualizar_progreso(self, procesados: int, total: int, estado_txt: str = ""):
        if total > 0:
            self.progress["maximum"] = total
            self.progress["value"] = procesados
            pct = int(procesados / total * 100)
            self._lbl_progreso.configure(
                text=f"Producto {procesados} de {total}  ({pct}%)  {estado_txt}"
            )
            self.quick_status_var.set(f"Procesando {procesados} de {total} productos")
        else:
            self.progress["value"] = 0
            self._lbl_progreso.configure(text=estado_txt or "Preparando...")
            self.quick_status_var.set(estado_txt or "Preparando ejecución")

    def _analizar_excel_actual(self, silencioso: bool = False):
        ruta = self.excel_var.get().strip()
        if self._es_modo_plazo() and not self._es_plazo_individual():
            self.validation_summary = None
            self._actualizar_resumen_excel_ui(None)
            self._actualizar_resumen_seleccion()
            return None

        if not ruta:
            self.validation_summary = None
            self._actualizar_resumen_excel_ui(None)
            return None

        if self._es_modo_cobertura():
            resumen, _ = bot.analizar_excel_coberturas(Path(ruta))
        elif self._es_modo_plazo():
            resumen, _ = bot.analizar_excel_plazos(Path(ruta))
        else:
            resumen, _ = bot.analizar_excel_productos(Path(ruta))
        self.validation_summary = resumen
        self._actualizar_resumen_excel_ui(resumen)
        self._actualizar_resumen_seleccion()

        if not silencioso:
            if resumen.is_ready:
                entidad = "región(es)" if self._es_modo_cobertura() else "producto(s)"
                mensaje = f"Archivo listo. {resumen.valid_rows} {entidad} podrán procesarse."
                if resumen.warnings:
                    mensaje += "\n\nAdvertencias:\n- " + "\n- ".join(resumen.warnings[:3])
                messagebox.showinfo("Validación completada", mensaje)
            else:
                mensaje = "Corrige el Excel antes de continuar:\n\n- " + "\n- ".join(resumen.blocking_issues)
                if resumen.issue_examples:
                    mensaje += "\n\nEjemplos detectados:\n- " + "\n- ".join(resumen.issue_examples[:4])
                messagebox.showwarning("Se encontraron problemas", mensaje)

        return resumen

    # ------------------------------------------------------------------
    # Logging con color
    # ------------------------------------------------------------------
    def _configurar_logging_gui(self):
        self.gui_handler = TextQueueLogHandler(self.log_queue)
        self.gui_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        log.addHandler(self.gui_handler)

    def _tick_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.txt_log.configure(state="normal")
                msg_lower = msg.lower()
                if "[error]" in msg_lower:
                    tag = "error"
                elif "[warning]" in msg_lower:
                    tag = "warning"
                elif "[ok]" in msg_lower or "exito" in msg_lower or "exitosamente" in msg_lower:
                    tag = "ok"
                elif "paso" in msg_lower or "=" * 10 in msg:
                    tag = "paso"
                else:
                    tag = ""
                self.txt_log.insert("end", msg + "\n", tag)
                self.txt_log.see("end")
                self.txt_log.configure(state="disabled")

                # Extraer progreso del mensaje de log
                # Formato esperado: "--- Producto X/Y: ..."
                import re as _re
                m = _re.search(r"Producto (\d+)/(\d+)", msg)
                if m:
                    proc, tot = int(m.group(1)), int(m.group(2))
                    self._procesados = proc
                    self._total_productos = tot
                    self.root.after(0, lambda p=proc, t=tot: self._actualizar_progreso(p, t))

                if "PASO" in msg:
                    txt = msg.split("] ", 1)[-1]
                    self._set_banner(txt)

                # Detectar login completado (para ocultar panel login)
                if "login confirmado" in msg_lower or "login completado" in msg_lower:
                    self.root.after(0, lambda: self._mostrar_panel_login(False))

        except Empty:
            pass
        self.root.after(150, self._tick_logs)

    def _seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls")],
            initialdir=str(bot.BASE_DIR),
        )
        if ruta:
            self.excel_var.set(ruta)
            self._set_banner(f"Archivo seleccionado: {Path(ruta).name}", self.C_INFO_BG, self.C_INFO_FG)
            self._analizar_excel_actual(silencioso=True)

    def _descargar_plantilla(self):
        if self._es_modo_plazo() and not self._es_plazo_individual():
            messagebox.showinfo(
                "Plantilla no necesaria",
                "El modo de plazo por bloque no usa Excel. Solo debes completar los filtros y el plazo general.",
            )
            return

        destino = filedialog.asksaveasfilename(
            title="Guardar plantilla como…",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=(
                "plantilla_coberturas.xlsx" if self._es_modo_cobertura()
                else "plantilla_plazos.xlsx" if self._es_modo_plazo()
                else "plantilla_productos.xlsx"
            ),
            initialdir=str(bot.BASE_DIR),
        )
        if not destino:
            return
        try:
            if self._es_modo_cobertura():
                bot.generar_plantilla_cobertura_excel(Path(destino))
            elif self._es_modo_plazo():
                bot.generar_plantilla_plazo_excel(Path(destino))
            else:
                bot.generar_plantilla_excel(Path(destino))
            if messagebox.askyesno(
                "Plantilla creada",
                f"Plantilla guardada en:\n{destino}\n\n"
                "¿Quieres abrir el archivo ahora para revisarlo?",
            ):
                os.startfile(destino)
            self.excel_var.set(destino)
            self._analizar_excel_actual(silencioso=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la plantilla:\n{e}")

    def _abrir_carpeta(self):
        os.startfile(str(bot.BASE_DIR))

    def _abrir_reporte(self):
        if self.reporte_generado and Path(self.reporte_generado).exists():
            os.startfile(str(self.reporte_generado))
            return
        messagebox.showinfo("Reporte", "Aún no hay reporte generado en esta sesión.")

    def _continuar_login(self):
        self.login_event.set()
        self.root.after(0, lambda: self._mostrar_panel_login(False))
        self._set_banner("Login confirmado; el bot continúa con la actualización.", self.C_OK_BG, self.C_OK_FG)
        self.quick_status_var.set("Login confirmado, retomando automatización")

    def _serializar_estado(self):
        return {
            "operation": self.operation_var.get(),
            "plazo_mode": self.plazo_mode_var.get(),
            "excel": self.excel_var.get().strip(),
            "acuerdo": self.acuerdo_var.get().strip(),
            "catalogo": self.catalogo_var.get().strip(),
            "categoria": self.categoria_var.get().strip(),
            "region": self.region_var.get().strip(),
            "provincia": self.provincia_var.get().strip(),
            "plazo_general": self.plazo_general_var.get().strip(),
            "pausa": self.pausa_var.get().strip(),
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }

    def _guardar_progreso(self):
        try:
            self._progress_file.write_text(
                json.dumps(self._serializar_estado(), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            self._set_banner("Progreso guardado correctamente. Podrás retomarlo luego desde esta misma ventana.", self.C_OK_BG, self.C_OK_FG)
            messagebox.showinfo("Progreso guardado", f"Se guardó el estado actual en:\n{self._progress_file}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el progreso:\n{exc}")

    def _cargar_progreso(self):
        if not self._progress_file.exists():
            messagebox.showinfo("Sin progreso guardado", "Aún no existe un progreso guardado en esta carpeta del proyecto.")
            return
        try:
            data = json.loads(self._progress_file.read_text(encoding="utf-8"))
            self.operation_var.set(data.get("operation", self.MODO_STOCK))
            self.plazo_mode_var.set(data.get("plazo_mode", self.PLAZO_BLOQUE))
            self.excel_var.set(data.get("excel", ""))
            self.acuerdo_var.set(data.get("acuerdo", ""))
            self.catalogo_var.set(data.get("catalogo", ""))
            self.categoria_var.set(data.get("categoria", ""))
            self.region_var.set(data.get("region", ""))
            self.provincia_var.set(data.get("provincia", ""))
            self.plazo_general_var.set(data.get("plazo_general", "5"))
            self.pausa_var.set(data.get("pausa", str(bot.PAUSA_ENTRE_PRODUCTOS)))
            self._aplicar_modo_operacion_ui()
            self._analizar_excel_actual(silencioso=True)
            self._actualizar_resumen_seleccion()
            saved_at = data.get("saved_at", "desconocido")
            self._set_banner(f"Progreso restaurado. Último guardado: {saved_at}", self.C_OK_BG, self.C_OK_FG)
        except Exception as exc:
            messagebox.showerror("Error al cargar", f"No se pudo restaurar el progreso:\n{exc}")

    def _notificar_login_ui(self):
        def _update():
            self._mostrar_panel_login(True)
            self._set_banner(
                "Inicia sesión en Chrome y luego vuelve a esta ventana para continuar.",
                self.C_LOGIN_BG, "#7B5800",
            )
        self.root.after(0, _update)

    # ------------------------------------------------------------------
    # Cascada de filtros
    # ------------------------------------------------------------------
    def _on_acuerdo_changed(self, event=None):
        self.combo_catalogo["values"] = []
        self.catalogo_var.set("")
        self.combo_categoria["values"] = []
        self.categoria_var.set("")
        self.combo_region["values"] = []
        self.region_var.set("")
        self.combo_provincia["values"] = []
        self.provincia_var.set("")
        self._actualizar_resumen_seleccion()
        self._set_banner("Acuerdo cambiado. Si necesitas sincronizar catálogos y categorías, importa opciones del portal.")

    def _on_catalogo_changed(self, event=None):
        self.combo_categoria["values"] = []
        self.categoria_var.set("")
        if self._es_modo_plazo():
            self.combo_region["values"] = []
            self.region_var.set("")
            self.combo_provincia["values"] = []
            self.provincia_var.set("")
        self._actualizar_resumen_seleccion()
        self._set_banner("Catálogo cambiado. Si necesitas sincronizar categorías, importa opciones del portal.")

    def _on_region_changed(self, event=None):
        self.combo_provincia["values"] = []
        self.provincia_var.set("")
        self._actualizar_resumen_seleccion()
        self._set_banner("Región cambiada. Si necesitas provincias exactas del portal, vuelve a importar opciones.")

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
        self.quick_status_var.set("Preparando importación de filtros desde el portal")
        self._set_banner("Abriendo Chrome para conectarse al portal...", self.C_INFO_BG, self.C_INFO_FG)
        self.worker = threading.Thread(target=self._cargar_opciones_worker, daemon=True)
        self.worker.start()

    def _cargar_opciones_worker(self):
        bot.MODO_GUI = True
        bot.EVENTO_LOGIN = self.login_event
        bot.GUI_NOTIFICAR_LOGIN = self._notificar_login_ui
        driver = None
        try:
            chrome_opts = bot.Options()
            chrome_opts.add_argument("--start-maximized")
            chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
            chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_opts.add_experimental_option("useAutomationExtension", False)
            service = bot.Service(bot.ChromeDriverManager().install())
            driver = bot.webdriver.Chrome(service=service, options=chrome_opts)

            bot.paso1_login(driver)
            if self._es_modo_cobertura():
                bot.paso2_navegacion_cobertura(driver)
            elif self._es_modo_plazo():
                bot.paso2_navegacion_plazo(driver)
            else:
                bot.paso2_navegacion(driver)

            acuerdo_select_id = "cboAcuerdo" if self._es_modo_plazo() else "ajaxAcuerdo"
            catalogo_select_id = "cboCatalogo" if self._es_modo_plazo() else "ajaxCatalogo"
            categoria_select_id = "cboCategoria" if self._es_modo_plazo() else "ajaxCategoria"

            acuerdo_opts = bot.leer_opciones_select(driver, acuerdo_select_id)
            log.info(f"Opciones Acuerdo ({len(acuerdo_opts)}): {acuerdo_opts}")

            catalogo_opts = []
            categoria_opts = []
            region_opts = []
            provincia_opts = []
            acuerdo_actual = self.acuerdo_var.get().strip()
            if acuerdo_opts and not self._es_modo_cobertura():
                try:
                    sel_a = bot.esperar_opciones_select(driver, acuerdo_select_id, bot.WAIT_LARGO)
                    texto_a = acuerdo_actual if acuerdo_actual else acuerdo_opts[0]
                    if self._es_modo_plazo():
                        bot.seleccionar_por_texto_flexible(sel_a, texto_a)
                        driver.execute_script(
                            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                            sel_a._el,
                        )
                    else:
                        bot.seleccionar_por_texto_parcial(sel_a, texto_a)
                    time.sleep(2)
                    catalogo_opts = bot.leer_opciones_select(driver, catalogo_select_id)
                    log.info(f"Opciones Catálogo ({len(catalogo_opts)}): {catalogo_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar catálogos: {e}")

            catalogo_actual = self.catalogo_var.get().strip()
            if catalogo_opts and not self._es_modo_cobertura():
                try:
                    sel_c = bot.esperar_opciones_select(driver, catalogo_select_id, bot.WAIT_LARGO)
                    texto_c = catalogo_actual if catalogo_actual else catalogo_opts[0]
                    if self._es_modo_plazo():
                        bot.seleccionar_por_texto_flexible(sel_c, texto_c)
                        driver.execute_script(
                            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                            sel_c._el,
                        )
                    else:
                        bot.seleccionar_por_texto_parcial(sel_c, texto_c)
                    time.sleep(2)
                    categoria_opts = bot.leer_opciones_select(driver, categoria_select_id)
                    log.info(f"Opciones Categoría ({len(categoria_opts)}): {categoria_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar categorías: {e}")

            if self._es_modo_plazo() and categoria_opts:
                try:
                    sel_cat = bot.esperar_opciones_select(driver, categoria_select_id, bot.WAIT_LARGO)
                    texto_cat = self.categoria_var.get().strip() if self.categoria_var.get().strip() else categoria_opts[0]
                    bot.seleccionar_por_texto_flexible(sel_cat, texto_cat)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                        sel_cat._el,
                    )
                    time.sleep(2)
                    region_opts = bot.leer_opciones_select(driver, "cboRegion")
                    log.info(f"Opciones Región ({len(region_opts)}): {region_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar regiones: {e}")

            if self._es_modo_plazo() and region_opts:
                try:
                    sel_region = bot.esperar_opciones_select(driver, "cboRegion", bot.WAIT_LARGO)
                    texto_region = self.region_var.get().strip() if self.region_var.get().strip() else region_opts[0]
                    bot.seleccionar_por_texto_flexible(sel_region, texto_region)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                        sel_region._el,
                    )
                    time.sleep(2)
                    provincia_opts = bot.leer_opciones_select(driver, "cboProvincia")
                    log.info(f"Opciones Provincia ({len(provincia_opts)}): {provincia_opts}")
                except Exception as e:
                    log.warning(f"No se pudo cargar provincias: {e}")

            self.root.after(0, lambda: self._actualizar_combos(acuerdo_opts, catalogo_opts, categoria_opts, region_opts, provincia_opts))

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

    def _actualizar_combos(self, acuerdos, catalogos, categorias, regiones=None, provincias=None):
        regiones = regiones or []
        provincias = provincias or []
        self.combo_acuerdo["values"] = acuerdos
        self.combo_catalogo["values"] = catalogos
        self.combo_categoria["values"] = categorias
        self.combo_region["values"] = regiones
        self.combo_provincia["values"] = provincias
        self._portal_snapshot = {
            "acuerdos": len(acuerdos),
            "catalogos": len(catalogos),
            "categorias": len(categorias),
        }
        self.metric_portal_var.set(
            f"{len(acuerdos)}/{len(catalogos)}/{len(categorias)}"
        )

        # Asignar primer valor si el campo estaba vacío
        if acuerdos and not self.acuerdo_var.get():
            self.acuerdo_var.set(acuerdos[0])
        if catalogos and not self.catalogo_var.get():
            self.catalogo_var.set(catalogos[0])
        if categorias and not self.categoria_var.get():
            self.categoria_var.set(categorias[0])
        if regiones and not self.region_var.get():
            self.region_var.set(regiones[0])
        if provincias and not self.provincia_var.get():
            self.provincia_var.set(provincias[0])

        if self._es_modo_cobertura():
            self.catalogo_var.set("")
            self.categoria_var.set("")

        self._actualizar_resumen_seleccion()

        if self._es_modo_cobertura():
            self._set_banner(
                f"Filtros cargados: {len(acuerdos)} acuerdo(s) disponibles para cobertura.",
                self.C_OK_BG, self.C_OK_FG,
            )
            self.quick_status_var.set("Acuerdos de cobertura importados correctamente")
            messagebox.showinfo(
                "Filtros cargados",
                f"Se cargaron desde el portal {len(acuerdos)} acuerdo(s) marco para cobertura.\n\n"
                "Ahora confirma el acuerdo correcto y ejecuta el proceso.",
            )
        elif self._es_modo_plazo():
            self._set_banner(
                f"Filtros de plazo cargados: {len(acuerdos)} acuerdos, {len(catalogos)} catálogos, {len(categorias)} categorías, {len(regiones)} regiones y {len(provincias)} provincias.",
                self.C_OK_BG, self.C_OK_FG,
            )
            self.quick_status_var.set("Filtros de plazo importados correctamente")
            messagebox.showinfo(
                "Filtros cargados",
                f"Se cargaron desde el portal:\n"
                f"  • {len(acuerdos)} Acuerdo(s) Marco\n"
                f"  • {len(catalogos)} Catálogo(s)\n"
                f"  • {len(categorias)} Categoría(s)\n"
                f"  • {len(regiones)} Región(es)\n"
                f"  • {len(provincias)} Provincia(s)\n\n"
                "Ahora confirma los valores correctos para el módulo de plazo.",
            )
        else:
            self._set_banner(
                f"Filtros cargados: {len(acuerdos)} acuerdos, {len(catalogos)} catálogos y {len(categorias)} categorías.",
                self.C_OK_BG, self.C_OK_FG,
            )
            self.quick_status_var.set("Filtros importados correctamente")
            messagebox.showinfo(
                "Filtros cargados",
                f"Se cargaron desde el portal:\n"
                f"  • {len(acuerdos)} Acuerdo(s) Marco\n"
                f"  • {len(catalogos)} Catálogo(s)\n"
                f"  • {len(categorias)} Categoría(s)\n\n"
                f"Ahora selecciona los valores correctos en los desplegables del Paso 2.",
            )

    # ------------------------------------------------------------------
    # Pausa / Detección / Aprendizaje
    # ------------------------------------------------------------------
    def _pausar_reanudar(self):
        if not self._pausado:
            if bot.PAUSA_EVENTO:
                bot.PAUSA_EVENTO.clear()
            self._pausado = True
            self.btn_pausar.configure(text="Reanudar")
            self._set_banner("Ejecución en pausa. Puedes reanudar cuando quieras.", self.C_LOGIN_BG, "#7B5800")
            self.quick_status_var.set("Proceso pausado por el usuario")
            log.info("Ejecución pausada por el usuario.")
        else:
            if bot.PAUSA_EVENTO:
                bot.PAUSA_EVENTO.set()
            self._pausado = False
            self.btn_pausar.configure(text="Pausar")
            self._set_banner("Reanudando proceso...", self.C_INFO_BG, self.C_INFO_FG)
            self.quick_status_var.set("Proceso reanudado")
            log.info("Ejecución reanudada por el usuario.")

    def _detener(self):
        if not messagebox.askyesno(
            "Detener proceso",
            "¿Seguro que quieres detener la automatización?\n\n"
            "Se generará el reporte Excel con los resultados hasta el momento.",
        ):
            return
        if bot.DETENER_EVENTO:
            bot.DETENER_EVENTO.set()
        if bot.PAUSA_EVENTO:
            bot.PAUSA_EVENTO.set()
        self._pausado = False
        self.btn_pausar.configure(text="Pausar")
        self._set_banner("Deteniendo el proceso. Se completará el producto actual y luego se generará el reporte.",
                         "#FFF3E0", "#E65100")
        self.quick_status_var.set("Detención solicitada; generando cierre controlado")
        log.info("Detener solicitado por el usuario.")

    def _ver_aprendizaje(self):
        arch = bot.BASE_DIR / "aprendizaje.json"
        if not arch.exists():
            messagebox.showinfo(
                "Estadísticas de errores",
                "Aún no hay datos registrados.\n\n"
                "El bot guarda estadísticas de cada error que encuentra.\n"
                "Después de la primera ejecución aparecerán aquí.",
            )
            return
        try:
            data = json.loads(arch.read_text(encoding="utf-8"))
            acum = data.get("acumulado", {})
            sesion = data.get("ultima_sesion", "Desconocida")
            if not acum:
                messagebox.showinfo("Estadísticas", "Aún no se han registrado errores.")
                return
            lineas = [
                f"Última sesión: {sesion}\n",
                "Errores encontrados (histórico):",
            ]
            for tipo, cnt in sorted(acum.items(), key=lambda x: -x[1]):
                estado = " ✔ ajuste activo" if cnt >= bot.AnalizadorFallos.UMBRAL else ""
                lineas.append(f"  • {tipo}: {cnt} vez/veces{estado}")
            lineas += [
                "",
                f"Cuando un error se repite {bot.AnalizadorFallos.UMBRAL}+ veces, el bot ajusta",
                "automáticamente su comportamiento para evitarlo.",
                "",
                "Para reiniciar las estadísticas: elimina el archivo 'aprendizaje.json'.",
            ]
            messagebox.showinfo("Estadísticas de errores", "\n".join(lineas))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def _iniciar(self):
        excel = Path(self.excel_var.get().strip())
        acuerdo = self.acuerdo_var.get().strip()
        catalogo = self.catalogo_var.get().strip()
        categoria = self.categoria_var.get().strip()
        region = self.region_var.get().strip()
        provincia = self.provincia_var.get().strip()
        plazo_general_txt = self.plazo_general_var.get().strip()
        pausa_txt = self.pausa_var.get().strip()

        resumen = self._analizar_excel_actual(silencioso=True)
        if self._requiere_excel() and (not resumen or not excel.exists()):
            messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontró el archivo:\n{excel}\n\n"
                "Usa el botón 'Buscar archivo' para seleccionar tu Excel."
            )
            return
        if self._requiere_excel() and not resumen.is_ready:
            mensaje = "Corrige el Excel antes de iniciar:\n\n- " + "\n- ".join(resumen.blocking_issues)
            if resumen.issue_examples:
                mensaje += "\n\nEjemplos detectados:\n- " + "\n- ".join(resumen.issue_examples[:4])
            messagebox.showerror("El archivo todavía no está listo", mensaje)
            return
        if self._es_modo_cobertura() and not acuerdo:
            messagebox.showerror(
                "Filtros incompletos",
                "Debes completar el Acuerdo Marco del Paso 2.\n\nUsa 'Importar opciones del portal' si el desplegable está vacío.",
            )
            return
        if self._es_modo_plazo() and (not acuerdo or not catalogo or not categoria or not region or not provincia):
            messagebox.showerror(
                "Filtros incompletos",
                "Debes completar los filtros de plazo:\n"
                "  • Acuerdo Marco\n  • Catálogo\n  • Categoría\n  • Región\n  • Provincia\n\n"
                "Usa 'Importar opciones del portal' si necesitas traer las listas exactas.",
            )
            return
        if not self._es_modo_cobertura() and not self._es_modo_plazo() and (not acuerdo or not catalogo or not categoria):
            messagebox.showerror(
                "Filtros incompletos",
                "Debes completar los tres filtros del Paso 2:\n"
                "  • Acuerdo Marco\n  • Catálogo\n  • Categoría\n\n"
                "Usa 'Importar opciones del portal' si los desplegables están vacíos."
            )
            return
        try:
            pausa = int(pausa_txt)
            if pausa < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Valor inválido", "La pausa debe ser un número entero mayor o igual a 0.")
            return

        if self._es_modo_plazo() and not self._es_plazo_individual():
            try:
                plazo_general = int(plazo_general_txt)
                if plazo_general < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Valor inválido", "El plazo general debe ser un número entero mayor o igual a 0.")
                return
        else:
            plazo_general = None

        # Resetear UI de resultado anterior
        self._mostrar_panel_resultado(False)
        self._mostrar_panel_ctrl(True)
        self._mostrar_panel_login(False)
        self._actualizar_progreso(0, 0, "Iniciando...")
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(text="Pausar")
        self._pausado = False
        self._set_banner("Iniciando automatización y abriendo Chrome...", self.C_INFO_BG, self.C_INFO_FG)
        self.quick_status_var.set(f"Ejecutando automatización de {self._texto_operacion()}")
        self._aplicar_estado_preparacion(
            "Automatización en curso",
            f"El bot está trabajando en {self._texto_operacion()}. Puedes pausar, detener o seguir el progreso desde esta pantalla.",
            "info",
        )
        self._actualizar_resumen_seleccion()

        self.worker = threading.Thread(
            target=self._worker_run,
            args=(excel, acuerdo, catalogo, categoria, region, provincia, pausa, self.operation_var.get(), self.plazo_mode_var.get(), plazo_general),
            daemon=True,
        )
        self.worker.start()

    def _worker_run(self, excel, acuerdo, catalogo, categoria, region, provincia, pausa, modo, plazo_mode, plazo_general):
        bot.MODO_GUI = True
        bot.EVENTO_LOGIN = self.login_event
        bot.GUI_NOTIFICAR_LOGIN = self._notificar_login_ui
        bot.PAUSA_EVENTO = threading.Event()
        bot.PAUSA_EVENTO.set()
        bot.DETENER_EVENTO = threading.Event()

        try:
            if modo == self.MODO_COBERTURA:
                reporte = bot.ejecutar_bot_cobertura(
                    excel_path=excel,
                    acuerdo_texto=acuerdo,
                    pausa_entre_productos=pausa,
                )
            elif modo == self.MODO_PLAZO:
                reporte = bot.ejecutar_bot_plazo(
                    acuerdo_texto=acuerdo,
                    catalogo_texto=catalogo,
                    categoria_texto=categoria,
                    region_texto=region,
                    provincia_texto=provincia,
                    modo_carga=plazo_mode,
                    pausa_entre_productos=pausa,
                    plazo_general=plazo_general,
                    excel_path=excel if plazo_mode == self.PLAZO_INDIVIDUAL else None,
                )
            else:
                reporte = bot.ejecutar_bot(
                    excel_path=excel,
                    acuerdo_texto=acuerdo,
                    catalogo_texto=catalogo,
                    categoria_texto=categoria,
                    pausa_entre_productos=pausa,
                )
            self.reporte_generado = reporte
            total = len(bot.RESULTADOS)
            exitos = sum(1 for r in bot.RESULTADOS if r["Estado"] == "EXITO")
            fallos = total - exitos
            etiqueta = "región(es)" if modo == self.MODO_COBERTURA else "bloque(s)/producto(s)"
            info = (
                f"{exitos} {etiqueta} actualizados correctamente"
                + (f"   ·   {fallos} con error(es)" if fallos else "")
                + f"\n\nReporte guardado en:\n{reporte}"
            )
            self.metric_reporte_var.set(Path(reporte).name)
            self.root.after(0, lambda: self._lbl_resultado_info.configure(text=info))
            self.root.after(0, lambda: self._mostrar_panel_ctrl(False))
            self.root.after(0, lambda: self._mostrar_panel_resultado(True))
            self.root.after(0, lambda: self._set_banner(
                f"Proceso completado: {exitos}/{total} {etiqueta} actualizados.",
                self.C_OK_BG, self.C_OK_FG,
            ))
            self.root.after(0, lambda: self._actualizar_progreso(total, total, "Completado"))
            self.root.after(0, lambda: self._aplicar_estado_preparacion(
                "Proceso finalizado",
                f"Se generó un reporte con {exitos} éxito(s) y {fallos} fallo(s).",
                "ok",
            ))
        except Exception as e:
            detalle = f"{e}\n\n{traceback.format_exc()}"
            log.error(f"Error fatal: {e}", exc_info=True)
            self.root.after(0, lambda: self._set_banner(
                "Error en la ejecución. Revisa la actividad detallada para soporte.",
                "#FDECEA", self.C_PELIGRO,
            ))
            self.root.after(0, lambda: self._aplicar_estado_preparacion(
                "Error en ejecución",
                "La automatización se interrumpió. Revisa el detalle técnico y corrige la causa antes de reintentar.",
                "error",
            ))
            self.root.after(0, lambda: messagebox.showerror("Error inesperado", detalle))
        finally:
            self.root.after(0, lambda: self.btn_iniciar.configure(state="normal"))
            self._pausado = False


def iniciar_interfaz():
    root = tk.Tk()
    root.title("Peru Compras Bot")
    root.geometry("1180x820")
    root.minsize(980, 700)
    root.resizable(True, True)

    icon_path = bot.BASE_DIR / "assets" / "app_mascot.ico"
    if icon_path.exists():
        try:
            root.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    root.withdraw()
    root.update_idletasks()
    root.deiconify()

    PeruComprasGUI(root)
    root.mainloop()
