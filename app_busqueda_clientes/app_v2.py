# ---------------------------------------------------------------------------
# Interfaz gráfica — aplicación
# ---------------------------------------------------------------------------

class ProspeccionApp(tk.Tk):
    """Buscador de negocios + prospectos guardados."""

    def __init__(self, ruta_db: str = RUTA_DB, ruta_palabras: str = RUTA_PALABRAS) -> None:
        super().__init__()
        self.title("Prospección de Clientes — Buscador de negocios")
        self.geometry("1280x780")
        self.minsize(1020, 640)

        self._db = BaseDatos(ruta_db)
        self._uso = ApiUsageTracker(self._db)
        self._ruta_palabras = ruta_palabras
        self._palabras, self._aviso_json, self._json_invalido = cargar_palabras_clave(ruta_palabras)
        self._cancelar = threading.Event()
        self._cola: "queue.Queue[tuple]" = queue.Queue()
        self._sesion: Optional[SesionBusqueda] = None
        self._pagina_actual = 0
        self._ocupado = False
        self._visibles: list[Negocio] = []
        self._por_iid: dict[str, Negocio] = {}
        self._guardados: dict[str, Negocio] = {}
        self._nota_id: Optional[str] = None
        self._suprimir_filtro = False
        self._editor_palabras: Optional[tk.Toplevel] = None

        self._configurar_estilos()
        self._build_ui()
        self._centrar_ventana()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.bind("<F5>", lambda _e: self._on_buscar())
        self.bind("<Control-f>", lambda _e: self.entry_filtro_b.focus_set())
        self.after(100, self._procesar_cola)
        self._refrescar_guardados()
        self._actualizar_botones()
        self._actualizar_uso()
        if self._aviso_json:
            self._set_estado(self._aviso_json, "aviso")

    # ------------------------------------------------------------------ estilos
    def _configurar_estilos(self) -> None:
        familias = set(tkfont.families(self))
        fam = next(
            (f for f in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Noto Sans", "DejaVu Sans", "Arial") if f in familias),
            tkfont.nametofont("TkDefaultFont").actual("family"),
        )
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkFixedFont"):
            try:
                tkfont.nametofont(nombre).configure(size=10)
                if nombre != "TkFixedFont":
                    tkfont.nametofont(nombre).configure(family=fam)
            except tk.TclError:
                pass
        self.f_peq = (fam, 9)
        self.f_neg = (fam, 10, "bold")
        self.f_tit = (fam, 16, "bold")
        c = COLORES
        st = ttk.Style(self)
        if "clam" in st.theme_names():
            st.theme_use("clam")
        self.configure(background=c["fondo"])
        st.configure(".", background=c["fondo"], foreground=c["texto"], font=(fam, 10), bordercolor=c["borde"])
        st.configure("TFrame", background=c["fondo"])
        st.configure("Card.TFrame", background=c["tarjeta"], relief="solid", borderwidth=1, bordercolor=c["borde"])
        st.configure("TLabel", background=c["fondo"], foreground=c["texto"])
        st.configure("Card.TLabel", background=c["tarjeta"])
        st.configure("Suave.TLabel", foreground=c["suave"], font=(fam, 9))
        st.configure("CardSuave.TLabel", background=c["tarjeta"], foreground=c["suave"], font=(fam, 9))
        st.configure("Negrita.TLabel", font=(fam, 10, "bold"))
        st.configure("Header.TFrame", background=c["primario"])
        st.configure("HeaderTitulo.TLabel", background=c["primario"], foreground="white", font=(fam, 16, "bold"))
        st.configure("HeaderSub.TLabel", background=c["primario"], foreground="#CFE3F7", font=(fam, 9))
        st.configure("TButton", padding=(12, 6), background="#FFFFFF", bordercolor=c["borde"], focuscolor="#FFFFFF")
        st.map("TButton", background=[("active", c["primario_claro"]), ("disabled", "#EEF0F3")],
               foreground=[("disabled", "#A0A7B1")])
        st.configure("Primary.TButton", background=c["primario"], foreground="white", bordercolor=c["primario"],
                     font=(fam, 10, "bold"), padding=(16, 7))
        st.map("Primary.TButton", background=[("active", c["primario_osc"]), ("disabled", "#9DBFE0")],
               foreground=[("disabled", "#FFFFFF")])
        st.configure("Header.TButton", background="#2A7FCB", foreground="white", bordercolor="#2A7FCB", padding=(12, 5))
        st.map("Header.TButton", background=[("active", c["primario_osc"])])
        st.configure("TMenubutton", padding=(12, 6), background="#FFFFFF", bordercolor=c["borde"], arrowcolor=c["texto"])
        st.map("TMenubutton", background=[("active", c["primario_claro"]), ("disabled", "#EEF0F3")],
               foreground=[("disabled", "#A0A7B1")])
        st.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], foreground=[("readonly", c["texto"])])
        st.configure("TEntry", fieldbackground="#FFFFFF", padding=5)
        st.configure("TNotebook", background=c["fondo"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        st.configure("TNotebook.Tab", padding=(22, 9), background="#E3E8EF", foreground=c["suave"], font=(fam, 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", c["tarjeta"])], foreground=[("selected", c["primario"])])
        st.configure("Datos.Treeview", rowheight=28, background=c["tarjeta"], fieldbackground=c["tarjeta"],
                     foreground=c["texto"], bordercolor=c["borde"], borderwidth=1, font=(fam, 10))
        st.map("Datos.Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["texto"])])
        st.configure("Datos.Treeview.Heading", background="#E6ECF4", foreground=c["texto"], font=(fam, 10, "bold"),
                     relief="flat", padding=(6, 7))
        st.map("Datos.Treeview.Heading", background=[("active", "#D6E2F0")])
        st.configure("Horizontal.TProgressbar", background=c["primario"], troughcolor="#E3E8EF", bordercolor=c["borde"])
        # Tk 8.6.9: los colores por etiqueta del Treeview no se aplicaban sin este arreglo
        def _sin_bug(opcion: str):
            return [e for e in st.map("Treeview", query_opt=opcion) if e[:2] != ("!disabled", "!selected")]
        st.map("Treeview", foreground=_sin_bug("foreground"), background=_sin_bug("background"))

    # --------------------------------------------------------------------- UI
    def _centrar_ventana(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
        y = max(0, (self.winfo_screenheight() // 2) - (h // 2))
        self.geometry(f"+{x}+{y}")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # --- cabecera ---
        cab = ttk.Frame(self, style="Header.TFrame", padding=(18, 12))
        cab.grid(row=0, column=0, sticky="ew")
        cab.columnconfigure(0, weight=1)
        titulos = ttk.Frame(cab, style="Header.TFrame")
        titulos.grid(row=0, column=0, sticky="w")
        ttk.Label(titulos, text="Buscador de negocios", style="HeaderTitulo.TLabel").pack(anchor="w")
        ttk.Label(titulos, text="Google Maps · prospectos guardados en SQLite", style="HeaderSub.TLabel").pack(anchor="w")
        ttk.Button(cab, text="Palabras clave", style="Header.TButton", command=self._abrir_editor_palabras).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(cab, text="Info de uso", style="Header.TButton", command=self._mostrar_info_api).grid(row=0, column=2)

        # --- pestañas ---
        self.nb = ttk.Notebook(self)
        self.nb.grid(row=1, column=0, sticky="nsew", padx=12, pady=(10, 4))
        self.tab_b = ttk.Frame(self.nb, padding=(0, 10, 0, 0))
        self.tab_g = ttk.Frame(self.nb, padding=(0, 10, 0, 0))
        self.nb.add(self.tab_b, text="Búsqueda")
        self.nb.add(self.tab_g, text="★ Guardados")
        self.nb.bind("<<NotebookTabChanged>>", self._on_cambio_pestana)
        self._build_tab_busqueda()
        self._build_tab_guardados()

        # --- barra de estado ---
        barra = ttk.Frame(self, padding=(14, 4, 14, 8))
        barra.grid(row=2, column=0, sticky="ew")
        barra.columnconfigure(0, weight=1)
        self.lbl_estado = ttk.Label(barra, text="Listo para buscar", style="Suave.TLabel")
        self.lbl_estado.grid(row=0, column=0, sticky="w")
        self.progreso = ttk.Progressbar(barra, mode="indeterminate", length=140)
        self.progreso.grid(row=0, column=1, padx=10)
        self.progreso.grid_remove()
        self.lbl_uso = ttk.Label(barra, text="", style="Suave.TLabel")
        self.lbl_uso.grid(row=0, column=2, sticky="e")

    def _crear_menu_exportar(self, parent: tk.Widget, origen: str) -> ttk.Menubutton:
        mb = ttk.Menubutton(parent, text="Exportar  ▾")
        menu = tk.Menu(mb, tearoff=0)
        for fmt, etiqueta in (("xlsx", "Excel"), ("csv", "CSV")):
            menu.add_command(label=f"{etiqueta} · todo lo cargado" if origen == "busqueda" else f"{etiqueta} · todos los guardados",
                             command=lambda f=fmt: self._exportar(origen, "todo", f))
            menu.add_command(label=f"{etiqueta} · solo lo filtrado en la tabla", command=lambda f=fmt: self._exportar(origen, "filtrado", f))
            menu.add_command(label=f"{etiqueta} · solo lo seleccionado", command=lambda f=fmt: self._exportar(origen, "seleccion", f))
            menu.add_separator()
        if origen == "busqueda":
            menu.add_command(label="Cargar todo y exportar a Excel…", command=self._on_cargar_todo)
        else:
            menu.delete(menu.index("end"))     # quita el último separador
        mb["menu"] = menu
        return mb

    def _build_tab_busqueda(self) -> None:
        t = self.tab_b
        t.columnconfigure(0, weight=1)
        t.rowconfigure(3, weight=1)

        card = ttk.Frame(t, style="Card.TFrame", padding=(14, 12))
        card.grid(row=0, column=0, sticky="ew")
        for col in (1, 3, 5, 7, 9):
            card.columnconfigure(col, weight=1)
        ttk.Label(card, text="País", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.var_pais = tk.StringVar(value="Colombia")
        self.combo_pais = ttk.Combobox(card, textvariable=self.var_pais, values=PAISES, state="readonly", width=16)
        self.combo_pais.grid(row=0, column=1, sticky="ew", padx=(0, 14))
        self.combo_pais.bind("<<ComboboxSelected>>", self._on_pais_change)

        ttk.Label(card, text="Ciudad", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.var_ciudad = tk.StringVar(value=CIUDADES_POR_PAIS["Colombia"][0])
        self.combo_ciudad = ttk.Combobox(card, textvariable=self.var_ciudad, values=CIUDADES_POR_PAIS["Colombia"], state="readonly", width=16)
        self.combo_ciudad.grid(row=0, column=3, sticky="ew", padx=(0, 14))

        ttk.Label(card, text="Tipo de negocio", style="Card.TLabel").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self.var_tipo = tk.StringVar(value="Médicos" if "Médicos" in self._palabras else next(iter(self._palabras)))
        self.combo_tipo = ttk.Combobox(card, textvariable=self.var_tipo, values=list(self._palabras), state="readonly", width=16)
        self.combo_tipo.grid(row=0, column=5, sticky="ew", padx=(0, 14))
        self.combo_tipo.bind("<<ComboboxSelected>>", lambda _e: self._actualizar_resumen_terminos())

        ttk.Label(card, text="Mín. reseñas", style="Card.TLabel").grid(row=0, column=6, sticky="w", padx=(0, 6))
        self.var_min = tk.StringVar(value="5")
        self.combo_min = ttk.Combobox(card, textvariable=self.var_min, values=MIN_RESEÑAS_OPCIONES, state="readonly", width=6)
        self.combo_min.grid(row=0, column=7, sticky="w", padx=(0, 14))
        self.combo_min.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        ttk.Label(card, text="Por página", style="Card.TLabel").grid(row=0, column=8, sticky="w", padx=(0, 6))
        self.var_pp = tk.StringVar(value="25")
        self.combo_pp = ttk.Combobox(card, textvariable=self.var_pp, values=OPCIONES_POR_PAGINA, state="readonly", width=6)
        self.combo_pp.grid(row=0, column=9, sticky="w", padx=(0, 14))
        self.combo_pp.bind("<<ComboboxSelected>>", self._on_cambio_reaplicar)

        botones = ttk.Frame(card, style="Card.TFrame")
        botones.grid(row=0, column=10, sticky="e")
        self.btn_buscar = ttk.Button(botones, text="Buscar negocios", style="Primary.TButton", command=self._on_buscar)
        self.btn_buscar.pack(side="left")
        self.btn_cancelar = ttk.Button(botones, text="Cancelar", command=self._on_cancelar, state="disabled")
        self.btn_cancelar.pack(side="left", padx=(8, 0))

        self.lbl_terminos = ttk.Label(card, text="", style="CardSuave.TLabel")
        self.lbl_terminos.grid(row=1, column=0, columnspan=11, sticky="w", pady=(8, 0))

        # barra de herramientas
        tools = ttk.Frame(t)
        tools.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        tools.columnconfigure(1, weight=1)
        ttk.Label(tools, text="Filtrar en la página").grid(row=0, column=0, padx=(0, 8))
        self.var_filtro_b = tk.StringVar()
        self.var_filtro_b.trace_add("write", lambda *_: self._on_filtro_b())
        self.entry_filtro_b = ttk.Entry(tools, textvariable=self.var_filtro_b, width=34)
        self.entry_filtro_b.grid(row=0, column=1, sticky="w")
        self.lbl_leads_b = ttk.Label(tools, text="0 leads", style="Negrita.TLabel", foreground=COLORES["ok"])
        self.lbl_leads_b.grid(row=0, column=2, padx=14)
        self.btn_guardar_b = ttk.Button(tools, text="★ Guardar / quitar", command=lambda: self._guardar_seleccion("busqueda"))
        self.btn_guardar_b.grid(row=0, column=3, padx=4)
        self.btn_msg_b = ttk.Button(tools, text="Copiar mensaje", command=self._on_copiar_mensaje)
        self.btn_msg_b.grid(row=0, column=4, padx=4)
        self.mb_export_b = self._crear_menu_exportar(tools, "busqueda")
        self.mb_export_b.grid(row=0, column=5, padx=(4, 0))

        # paginación
        pag = ttk.Frame(t)
        pag.grid(row=2, column=0, sticky="w", pady=(8, 8))
        self.btn_ant = ttk.Button(pag, text="◀ Anterior", command=self._pagina_anterior)
        self.btn_ant.pack(side="left")
        self.lbl_pagina = ttk.Label(pag, text="Página —", style="Negrita.TLabel", padding=(14, 0))
        self.lbl_pagina.pack(side="left")
        self.btn_sig = ttk.Button(pag, text="Siguiente ▶", command=self._pagina_siguiente)
        self.btn_sig.pack(side="left")
        ttk.Label(pag, text="  ·  Clic en ☆ para guardar  ·  doble clic en una celda para editar  ·  clic derecho para más acciones",
                  style="Suave.TLabel").pack(side="left", padx=(10, 0))

        self.tabla_b = TablaDatos(
            t, COLS_BUSQUEDA,
            on_estrella=lambda iid: self._toggle_guardado(self._por_iid.get(iid)),
            on_editar=lambda iid, col, val: self._aplicar_edicion(self._por_iid.get(iid), col, val),
            on_contexto=lambda ev: self._menu_contexto("busqueda", ev),
            on_doble_estado=lambda iid, x, y: self._menu_estado_en(("busqueda", x, y)),
            on_seleccion=self._actualizar_botones,
            on_mensaje=self._mensaje_tabla,
        )
        self.tabla_b.grid(row=3, column=0, sticky="nsew")
        self._actualizar_resumen_terminos()

    def _build_tab_guardados(self) -> None:
        t = self.tab_g
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)

        tools = ttk.Frame(t)
        tools.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tools.columnconfigure(1, weight=1)
        ttk.Label(tools, text="Filtrar").grid(row=0, column=0, padx=(0, 8))
        self.var_filtro_g = tk.StringVar()
        self.var_filtro_g.trace_add("write", lambda *_: self._refrescar_guardados())
        self.entry_filtro_g = ttk.Entry(tools, textvariable=self.var_filtro_g, width=34)
        self.entry_filtro_g.grid(row=0, column=1, sticky="w")
        self.lbl_leads_g = ttk.Label(tools, text="0 guardados", style="Negrita.TLabel", foreground=COLORES["ok"])
        self.lbl_leads_g.grid(row=0, column=2, padx=14)
        self.mb_estado_g = ttk.Menubutton(tools, text="Cambiar estado  ▾")
        self.mb_estado_g["menu"] = self._crear_menu_estado(self.mb_estado_g, "guardados")
        self.mb_estado_g.grid(row=0, column=3, padx=4)
        self.btn_tel_g = ttk.Button(tools, text="Copiar teléfono", command=lambda: self._copiar_telefonos("guardados"))
        self.btn_tel_g.grid(row=0, column=4, padx=4)
        self.btn_quitar_g = ttk.Button(tools, text="Quitar de guardados", command=lambda: self._quitar_guardados())
        self.btn_quitar_g.grid(row=0, column=5, padx=4)
        self.mb_export_g = self._crear_menu_exportar(tools, "guardados")
        self.mb_export_g.grid(row=0, column=6, padx=(4, 0))

        self.tabla_g = TablaDatos(
            t, COLS_GUARDADOS,
            on_estrella=lambda iid: self._quitar_guardados([self._guardados[iid]] if iid in self._guardados else []),
            on_editar=lambda iid, col, val: self._aplicar_edicion(self._guardados.get(iid), col, val),
            on_contexto=lambda ev: self._menu_contexto("guardados", ev),
            on_doble_estado=lambda iid, x, y: self._menu_estado_en(("guardados", x, y)),
            on_seleccion=self._on_sel_guardados,
            on_mensaje=self._mensaje_tabla,
        )
        self.tabla_g.grid(row=1, column=0, sticky="nsew")

        panel = ttk.Frame(t, style="Card.TFrame", padding=(12, 8))
        panel.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        panel.columnconfigure(0, weight=1)
        self.lbl_nota = ttk.Label(panel, text="Notas — selecciona un prospecto", style="Card.TLabel", font=self.f_neg)
        self.lbl_nota.grid(row=0, column=0, sticky="w")
        self.txt_nota = tk.Text(panel, height=3, wrap="word", relief="solid", bd=1, state="disabled",
                                highlightthickness=1, highlightbackground=COLORES["borde"], highlightcolor=COLORES["primario"])
        self.txt_nota.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.txt_nota.bind("<FocusOut>", lambda _e: self._guardar_nota_pendiente())
        self.txt_nota.bind("<Control-Return>", lambda _e: (self._guardar_nota_pendiente(), "break")[1])
        self.btn_nota = ttk.Button(panel, text="Guardar nota", command=self._guardar_nota_pendiente, state="disabled")
        self.btn_nota.grid(row=1, column=1, padx=(10, 0), sticky="s")

    # --------------------------------------------------------------- utilidades UI
    def _set_estado(self, texto: str, nivel: str = "normal") -> None:
        color = {"normal": COLORES["suave"], "ok": COLORES["ok"], "aviso": COLORES["aviso"], "error": COLORES["error"]}[nivel]
        self.lbl_estado.config(text=texto, foreground=color)

    def _mensaje_tabla(self, texto: str) -> None:
        if not self._ocupado:
            self._set_estado(texto)

    def _actualizar_uso(self) -> None:
        self.lbl_uso.config(text=f"Peticiones: {self._uso.peticiones_sesion} esta sesión · {self._db.uso_mes()} este mes")

    def _tipo_terminos(self, tipo: str) -> list[str]:
        return self._palabras.get(tipo) or [tipo]

    def _actualizar_resumen_terminos(self) -> None:
        n = len(self._tipo_terminos(self.var_tipo.get()))
        self.lbl_terminos.config(
            text=f"«{self.var_tipo.get()}»: {n} términos de búsqueda · 1 consulta por término (hasta 3 páginas de 20 resultados cada una)"
        )

    def _on_pais_change(self, *_: object) -> None:
        ciudades = CIUDADES_POR_PAIS.get(self.var_pais.get(), [])
        if ciudades:
            self.combo_ciudad["values"] = ciudades
            self.var_ciudad.set(ciudades[0])

    def _limite_pagina(self) -> int:
        if self.var_pp.get().strip().lower() == "todos":
            return PAGINA_TODOS
        try:
            return max(1, int(self.var_pp.get()))
        except ValueError:
            return 10

    def _min_resenas(self) -> int:
        try:
            return max(0, int(self.var_min.get()))
        except ValueError:
            return 0

    def _actualizar_botones(self) -> None:
        busy = self._ocupado
        s = self._sesion
        tiene = s is not None and self._pagina_actual > 0 and self._pagina_actual in s.paginas
        sel_b = bool(self.tabla_b.seleccion())
        sel_g = bool(self.tabla_g.seleccion())
        est = lambda cond: "normal" if cond else "disabled"
        self.btn_buscar.config(state=est(not busy))
        self.btn_cancelar.config(state=est(busy))
        self.btn_ant.config(state=est(tiene and not busy and self._pagina_actual > 1))
        sig = tiene and ((self._pagina_actual + 1) in s.paginas or s.hay_mas())
        self.btn_sig.config(state=est(sig and not busy))
        self.lbl_pagina.config(text=f"Página {self._pagina_actual}" if tiene else "Página —")
        self.mb_export_b.config(state=est(s is not None and not busy))
        self.btn_guardar_b.config(state=est(sel_b and not busy))
        self.btn_msg_b.config(state=est(sel_b))
        self.mb_estado_g.config(state=est(sel_g))
        self.btn_tel_g.config(state=est(sel_g))
        self.btn_quitar_g.config(state=est(sel_g))
        self.mb_export_g.config(state=est(bool(self._guardados)))

    def _actualizar_contador_b(self) -> None:
        visibles = len(self._visibles)
        s = self._sesion
        if s is not None and self._pagina_actual > 0:
            texto = f"Pág. {self._pagina_actual}: {visibles} leads · {len(s.raw)} en caché"
        elif s is not None:
            texto = "0 leads"
        else:
            texto = "0 leads"
        self.lbl_leads_b.config(text=texto, foreground=COLORES["ok"] if visibles else COLORES["error"])

    # -------------------------------------------------- hilos / cola de mensajes
    def _lanzar(self, trabajo: Callable[[], object], al_terminar: Callable[[object, Optional[BaseException]], None], texto: str) -> None:
        self._ocupado = True
        self._cancelar.clear()
        self.progreso.grid()
        self.progreso.start(12)
        self._set_estado(texto, "aviso")
        self._actualizar_botones()

        def run() -> None:
            try:
                self._cola.put(("fin", al_terminar, trabajo(), None))
            except BaseException as exc:      # noqa: BLE001 — se informa en la interfaz
                self._cola.put(("fin", al_terminar, None, exc))

        threading.Thread(target=run, daemon=True).start()

    def _procesar_cola(self) -> None:
        try:
            while True:
                item = self._cola.get_nowait()
                if item[0] == "progreso":
                    if self._ocupado:
                        self._set_estado(item[1], "aviso")
                        self._actualizar_uso()
                elif item[0] == "fin":
                    _, callback, resultado, exc = item
                    self._ocupado = False
                    self.progreso.stop()
                    self.progreso.grid_remove()
                    self._actualizar_uso()
                    try:
                        callback(resultado, exc)
                    except Exception as err:   # noqa: BLE001
                        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{err}")
                    self._actualizar_botones()
        except queue.Empty:
            pass
        try:
            self.after(100, self._procesar_cola)
        except tk.TclError:
            pass

    def _on_cancelar(self) -> None:
        self._cancelar.set()
        self._set_estado("Cancelando… (termina la petición en curso)", "aviso")
        self.btn_cancelar.config(state="disabled")

    # ------------------------------------------------------------- búsqueda
    def _on_buscar(self) -> None:
        if self._ocupado:
            return
        self._guardar_nota_pendiente()
        limite = self._limite_pagina()
        if limite >= PAGINA_TODOS:
            tope = TOPE_PETICIONES_MAX
            if not messagebox.askyesno(
                "Cargar todos los resultados",
                f"Con «todos» la búsqueda puede usar hasta {tope} peticiones "
                f"(≈ ${tope * COSTO_ESTIMADO_POR_PETICION_USD:.2f} USD estimados).\n\n¿Continuar?",
            ):
                return
        try:
            cliente = ClienteGooglePlaces(_leer_google_api_key(), self._uso, self._cancelar)
        except ValueError as exc:
            messagebox.showerror("Falta la API key", str(exc))
            return
        pais, ciudad, tipo = self.var_pais.get(), self.var_ciudad.get(), self.var_tipo.get()
        sesion = SesionBusqueda(
            cliente, self._db, pais, ciudad, tipo, self._tipo_terminos(tipo),
            self._min_resenas(), limite, progreso=lambda txt: self._cola.put(("progreso", txt)),
        )
        self._lanzar(
            sesion.construir_pagina,
            lambda res, exc: self._fin_pagina(sesion, True, res, exc),
            f"Consultando Google Maps — {ciudad}: {tipo}…",
        )

    def _pagina_siguiente(self) -> None:
        s = self._sesion
        if s is None or self._ocupado:
            return
        siguiente = self._pagina_actual + 1
        if siguiente in s.paginas:
            self._mostrar_pagina(siguiente)
        elif s.hay_mas():
            self._lanzar(
                s.construir_pagina,
                lambda res, exc: self._fin_pagina(s, False, res, exc),
                f"Consultando Google Maps (página {siguiente})…",
            )

    def _pagina_anterior(self) -> None:
        if self._sesion is not None and not self._ocupado and self._pagina_actual > 1:
            self._mostrar_pagina(self._pagina_actual - 1)

    def _fin_pagina(self, sesion: SesionBusqueda, nueva: bool, pagina: object, exc: Optional[BaseException]) -> None:
        if exc is not None:
            messagebox.showerror("Error", f"No se pudo completar la búsqueda:\n{exc}")
            self._set_estado("Error en la búsqueda", "error")
            return
        if nueva:
            self._sesion = sesion
            self._pagina_actual = 0
        if not pagina:
            self._render_busqueda()
            if nueva:
                self._set_estado("Sin resultados para esta búsqueda", "error")
            else:
                messagebox.showinfo("Sin más resultados", "Google Maps no devolvió más negocios nuevos para esta búsqueda.")
            if sesion.error:
                messagebox.showwarning("Búsqueda incompleta", sesion.error)
            return
        self._mostrar_pagina(max(sesion.paginas))
        if sesion.error:
            messagebox.showwarning(
                "Búsqueda incompleta",
                f"Google devolvió un error; se muestran los resultados que ya estaban cargados.\n\n{sesion.error}",
            )
        if sesion.cancelado:
            self._set_estado(f"Búsqueda cancelada — se muestran {len(pagina)} resultados ya cargados", "aviso")

    def _on_cambio_reaplicar(self, *_: object) -> None:
        """Cambiar el mínimo de reseñas o el tamaño de página reparte de nuevo lo que ya está en caché."""
        s = self._sesion
        if s is None or self._ocupado:
            return
        s.reconfigurar(self._min_resenas(), self._limite_pagina())
        pagina = s.construir_pagina(permitir_api=False)
        self._pagina_actual = 0
        if pagina:
            self._mostrar_pagina(1)
            extra = " · pulsa «Siguiente» para pedir más a Google" if s.hay_mas() and not s.agotado else ""
            self._set_estado(f"Reaplicado sobre {len(s.raw)} negocios en caché (sin consultar Google){extra}", "ok")
        else:
            self._render_busqueda()
        self._actualizar_botones()

    def _mostrar_pagina(self, numero: int) -> None:
        s = self._sesion
        if s is None or numero not in s.paginas:
            return
        self._pagina_actual = numero
        self._db.sincronizar(s.paginas[numero])
        self._suprimir_filtro = True
        self.var_filtro_b.set("")
        self._suprimir_filtro = False
        self._render_busqueda()
        omit = f" · {len(s.omitidos)} términos omitidos" if s.omitidos else ""
        self._set_estado(
            f"Página {numero}: {len(s.paginas[numero])} leads · {s.ciudad} — {s.tipo} · {s.peticiones} peticiones{omit}",
            "ok",
        )
        self._actualizar_botones()

    def _negocios_pagina(self) -> list[Negocio]:
        s = self._sesion
        return s.paginas.get(self._pagina_actual, []) if s is not None else []

    def _on_filtro_b(self) -> None:
        if not self._suprimir_filtro:
            self._render_busqueda()

    def _render_busqueda(self) -> None:
        termino = self.var_filtro_b.get().strip().lower()
        base = self._negocios_pagina()
        self._visibles = [n for n in base if not termino or termino in _texto_busqueda(n)]
        self._por_iid = {n.place_id: n for n in self._visibles}
        self.tabla_b.set_filas([(n.place_id, _valores_busqueda(n), n.estado) for n in self._visibles])
        self._actualizar_contador_b()
        self._actualizar_botones()

    # -------------------------------------------------- guardados / estados / edición
    def _tabla(self, origen: str) -> TablaDatos:
        return self.tabla_b if origen == "busqueda" else self.tabla_g

    def _mapa(self, origen: str) -> dict[str, Negocio]:
        return self._por_iid if origen == "busqueda" else self._guardados

    def _seleccionados(self, origen: str) -> list[Negocio]:
        mapa = self._mapa(origen)
        return [mapa[i] for i in self._tabla(origen).seleccion() if i in mapa]

    def _propagar(self, n: Negocio) -> None:
        """Copia los datos del usuario a las demás copias en memoria del mismo prospecto y repinta."""
        copias: list[Negocio] = []
        if self._sesion is not None and n.place_id in self._sesion.raw:
            copias.append(self._sesion.raw[n.place_id])
        if n.place_id in self._guardados:
            copias.append(self._guardados[n.place_id])
        for o in copias:
            if o is n:
                continue
            o.estado, o.notas, o.guardado, o.guardado_en = n.estado, n.notas, n.guardado, n.guardado_en
            for campo in CAMPOS_EDITABLES:
                setattr(o, campo, getattr(n, campo))
        if n.place_id in self._por_iid:
            o = self._por_iid[n.place_id]
            self.tabla_b.actualizar_fila(n.place_id, _valores_busqueda(o), o.estado)
        if n.place_id in self._guardados:
            o = self._guardados[n.place_id]
            self.tabla_g.actualizar_fila(n.place_id, _valores_guardados(o), o.estado)

    def _toggle_guardado(self, n: Optional[Negocio]) -> None:
        if n is None:
            return
        nuevo = not n.guardado
        guardado_en = self._db.set_guardado(n.place_id, nuevo)
        if guardado_en is None:
            self._set_estado("No se pudo guardar: el prospecto no está en la base de datos", "error")
            return
        n.guardado, n.guardado_en = nuevo, guardado_en
        self._propagar(n)
        self._refrescar_guardados()
        self._set_estado(f"«{n.nombre}» {'guardado en ★ Guardados' if nuevo else 'quitado de Guardados'}", "ok")

    def _guardar_seleccion(self, origen: str) -> None:
        sel = self._seleccionados(origen)
        if not sel:
            return
        guardar = any(not n.guardado for n in sel)
        cambiados = 0
        for n in sel:
            if n.guardado != guardar:
                ge = self._db.set_guardado(n.place_id, guardar)
                if ge is not None:
                    n.guardado, n.guardado_en = guardar, ge
                    self._propagar(n)
                    cambiados += 1
        self._refrescar_guardados()
        self._set_estado(f"{cambiados} prospecto(s) {'guardados' if guardar else 'quitados de Guardados'}", "ok")

    def _quitar_guardados(self, negocios: Optional[list[Negocio]] = None) -> None:
        sel = negocios if negocios is not None else self._seleccionados("guardados")
        if not sel:
            return
        con_datos = [n for n in sel if n.notas.strip() or n.estado != "nuevo"]
        if con_datos and not messagebox.askyesno(
            "Quitar de guardados",
            f"{len(con_datos)} de los prospectos tienen notas o un estado distinto de «nuevo».\n"
            "Se conservan en la base de datos y reaparecerán si los vuelves a guardar.\n\n¿Quitarlos de Guardados?",
        ):
            return
        self._guardar_nota_pendiente()
        for n in sel:
            if self._db.set_guardado(n.place_id, False) is not None:
                n.guardado, n.guardado_en = False, ""
                self._propagar(n)
        self._refrescar_guardados()
        self._set_estado(f"{len(sel)} prospecto(s) quitados de Guardados", "ok")

    def _crear_menu_estado(self, master: tk.Widget, origen: str) -> tk.Menu:
        menu = tk.Menu(master, tearoff=0)
        for estado in ESTADOS:
            menu.add_command(label=estado.capitalize(), command=lambda e=estado: self._cambiar_estado(origen, e))
        return menu

    def _menu_estado_en(self, destino: tuple[str, int, int]) -> None:
        origen, x, y = destino
        menu = self._crear_menu_estado(self, origen)
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _cambiar_estado(self, origen: str, estado: str) -> None:
        sel = self._seleccionados(origen)
        for n in sel:
            if self._db.set_estado(n.place_id, estado):
                n.estado = estado
                self._propagar(n)
        if sel:
            self._set_estado(f"Estado «{estado}» aplicado a {len(sel)} prospecto(s)", "ok")

    def _aplicar_edicion(self, n: Optional[Negocio], col_id: str, valor: str) -> None:
        if n is None or col_id not in CAMPO_DE_COLUMNA:
            return
        campo = CAMPO_DE_COLUMNA[col_id]
        valor = valor.strip()
        if campo == "resenas":
            if not valor.isdigit():
                self._set_estado("Reseñas debe ser un número entero; no se cambió", "error")
                self._propagar(n)
                return
            nuevo: object = int(valor)
        elif campo == "nombre" and not valor:
            self._set_estado("El nombre no puede quedar vacío; no se cambió", "error")
            self._propagar(n)
            return
        else:
            nuevo = valor
        setattr(n, campo, nuevo)
        self._db.set_edicion(n, campo, nuevo)
        self._propagar(n)
        self._set_estado("Cambio guardado (se conserva aunque Google actualice el prospecto)", "ok")

    def _refrescar_guardados(self) -> None:
        self._guardar_nota_pendiente()
        termino = self.var_filtro_g.get().strip().lower() if hasattr(self, "var_filtro_g") else ""
        todos = self._db.listar_guardados()
        self._guardados = {n.place_id: n for n in todos if not termino or termino in _texto_busqueda(n)}
        self.tabla_g.set_filas([(i, _valores_guardados(n), n.estado) for i, n in self._guardados.items()])
        self.nb.tab(self.tab_g, text=f"★ Guardados ({len(todos)})")
        self.lbl_leads_g.config(
            text=f"{len(self._guardados)} guardados" + (f" de {len(todos)}" if len(todos) != len(self._guardados) else ""),
            foreground=COLORES["ok"] if todos else COLORES["suave"],
        )
        self._nota_id = None
        self._on_sel_guardados()

    def _on_cambio_pestana(self, _event: Optional[tk.Event] = None) -> None:
        self._guardar_nota_pendiente()
        if self.nb.index(self.nb.select()) == 1:
            self._refrescar_guardados()
        else:
            self._db.sincronizar(self._negocios_pagina())
            for n in self._visibles:
                self.tabla_b.actualizar_fila(n.place_id, _valores_busqueda(n), n.estado)

    # --- notas ---
    def _on_sel_guardados(self) -> None:
        self._guardar_nota_pendiente()
        sel = self.tabla_g.seleccion()
        self.txt_nota.config(state="normal")
        self.txt_nota.delete("1.0", "end")
        if len(sel) == 1 and sel[0] in self._guardados:
            n = self._guardados[sel[0]]
            self._nota_id = n.place_id
            self.txt_nota.insert("1.0", n.notas)
            self.lbl_nota.config(text=f"Notas — {n.nombre}")
            self.btn_nota.config(state="normal")
        else:
            self._nota_id = None
            self.txt_nota.config(state="disabled")
            self.lbl_nota.config(text="Notas — selecciona un prospecto" if not sel else "Notas — selecciona un solo prospecto")
            self.btn_nota.config(state="disabled")
        self._actualizar_botones()

    def _guardar_nota_pendiente(self) -> None:
        if not self._nota_id or not hasattr(self, "txt_nota"):
            return
        n = self._guardados.get(self._nota_id)
        if n is None:
            return
        texto = self.txt_nota.get("1.0", "end-1c")
        if texto != n.notas:
            if self._db.set_notas(n.place_id, texto):
                n.notas = texto
                self._propagar(n)
                self._set_estado("Nota guardada", "ok")

    # ---------------------------------------------------- menú contextual y portapapeles
    def _menu_contexto(self, origen: str, event: tk.Event) -> None:
        sel = self._seleccionados(origen)
        if not sel:
            return
        n = sel[0]
        menu = tk.Menu(self, tearoff=0)
        todos_guardados = all(x.guardado for x in sel)
        etiqueta = f" ({len(sel)})" if len(sel) > 1 else ""
        menu.add_command(
            label=("☆ Quitar de guardados" if todos_guardados else "★ Guardar") + etiqueta,
            command=(lambda: self._quitar_guardados()) if (origen == "guardados") else (lambda: self._guardar_seleccion(origen)),
        )
        menu.add_cascade(label="Cambiar estado", menu=self._crear_menu_estado(menu, origen))
        menu.add_separator()
        menu.add_command(label="Copiar celda", command=lambda: self._tabla(origen)._copiar())
        menu.add_command(label="Copiar teléfono" + etiqueta, command=lambda: self._copiar_telefonos(origen))
        if origen == "busqueda":
            menu.add_command(label="Copiar mensaje", command=self._on_copiar_mensaje)
        menu.add_separator()
        menu.add_command(label="Abrir en Google Maps", command=lambda: self._abrir(n.link_maps, "enlace de Maps"),
                         state="normal" if n.link_maps else "disabled")
        menu.add_command(label="Abrir sitio web", command=lambda: self._abrir(n.web, "sitio web"),
                         state="normal" if n.web.lower().startswith(("http://", "https://")) else "disabled")
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _abrir(self, url: str, nombre: str) -> None:
        if not abrir_url(url):
            self._set_estado(f"No hay un {nombre} válido para abrir", "error")

    def _copiar(self, texto: str, mensaje: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(texto)
        self._set_estado(mensaje, "ok")

    def _copiar_telefonos(self, origen: str) -> None:
        tels = [n.telefono for n in self._seleccionados(origen) if n.telefono.strip()]
        if tels:
            self._copiar("\n".join(tels), f"{len(tels)} teléfono(s) copiado(s)")
        else:
            self._set_estado("Los prospectos seleccionados no tienen teléfono", "aviso")

    def _on_copiar_mensaje(self) -> None:
        sel = self._seleccionados("busqueda") or self._seleccionados("guardados")
        if sel:
            self._copiar(mensaje_para(sel[0]), "Mensaje copiado al portapapeles")

    # ------------------------------------------------------------------ exportación
    def _lista_export(self, origen: str, modo: str) -> list[Negocio]:
        if origen == "busqueda":
            if modo == "todo":
                return self._sesion.todos_los_negocios() if self._sesion else []
            mapa = self._por_iid
        else:
            if modo == "todo":
                return list(self._db.listar_guardados())
            mapa = self._guardados
        tabla = self._tabla(origen)
        ids = tabla.iids() if modo == "filtrado" else tabla.seleccion()
        return [mapa[i] for i in ids if i in mapa]

    def _exportar(self, origen: str, modo: str, fmt: str) -> None:
        if self._ocupado:
            return
        lista = self._lista_export(origen, modo)
        if not lista:
            messagebox.showinfo("Exportar", "No hay prospectos para exportar con esa opción.")
            return
        self._escribir_export(lista, origen, fmt)

    def _escribir_export(self, lista: list[Negocio], origen: str, fmt: str) -> None:
        self._guardar_nota_pendiente()
        self._db.sincronizar(lista)
        if origen == "busqueda" and self._sesion is not None:
            s = self._sesion
            leads = "todos" if s.limite >= PAGINA_TODOS else str(len(lista))
            partes = [s.pais, s.ciudad, leads, s.tipo]
        else:
            partes = ["Guardados", str(len(lista)), "prospectos"]
        try:
            os.makedirs(INFORMES_DIR, exist_ok=True)
            ruta = ruta_unica(os.path.join(INFORMES_DIR, nombre_exportacion(partes, fmt)))
            self.config(cursor="watch")
            self.update_idletasks()
            (exportar_excel if fmt == "xlsx" else exportar_csv)(ruta, lista)
        except OSError as exc:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{exc}")
            self._set_estado("Error al exportar", "error")
            return
        finally:
            self.config(cursor="")
        self._set_estado(f"Exportado: {ruta}", "ok")
        if messagebox.askyesno("Exportar", f"Informe guardado ({len(lista)} prospectos):\n{ruta}\n\n¿Abrir la carpeta?"):
            abrir_ruta(INFORMES_DIR)

    def _on_cargar_todo(self) -> None:
        s = self._sesion
        if s is None or self._ocupado:
            messagebox.showinfo("Exportar", "Primero haz una búsqueda.")
            return
        faltan = s.estimar_peticiones_restantes()
        if faltan > 0:
            usar = min(faltan, TOPE_CARGAR_TODO)
            if not messagebox.askyesno(
                "Cargar todo y exportar",
                f"Para cargar todo el resultado faltan hasta {faltan} peticiones a Google "
                f"(≈ ${faltan * COSTO_ESTIMADO_POR_PETICION_USD:.2f} USD).\n"
                f"Esta operación usará como máximo {usar} (tope {TOPE_CARGAR_TODO}).\n\n¿Continuar?",
            ):
                return
        self._lanzar(
            lambda: s.cargar_todo(TOPE_CARGAR_TODO if faltan > 0 else 0),
            lambda res, exc: self._fin_cargar_todo(s, exc),
            "Cargando todos los resultados…",
        )

    def _fin_cargar_todo(self, s: SesionBusqueda, exc: Optional[BaseException]) -> None:
        if exc is not None:
            messagebox.showerror("Error", f"No se pudo completar la carga:\n{exc}")
            return
        if self._pagina_actual in s.paginas:
            self._render_busqueda()
        lista = s.todos_los_negocios()
        if s.error:
            messagebox.showwarning("Carga incompleta", f"Google devolvió un error; se exportará lo ya cargado.\n\n{s.error}")
        elif s.cancelado:
            self._set_estado("Carga cancelada — se exporta lo ya cargado", "aviso")
        elif not s.agotado:
            messagebox.showinfo("Tope alcanzado", f"Se alcanzó el tope de {TOPE_CARGAR_TODO} peticiones; se exporta lo cargado hasta ahora.")
        if lista:
            self._escribir_export(lista, "busqueda", "xlsx")
        else:
            messagebox.showinfo("Exportar", "No hay resultados para exportar.")

    # --------------------------------------------------------------- diálogos
    def _centrar_sobre(self, ventana: tk.Toplevel) -> None:
        ventana.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - ventana.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - ventana.winfo_height()) // 2
        ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _mostrar_info_api(self) -> None:
        ventana = tk.Toplevel(self)
        ventana.title("Info — Uso de Google Places API")
        ventana.geometry("560x640")
        ventana.transient(self)
        ventana.configure(background=COLORES["fondo"])
        marco = ttk.Frame(ventana, padding=16)
        marco.pack(fill="both", expand=True)
        texto = tk.Text(marco, wrap="word", font=("Consolas", 10), relief="solid", bd=1, bg="#FFFFFF", padx=10, pady=10)
        texto.pack(fill="both", expand=True)

        def actualizar() -> None:
            texto.config(state="normal")
            texto.delete("1.0", tk.END)
            texto.insert("1.0", self._uso.resumen_texto() + f"\n\nBase de datos: {self._db.ruta}")
            texto.config(state="disabled")

        actualizar()
        barra = ttk.Frame(marco)
        barra.pack(fill="x", pady=(12, 0))
        ttk.Button(barra, text="Actualizar", command=actualizar).pack(side="left")
        ttk.Button(barra, text="Cerrar", command=ventana.destroy).pack(side="right")
        self._centrar_sobre(ventana)

    def _abrir_editor_palabras(self) -> None:
        if self._editor_palabras is not None and self._editor_palabras.winfo_exists():
            self._editor_palabras.lift()
            return
        v = tk.Toplevel(self)
        self._editor_palabras = v
        v.title("Palabras clave por tipo de negocio")
        v.geometry("560x620")
        v.transient(self)
        v.configure(background=COLORES["fondo"])
        marco = ttk.Frame(v, padding=16)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(2, weight=1)

        fila = ttk.Frame(marco)
        fila.grid(row=0, column=0, sticky="ew")
        fila.columnconfigure(1, weight=1)
        ttk.Label(fila, text="Tipo de negocio").grid(row=0, column=0, padx=(0, 8))
        var_t = tk.StringVar(value=self.var_tipo.get())
        combo = ttk.Combobox(fila, textvariable=var_t, values=list(self._palabras), state="readonly")
        combo.grid(row=0, column=1, sticky="ew")
        ttk.Label(marco, text="Un término por línea. Cada término hace 1 consulta («término en ciudad, país»).",
                  style="Suave.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 4))
        txt = tk.Text(marco, wrap="word", relief="solid", bd=1, padx=8, pady=8)
        txt.grid(row=2, column=0, sticky="nsew")
        lbl = ttk.Label(marco, text="", style="Suave.TLabel")
        lbl.grid(row=3, column=0, sticky="w", pady=(6, 0))

        def cargar(*_: object) -> None:
            txt.delete("1.0", "end")
            txt.insert("1.0", "\n".join(self._palabras.get(var_t.get(), [])))
            btn_elim.config(state="disabled" if var_t.get() in PALABRAS_CLAVE_DEFECTO else "normal")
            lbl.config(text=f"{len(self._palabras.get(var_t.get(), []))} términos guardados")

        def refrescar_combos() -> None:
            combo["values"] = list(self._palabras)
            self.combo_tipo["values"] = list(self._palabras)
            if self.var_tipo.get() not in self._palabras:
                self.var_tipo.set(next(iter(self._palabras)))
            self._actualizar_resumen_terminos()

        def persistir() -> bool:
            try:
                guardar_palabras_clave(self._ruta_palabras, self._palabras, self._json_invalido)
                self._json_invalido = False
                return True
            except OSError as exc:
                messagebox.showerror("Error", f"No se pudo guardar palabras_clave.json:\n{exc}", parent=v)
                return False

        def guardar() -> None:
            terminos = normalizar_terminos(txt.get("1.0", "end").splitlines())
            if not terminos:
                messagebox.showwarning("Palabras clave", "Escribe al menos un término.", parent=v)
                return
            self._palabras[var_t.get()] = terminos
            if persistir():
                refrescar_combos()
                cargar()
                lbl.config(text=f"Guardado: {len(terminos)} términos")

        def restaurar() -> None:
            defecto = PALABRAS_CLAVE_DEFECTO.get(var_t.get())
            if defecto:
                txt.delete("1.0", "end")
                txt.insert("1.0", "\n".join(defecto))
                lbl.config(text="Valores por defecto cargados — pulsa «Guardar» para aplicarlos")

        def nuevo() -> None:
            nombre = simpledialog.askstring("Nuevo tipo", "Nombre del nuevo tipo de negocio:", parent=v)
            nombre = re.sub(r"\s+", " ", nombre or "").strip()
            if not nombre:
                return
            if nombre in self._palabras:
                var_t.set(nombre)
                cargar()
                return
            self._palabras[nombre] = [nombre.lower()]
            if persistir():
                refrescar_combos()
                var_t.set(nombre)
                cargar()

        def eliminar() -> None:
            tipo = var_t.get()
            if tipo in PALABRAS_CLAVE_DEFECTO:
                return
            if messagebox.askyesno("Eliminar tipo", f"¿Eliminar el tipo «{tipo}» y sus términos?", parent=v):
                self._palabras.pop(tipo, None)
                if persistir():
                    refrescar_combos()
                    var_t.set(next(iter(self._palabras)))
                    cargar()

        def reiniciar_omitidos() -> None:
            self._db.reiniciar_terminos()
            lbl.config(text="Listo: se volverán a consultar los términos que se habían omitido")

        barra = ttk.Frame(marco)
        barra.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(barra, text="Guardar", style="Primary.TButton", command=guardar).pack(side="left")
        ttk.Button(barra, text="Valores por defecto", command=restaurar).pack(side="left", padx=6)
        ttk.Button(barra, text="Nuevo tipo…", command=nuevo).pack(side="left")
        btn_elim = ttk.Button(barra, text="Eliminar tipo", command=eliminar)
        btn_elim.pack(side="left", padx=6)
        ttk.Button(barra, text="Cerrar", command=v.destroy).pack(side="right")
        ttk.Button(marco, text="Volver a consultar términos omitidos (aportaban 0 nuevos)", command=reiniciar_omitidos).grid(
            row=5, column=0, sticky="w", pady=(10, 0))
        combo.bind("<<ComboboxSelected>>", cargar)
        cargar()
        self._centrar_sobre(v)

    # ------------------------------------------------------------------ cierre
    def _cerrar(self) -> None:
        try:
            self._guardar_nota_pendiente()
            self.tabla_b.cerrar_editor(True)
            self.tabla_g.cerrar_editor(True)
        finally:
            self._cancelar.set()
            self._db.cerrar()
            self.destroy()


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    app = ProspeccionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
