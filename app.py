import threading
import os
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import db
import importer
import exporter

import sys as _sys
_BASE = os.path.dirname(_sys.executable) if getattr(_sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(_BASE, "Cartas YUGIOH.xlsx")
TIPOS = ["Monstruos", "Rituales", "Fusion", "Sincronia", "Enlace", "XYZ", "Pendulo", "Magia", "Trampas"]

TIPO_COLORS = {
    "Monstruos": "#e67e22",
    "Rituales":  "#9b59b6",
    "Fusion":    "#8e44ad",
    "Sincronia": "#bdc3c7",
    "Enlace":    "#3498db",
    "XYZ":       "#2c3e50",
    "Pendulo":   "#27ae60",
    "Magia":     "#2ecc71",
    "Trampas":   "#e74c3c",
}

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def apply_treeview_style():
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview",
        background="#1a1a30", foreground="white",
        rowheight=30, fieldbackground="#1a1a30", font=("Segoe UI", 12))
    style.configure("Treeview.Heading",
        background="#0d0d1f", foreground="#f0c040",
        font=("Segoe UI", 12, "bold"), relief="flat")
    style.map("Treeview",
        background=[("selected", "#1f6aa5")],
        foreground=[("selected", "white")])


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Biblioteca Yugioh")
        self.geometry("1350x800")
        self.minsize(1100, 640)
        apply_treeview_style()
        db.init_db()
        self._build_ui()

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Sidebar ──────────────────────────────────────────────────────────
        sb = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color="#0d0d1f")
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(2, weight=1)
        sb.grid_columnconfigure(0, weight=1)

        # Logo
        self._build_logo(sb).grid(row=0, column=0, pady=(10, 0), sticky="ew")

        # Separador
        self._sep(sb).grid(row=1, column=0, pady=(2, 4), sticky="ew")

        # Área nav scrollable
        nav = ctk.CTkScrollableFrame(sb, fg_color="transparent",
                                     scrollbar_button_color="#f0c040",
                                     scrollbar_button_hover_color="#c8900a")
        nav.grid(row=2, column=0, sticky="nsew", padx=4)
        nav.grid_columnconfigure(0, weight=1)

        self._nav_btns = {}
        self._cartas_expanded = False

        # ── Buscador global ──
        search_frame = ctk.CTkFrame(nav, fg_color="transparent")
        search_frame.pack(fill="x", pady=(2, 4))
        search_frame.grid_columnconfigure(0, weight=1)
        self._global_search_var = ctk.StringVar()
        gsearch_entry = ctk.CTkEntry(
            search_frame, textvariable=self._global_search_var,
            placeholder_text="🔍 Buscar…", font=ctk.CTkFont(size=12))
        gsearch_entry.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        gsearch_entry.bind("<Return>", lambda _: self._do_global_search())
        ctk.CTkButton(
            search_frame, text="▶", width=28,
            command=self._do_global_search,
            font=ctk.CTkFont(size=12), corner_radius=6
        ).grid(row=0, column=1)

        # ── Cartas (desplegable) ──
        self._cartas_toggle = ctk.CTkButton(
            nav, text="🃏  Cartas  ▶", anchor="w",
            font=ctk.CTkFont(size=13), corner_radius=6,
            command=self._toggle_cartas)
        self._cartas_toggle.pack(fill="x", pady=(2, 0))
        self._nav_btns["cartas"] = self._cartas_toggle

        # Sub-frame de tipos (oculto inicialmente, no empaquetado)
        self._tipos_sub = ctk.CTkFrame(nav, fg_color="#08081a", corner_radius=6)
        tipo_icons = {
            "Todos":"▪","Monstruos":"🟠","Rituales":"🟣","Fusion":"🟤",
            "Sincronia":"⚪","Enlace":"🔵","XYZ":"⬛","Pendulo":"🟩",
            "Magia":"💚","Trampas":"🔴",
        }
        self._tipo_nav_btns = {}
        for tipo in ["Todos"] + TIPOS:
            icon = tipo_icons.get(tipo, "▪")
            b = ctk.CTkButton(
                self._tipos_sub, text=f"   {icon}  {tipo}", anchor="w",
                fg_color="transparent", hover_color="#1f3060",
                font=ctk.CTkFont(size=12), corner_radius=4, height=28,
                command=lambda t=tipo: self._nav_to_tipo(t))
            b.pack(fill="x", padx=4, pady=1)
            self._tipo_nav_btns[tipo] = b

        # ── Resto de secciones ──
        other_items = [
            ("♻  Repetidos",    "repetidos"),
            ("⚔  Decks",        "decks"),
            ("📋  Pendientes",   "pendientes"),
            ("🏛  Álbumes",      "albumes"),
            ("📊  Estadísticas", "stats"),
            ("📈  Gráficas",     "graficas"),
        ]
        for label, key in other_items:
            b = ctk.CTkButton(nav, text=label, anchor="w",
                              font=ctk.CTkFont(size=13), corner_radius=6,
                              command=lambda k=key: self._show(k))
            b.pack(fill="x", pady=3)
            self._nav_btns[key] = b

        # Separador inferior
        self._sep(sb).grid(row=3, column=0, pady=4, sticky="ew")

        # Botones Excel
        bot = ctk.CTkFrame(sb, fg_color="transparent")
        bot.grid(row=4, column=0, sticky="ew", padx=6, pady=(0, 8))
        bot.grid_columnconfigure(0, weight=1)
        for i, (txt, color, hov, cmd) in enumerate([
            ("⟳  Reimportar", "#1e2a3a", "#2a3d55", self._reimport),
            ("↓  Exportar",   "#1a3a1a", "#2d6a4f", self._exportar_excel),
            ("↑  Actualizar", "#3a1a3a", "#6d2b69", self._actualizar_excel),
            ("📂  Importar BD","#2a1a0a", "#5a3a10", self._importar_bd),
        ]):
            ctk.CTkButton(bot, text=txt, fg_color=color, hover_color=hov,
                          font=ctk.CTkFont(size=12), corner_radius=6,
                          command=cmd).grid(row=i, column=0, pady=2, sticky="ew")

        # ── Área principal ────────────────────────────────────────────────────
        self._main = ctk.CTkFrame(self, fg_color="#12122a")
        self._main.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self._main.grid_columnconfigure(0, weight=1)
        self._main.grid_rowconfigure(0, weight=1)

        self._dlg_ubicacion_widget = None
        self._dlg_copia_widget = None
        self._dlg_copia_var = None

        self._frames = {}
        self._build_cartas()
        self._build_repetidos()
        self._build_decks()
        self._build_pendientes()
        self._build_albumes()
        self._build_stats()
        self._build_graficas()
        self._build_global_search_frame()
        self._show("cartas")

    def _sep(self, parent):
        c = tk.Canvas(parent, width=180, height=2, bg="#0d0d1f", highlightthickness=0)
        c.create_line(8, 1, 172, 1, fill="#f0c040", width=1)
        return c

    def _build_logo(self, parent):
        import math
        gold   = "#f0c040"
        gold2  = "#c8900a"
        bg     = "#0d0d1f"
        cx, cy = 92, 90
        w, h   = 185, 145

        c = tk.Canvas(parent, width=w, height=h, bg=bg, highlightthickness=0)

        # Título
        c.create_text(cx, 14, text="B I B L I O T E C A",
                      fill=gold, font=("Segoe UI", 9, "bold"))
        c.create_text(cx, 28, text="Y U G I O H",
                      fill=gold2, font=("Segoe UI", 11, "bold"))

        # Anillos exteriores decorativos
        for r, col, w_ in [(50, "#2a2000", 1), (47, "#5a4500", 1), (44, gold, 2)]:
            c.create_oval(cx - r, cy - r, cx + r, cy + r,
                          outline=col, fill="", width=w_)

        # Puntos en el anillo
        for deg in range(0, 360, 40):
            rad = math.radians(deg)
            px = cx + 44 * math.cos(rad)
            py = cy + 44 * math.sin(rad)
            c.create_oval(px - 2.5, py - 2.5, px + 2.5, py + 2.5, fill=gold, outline="")

        # Triángulo del Puzzle del Milenio
        r_tri = 32
        pts = []
        for deg in [270, 30, 150]:
            rad = math.radians(deg)
            pts += [cx + r_tri * math.cos(rad), cy + r_tri * math.sin(rad)]
        c.create_polygon(pts, outline=gold, fill="#08081a", width=2)

        # Ojo (dentro del triángulo)
        ey = cy + 4
        # Blanco del ojo
        c.create_oval(cx - 14, ey - 7, cx + 14, ey + 7,
                      outline=gold, fill="#08081a", width=1.5)
        # Iris
        c.create_oval(cx - 6, ey - 6, cx + 6, ey + 6, fill=gold2, outline="")
        # Pupila
        c.create_oval(cx - 3, ey - 3, cx + 3, ey + 3, fill="#000", outline="")
        # Cola del ojo de Horus (abajo-derecha)
        c.create_line(cx + 10, ey + 5, cx + 20, ey + 14,
                      fill=gold, width=2, capstyle="round")
        c.create_line(cx + 20, ey + 14, cx + 16, ey + 20,
                      fill=gold, width=2, capstyle="round")
        # Ceja (arriba-izquierda)
        c.create_line(cx - 14, ey - 7, cx - 20, ey - 17,
                      fill=gold, width=2, capstyle="round")

        # Cuatro líneas diagonales decorativas (marco tipo carta)
        for deg in [45, 135, 225, 315]:
            rad = math.radians(deg)
            x1 = cx + 44 * math.cos(rad)
            y1 = cy + 44 * math.sin(rad)
            x2 = cx + 52 * math.cos(rad)
            y2 = cy + 52 * math.sin(rad)
            c.create_line(x1, y1, x2, y2, fill=gold2, width=1.5)

        return c

        # Main area
        self._main = ctk.CTkFrame(self, fg_color="#1e1e2e")
        self._main.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self._main.grid_columnconfigure(0, weight=1)
        self._main.grid_rowconfigure(0, weight=1)

        self._frames = {}
        self._build_cartas()
        self._build_repetidos()
        self._build_decks()
        self._build_pendientes()
        self._build_stats()
        self._build_graficas()
        self._show("cartas")

    def _build_global_search_frame(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        f.grid_rowconfigure(3, weight=1)
        self._frames["global_search"] = f

        ctk.CTkLabel(f, text="Cartas en colección",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#f0c040").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 2))

        tf1 = ctk.CTkFrame(f, fg_color="transparent")
        tf1.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        tf1.grid_columnconfigure(0, weight=1)
        tf1.grid_rowconfigure(0, weight=1)

        cols1 = ("tipo", "cant", "nombre", "ubicacion", "ubicacion_copia")
        self._gs_cartas_tree = ttk.Treeview(tf1, columns=cols1, show="headings")
        for col, head, width, anchor in [
            ("tipo",           "Tipo",       100, "w"),
            ("cant",           "Cant",        50, "center"),
            ("nombre",         "Nombre",     300, "w"),
            ("ubicacion",      "Ubicación",  160, "w"),
            ("ubicacion_copia","Copia",      160, "w"),
        ]:
            self._gs_cartas_tree.heading(col, text=head)
            self._gs_cartas_tree.column(col, width=width, anchor=anchor)
        vsb1 = ttk.Scrollbar(tf1, orient="vertical", command=self._gs_cartas_tree.yview)
        self._gs_cartas_tree.configure(yscrollcommand=vsb1.set)
        self._gs_cartas_tree.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")

        ctk.CTkLabel(f, text="Cartas en decks",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#f0c040").grid(row=2, column=0, sticky="w", padx=8, pady=(4, 2))

        tf2 = ctk.CTkFrame(f, fg_color="transparent")
        tf2.grid(row=3, column=0, sticky="nsew", padx=4, pady=(0, 4))
        tf2.grid_columnconfigure(0, weight=1)
        tf2.grid_rowconfigure(0, weight=1)

        cols2 = ("deck", "cant", "nombre", "en_album")
        self._gs_decks_tree = ttk.Treeview(tf2, columns=cols2, show="headings")
        for col, head, width, anchor in [
            ("deck",     "Deck",          150, "w"),
            ("cant",     "Cant",           50, "center"),
            ("nombre",   "Nombre",        300, "w"),
            ("en_album", "Copia en álbum",130, "center"),
        ]:
            self._gs_decks_tree.heading(col, text=head)
            self._gs_decks_tree.column(col, width=width, anchor=anchor)
        vsb2 = ttk.Scrollbar(tf2, orient="vertical", command=self._gs_decks_tree.yview)
        self._gs_decks_tree.configure(yscrollcommand=vsb2.set)
        self._gs_decks_tree.tag_configure("en_album", foreground="#2ecc71")
        self._gs_decks_tree.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")

    def _do_global_search(self):
        query = self._global_search_var.get().strip()
        if not query:
            return
        self._show("global_search")
        self._refresh_global_search(query)

    def _refresh_global_search(self, query=None):
        if query is None:
            query = self._global_search_var.get().strip()

        self._gs_cartas_tree.delete(*self._gs_cartas_tree.get_children())
        for r in db.get_cartas(search=query, mostrar_vacios=False, include_pendientes=False):
            r = dict(r)
            self._gs_cartas_tree.insert("", "end", values=(
                r.get("tipo", ""), r.get("cant", ""),
                r.get("nombre", ""), r.get("ubicacion", ""),
                r.get("ubicacion_copia", ""),
            ))

        self._gs_decks_tree.delete(*self._gs_decks_tree.get_children())
        for r in db.get_decks(search=query):
            r = dict(r)
            en = r.get("repetidos", 0)
            self._gs_decks_tree.insert("", "end", iid=str(r["id"]),
                tags=("en_album",) if en else (),
                values=(r.get("deck_nombre", ""), r["cant"], r["nombre"],
                        "Sí" if en else "No"))

    def _show(self, key):
        for f in self._frames.values():
            f.grid_remove()
        self._frames[key].grid(row=0, column=0, sticky="nsew")
        for k, b in self._nav_btns.items():
            b.configure(fg_color="#1f6aa5" if k == key else "transparent")
        refreshers = {
            "cartas":         self._refresh_cartas,
            "repetidos":      self._refresh_repetidos,
            "decks":          self._refresh_decks,
            "pendientes":     self._refresh_pendientes,
            "albumes":        self._refresh_albumes,
            "stats":          self._refresh_stats,
            "graficas":       self._refresh_graficas,
            "global_search":  self._refresh_global_search,
        }
        refreshers[key]()

    def _toggle_cartas(self):
        if self._cartas_expanded:
            self._tipos_sub.pack_forget()
            self._cartas_toggle.configure(text="🃏  Cartas  ▶")
            self._cartas_expanded = False
        else:
            self._tipos_sub.pack(fill="x", pady=(0, 2),
                                 after=self._cartas_toggle)
            self._cartas_toggle.configure(text="🃏  Cartas  ▼")
            self._cartas_expanded = True
        self._show("cartas")

    def _nav_to_tipo(self, tipo):
        self._tipo_seg.set(tipo)
        # resaltar sub-botón activo
        for t, b in self._tipo_nav_btns.items():
            b.configure(fg_color="#1f3060" if t == tipo else "transparent")
        self._show("cartas")

    # ── Cartas ───────────────────────────────────────────────────────────────

    def _build_cartas(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(3, weight=1)
        self._frames["cartas"] = f

        # Tabs por tipo
        self._tipo_seg = ctk.CTkSegmentedButton(
            f, values=["Todos"] + TIPOS,
            font=ctk.CTkFont(size=14),
            command=lambda _: self._refresh_cartas())
        self._tipo_seg.set("Todos")
        self._tipo_seg.grid(row=0, column=0, sticky="ew", padx=4, pady=(6, 2))

        # Toolbar
        tb = ctk.CTkFrame(f, fg_color="transparent")
        tb.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 0))
        tb.grid_columnconfigure(0, weight=1)

        self._search_var = ctk.StringVar()
        e = ctk.CTkEntry(tb, textvariable=self._search_var,
                         placeholder_text="Buscar carta…")
        e.grid(row=0, column=0, padx=4, sticky="ew")
        e.bind("<Return>", lambda _: self._refresh_cartas())

        ctk.CTkButton(tb, text="Buscar", width=75,
                      command=self._refresh_cartas).grid(row=0, column=1, padx=4)
        ctk.CTkButton(tb, text="Limpiar", width=75,
                      fg_color="gray30",
                      command=self._clear_search).grid(row=0, column=2, padx=4)

        self._vacios_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(tb, text="Mostrar huecos",
                        variable=self._vacios_var,
                        command=self._refresh_cartas).grid(row=0, column=3, padx=8)

        ctk.CTkButton(tb, text="+ Añadir carta", width=110,
                      command=self._add_carta_dialog).grid(row=0, column=4, padx=4)

        # Count
        self._cartas_count = ctk.CTkLabel(f, text="", text_color="gray60",
                                          font=ctk.CTkFont(size=14))
        self._cartas_count.grid(row=2, column=0, sticky="w", padx=8, pady=2)

        # Treeview
        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ("tipo", "cant", "nombre", "ubicacion", "ubicacion_copia", "cant2", "repetidos", "pendiente")
        self._ct = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")

        specs = [
            ("tipo",           "Tipo",          95,  "center"),
            ("cant",           "Cant",           45,  "center"),
            ("nombre",         "Nombre",        310,  "w"),
            ("ubicacion",      "Ubicación",     130,  "w"),
            ("ubicacion_copia","Ubi. Copia",    120,  "w"),
            ("cant2",          "Cant2",          50,  "center"),
            ("repetidos",      "Repetido",       70,  "center"),
            ("pendiente",      "Pendiente",      70,  "center"),
        ]
        for col, head, width, anchor in specs:
            self._ct.heading(col, text=head,
                             command=lambda c=col: self._sort_cartas(c))
            self._ct.column(col, width=width, anchor=anchor, minwidth=40)

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._ct.yview)
        self._ct.configure(yscrollcommand=vsb.set)
        self._ct.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._ct.tag_configure("vacio",     foreground="#555577")
        self._ct.tag_configure("pendiente", foreground="#f39c12")
        self._ct.tag_configure("repetido",  foreground="#e74c3c")

        self._ct.bind("<Double-1>",   self._edit_carta_dialog)
        self._ct.bind("<Delete>",     self._delete_or_clear_carta)
        self._ct.bind("<BackSpace>",  self._delete_or_clear_carta)
        self._ct.bind("<Button-3>",   self._cartas_context_menu)

        self._cartas_sort = {"col": None, "rev": False}
        self._ctx_menu = tk.Menu(self, tearoff=0, bg="#2b2b2b", fg="white",
                                 activebackground="#1f6aa5", activeforeground="white")

    def _clear_search(self):
        self._search_var.set("")
        self._refresh_cartas()

    def _sort_cartas(self, col):
        rev = self._cartas_sort["col"] == col and not self._cartas_sort["rev"]
        self._cartas_sort = {"col": col, "rev": rev}
        items = [(self._ct.set(k, col), k) for k in self._ct.get_children("")]
        items.sort(reverse=rev, key=lambda x: x[0].lower() if x[0] else "")
        for idx, (_, k) in enumerate(items):
            self._ct.move(k, "", idx)

    def _refresh_cartas(self, *_):
        tipo = self._tipo_seg.get()
        search = self._search_var.get().strip()
        mostrar_vacios = self._vacios_var.get()

        if tipo == "Todos":
            self._ct["displaycolumns"] = ("tipo", "cant", "nombre", "ubicacion", "ubicacion_copia", "cant2", "repetidos", "pendiente")
        else:
            self._ct["displaycolumns"] = ("cant", "nombre", "ubicacion", "ubicacion_copia", "cant2", "repetidos", "pendiente")

        rows = db.get_cartas(
            tipo=tipo if tipo != "Todos" else None,
            search=search or None,
            mostrar_vacios=mostrar_vacios,
        )

        self._ct.delete(*self._ct.get_children())
        filled = vacios = 0
        for r in rows:
            r = dict(r)
            is_vacio = r.get("vacio", 0)
            is_pend = r.get("pendiente", 0)
            is_rep = r.get("repetidos", 0)
            c2 = r.get("cant2")
            try:
                cant2_str = str(int(float(c2))) if c2 is not None and str(c2).strip() not in ("", "None") else ""
            except (ValueError, TypeError):
                cant2_str = ""

            if is_vacio:
                vacios += 1
                tags = ("vacio",)
                vals = (r.get("tipo", ""), "", "(Vacío)", "", "", "", "", "")
            else:
                filled += 1
                tags = ("pendiente",) if is_pend else (("repetido",) if is_rep else ())
                vals = (
                    r.get("tipo", "Pendiente") if is_pend else r.get("tipo", ""),
                    r.get("cant", ""),
                    r.get("nombre", ""),
                    r.get("ubicacion", "") or "",
                    r.get("ubicacion_copia", "") or "",
                    cant2_str,
                    "Sí" if is_rep else "No",
                    "Sí" if is_pend else "No",
                )

            iid = f"p_{r['id']}" if is_pend else f"c_{r['id']}_{is_vacio}"
            self._ct.insert("", "end", iid=iid, values=vals, tags=tags)

        txt = f"{filled} cartas"
        if mostrar_vacios:
            txt += f"  |  {vacios} huecos vacíos"
        self._cartas_count.configure(text=txt)

    def _add_carta_dialog(self):
        self._carta_dialog(None)

    def _edit_carta_dialog(self, event=None):
        sel = self._ct.selection()
        if not sel:
            return
        iid = sel[0]
        if iid.startswith("p_") or "_1" in iid:
            return  # pendiente or vacant — open fill dialog for vacant
        if iid.endswith("_1"):
            return
        parts = iid.split("_")
        carta_id = int(parts[1])
        row = dict(db.get_conn().execute(
            "SELECT * FROM cartas WHERE id=?", (carta_id,)
        ).fetchone())
        self._carta_dialog(row)

    def _carta_dialog(self, existing=None):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Editar carta" if existing else "Añadir carta")
        dlg.geometry("420x400")
        dlg.grab_set()
        dlg.resizable(False, False)

        ubicaciones = db.get_ubicaciones()
        deck_names = db.get_deck_names()
        opciones_copia = [""] + ["Caja de copias"] + deck_names + \
                         [f"Caja de copias + {d}" for d in deck_names]

        fields = {}
        self._dlg_ubicacion_widget = None
        self._dlg_copia_widget = None
        self._dlg_copia_var = None

        specs = [
            ("Nombre",     "nombre",          "entry"),
            ("Tipo",       "tipo",            "option"),
            ("Cant",       "cant",            "entry"),
            ("Ubicación",  "ubicacion",       "combo"),
            ("Ubi. Copia", "ubicacion_copia", "copia"),
            ("Cant2",      "cant2",           "entry"),
        ]
        for i, (lbl, key, wtype) in enumerate(specs):
            ctk.CTkLabel(dlg, text=lbl, anchor="w").grid(
                row=i, column=0, padx=14, pady=5, sticky="w")
            v = (existing or {}).get(key)
            if key == "cant2" and v is not None:
                try:
                    v = str(int(float(v)))
                except (ValueError, TypeError):
                    v = ""
            val = "" if v is None else str(v)

            if wtype == "option":
                var = ctk.StringVar(value=val or "Monstruos")
                w = ctk.CTkOptionMenu(dlg, variable=var, values=TIPOS, width=220)
                fields[key] = var
            elif wtype == "combo":
                var = ctk.StringVar(value=val)
                w = ctk.CTkComboBox(dlg, variable=var, values=ubicaciones,
                                    width=220, state="normal")
                w.set(val)
                fields[key] = var
                self._dlg_ubicacion_widget = w
            elif wtype == "copia":
                # Valor actual puede no estar en la lista → añadirlo
                opts = opciones_copia if val in opciones_copia \
                       else opciones_copia + [val]
                var = ctk.StringVar(value=val)
                w = ctk.CTkOptionMenu(dlg, variable=var, values=opts, width=220)
                fields[key] = var
                self._dlg_copia_widget = w
                self._dlg_copia_var = var
            else:
                var = ctk.StringVar(value=val or ("1" if key == "cant" else ""))
                w = ctk.CTkEntry(dlg, textvariable=var, width=220)
                fields[key] = var
            w.grid(row=i, column=1, padx=14, pady=5, sticky="w")

        def _clear_dlg_refs():
            self._dlg_ubicacion_widget = None
            self._dlg_copia_widget = None
            self._dlg_copia_var = None

        dlg.protocol("WM_DELETE_WINDOW", lambda: (_clear_dlg_refs(), dlg.destroy()))

        def save():
            nombre = fields["nombre"].get().strip()
            if not nombre:
                messagebox.showwarning("Error", "Nombre requerido", parent=dlg)
                return

            es_repetido = 0
            if not existing and db.check_repetido(nombre):
                resp = messagebox.askyesno(
                    "Carta ya existe",
                    f"'{nombre}' ya está en el álbum.\n\n"
                    "¿Añadir igualmente como repetida?",
                    parent=dlg)
                if not resp:
                    return
                es_repetido = 1
            elif existing:
                es_repetido = existing.get("repetidos", 0)

            try:
                cant = int(fields["cant"].get() or 1)
            except ValueError:
                cant = 1

            try:
                cant2_raw = fields["cant2"].get().strip()
                cant2 = int(float(cant2_raw)) if cant2_raw else None
            except ValueError:
                cant2 = None

            if existing:
                db.update_carta(existing["id"], cant, nombre,
                                fields["ubicacion"].get() or None,
                                fields["ubicacion_copia"].get() or None,
                                cant2, es_repetido)
            else:
                db.add_carta(fields["tipo"].get(), cant, nombre,
                             fields["ubicacion"].get() or None,
                             fields["ubicacion_copia"].get() or None,
                             cant2, es_repetido)
            _clear_dlg_refs()
            dlg.destroy()
            self._refresh_cartas()

        ctk.CTkButton(dlg, text="Guardar", command=save).grid(
            row=len(specs), column=0, columnspan=2, pady=12)

    def _delete_or_clear_carta(self, event=None):
        sel = self._ct.selection()
        if not sel:
            return
        iid = sel[0]
        if iid.startswith("p_"):
            return

        parts = iid.split("_")
        carta_id = int(parts[1])
        is_vacio = parts[2] == "1" if len(parts) > 2 else False

        if is_vacio:
            messagebox.showinfo("Hueco vacío",
                                "Este hueco está vacío. Doble clic para añadir una carta.")
            return

        db.clear_carta(carta_id)
        self._refresh_cartas()

    def _cartas_context_menu(self, event):
        iid = self._ct.identify_row(event.y)
        if not iid or iid.startswith("p_"):
            return
        parts = iid.split("_")
        is_vacio = parts[2] == "1" if len(parts) > 2 else False
        if is_vacio:
            return
        self._ct.selection_set(iid)

        self._ctx_menu.delete(0, "end")
        self._ctx_menu.add_command(
            label="Editar carta",
            command=lambda: self._edit_carta_dialog())
        self._ctx_menu.add_command(
            label="Mover a deck…",
            command=lambda: self._mover_a_deck_dialog(int(parts[1])))
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(
            label="Dejar hueco vacío",
            command=lambda: (db.clear_carta(int(parts[1])), self._refresh_cartas()))
        self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _mover_a_deck_dialog(self, carta_id):
        nombres_deck = db.get_deck_names()

        dlg = ctk.CTkToplevel(self)
        dlg.title("Mover a deck")
        dlg.geometry("380x240")
        dlg.grab_set()
        dlg.resizable(False, False)

        ctk.CTkLabel(dlg, text="Selecciona o crea un deck:",
                     font=ctk.CTkFont(size=14)).pack(pady=(18, 6))

        deck_var = ctk.StringVar(value=nombres_deck[0] if nombres_deck else "")
        opciones = nombres_deck + ["── Deck nuevo ──"]
        menu = ctk.CTkOptionMenu(dlg, variable=deck_var, values=opciones, width=280)
        menu.pack(pady=4)

        nuevo_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        nuevo_frame.pack(pady=4)
        ctk.CTkLabel(nuevo_frame, text="Nombre nuevo:").grid(row=0, column=0, padx=6)
        nuevo_entry = ctk.CTkEntry(nuevo_frame, width=200)
        nuevo_entry.grid(row=0, column=1, padx=6)

        hueco_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(dlg, text="Dejar hueco vacío en el álbum",
                        variable=hueco_var).pack(pady=6)

        def confirmar():
            seleccion = deck_var.get()
            if seleccion == "── Deck nuevo ──":
                nombre_deck = nuevo_entry.get().strip()
                if not nombre_deck:
                    messagebox.showwarning("Error", "Escribe el nombre del deck", parent=dlg)
                    return
            else:
                nombre_deck = seleccion

            db.move_carta_to_deck(carta_id, nombre_deck, dejar_hueco=hueco_var.get())
            dlg.destroy()
            self._refresh_cartas()
            self._refresh_decks()
            self._refresh_dialog_opts()

        ctk.CTkButton(dlg, text="Mover", command=confirmar).pack(pady=10)

    def _refresh_dialog_opts(self):
        if self._dlg_ubicacion_widget and self._dlg_ubicacion_widget.winfo_exists():
            self._dlg_ubicacion_widget.configure(values=db.get_ubicaciones())
        if self._dlg_copia_widget and self._dlg_copia_widget.winfo_exists():
            deck_names = db.get_deck_names()
            opciones = [""] + ["Caja de copias"] + deck_names + \
                       [f"Caja de copias + {d}" for d in deck_names]
            current = self._dlg_copia_var.get() if self._dlg_copia_var else ""
            if current and current not in opciones:
                opciones = opciones + [current]
            self._dlg_copia_widget.configure(values=opciones)

    # ── Repetidos ─────────────────────────────────────────────────────────────

    def _build_repetidos(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        self._frames["repetidos"] = f

        ctk.CTkLabel(f,
            text="Cartas marcadas como repetidas — ya existían en el álbum al añadirlas.",
            text_color="gray60", font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))

        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ("tipo", "cant", "nombre", "ubicacion", "ubicacion_copia", "cant2")
        self._rept = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for col, head, width, anchor in [
            ("tipo",           "Tipo",       100, "center"),
            ("cant",           "Cant",        45, "center"),
            ("nombre",         "Nombre",     340, "w"),
            ("ubicacion",      "Ubicación",  140, "w"),
            ("ubicacion_copia","Ubi. Copia", 130, "w"),
            ("cant2",          "Cant2",       55, "center"),
        ]:
            self._rept.heading(col, text=head)
            self._rept.column(col, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._rept.yview)
        self._rept.configure(yscrollcommand=vsb.set)
        self._rept.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._rept.tag_configure("rep", foreground="#e74c3c")
        self._rept.bind("<Delete>", self._delete_repetido)

    def _refresh_repetidos(self):
        conn = db.get_conn()
        rows = conn.execute(
            "SELECT * FROM cartas WHERE repetidos=1 AND vacio=0 ORDER BY tipo, nombre"
        ).fetchall()
        conn.close()

        self._rept.delete(*self._rept.get_children())
        for r in rows:
            r = dict(r)
            c2 = r.get("cant2")
            try:
                cant2_str = str(int(float(c2))) if c2 is not None and str(c2).strip() not in ("", "None") else ""
            except (ValueError, TypeError):
                cant2_str = ""
            self._rept.insert("", "end", iid=str(r["id"]), tags=("rep",),
                values=(r["tipo"], r["cant"], r["nombre"],
                        r.get("ubicacion") or "",
                        r.get("ubicacion_copia") or "",
                        cant2_str))

    def _delete_repetido(self, _=None):
        sel = self._rept.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar esta carta repetida?"):
            db.delete_carta(int(sel[0]))
            self._refresh_repetidos()

    # ── Decks ─────────────────────────────────────────────────────────────────

    def _build_decks(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(2, weight=1)
        self._frames["decks"] = f

        # Toolbar
        tb = ctk.CTkFrame(f, fg_color="transparent")
        tb.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        tb.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(tb, text="Deck:").grid(row=0, column=0, padx=(4, 2))
        self._deck_filtro_var = ctk.StringVar(value="Todos")
        self._deck_filtro_menu = ctk.CTkOptionMenu(
            tb, variable=self._deck_filtro_var, values=["Todos"],
            width=160, command=lambda _: self._refresh_decks())
        self._deck_filtro_menu.grid(row=0, column=1, padx=4)

        self._deck_search = ctk.StringVar()
        e = ctk.CTkEntry(tb, textvariable=self._deck_search,
                         placeholder_text="Buscar carta en deck…")
        e.grid(row=0, column=2, padx=4, sticky="ew")
        e.bind("<Return>", lambda _: self._refresh_decks())
        ctk.CTkButton(tb, text="Buscar", width=75,
                      command=self._refresh_decks).grid(row=0, column=3, padx=4)
        ctk.CTkButton(tb, text="Renombrar deck", width=120,
                      fg_color="#1a3a5a", hover_color="#1f6aa5",
                      command=self._renombrar_deck).grid(row=0, column=4, padx=4)
        ctk.CTkButton(tb, text="Eliminar deck", width=110,
                      fg_color="#8b0000", hover_color="#a00000",
                      command=self._eliminar_deck).grid(row=0, column=5, padx=4)

        ctk.CTkLabel(f,
            text="Cartas movidas al deck desde el álbum.  Verde = hay copia en álbum.",
            text_color="gray60", font=ctk.CTkFont(size=14)
        ).grid(row=1, column=0, sticky="w", padx=8, pady=(0, 2))

        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ("deck", "cant", "nombre", "en_album")
        self._dt = ttk.Treeview(tf, columns=cols, show="headings")
        for col, head, width, anchor in [
            ("deck",     "Deck",          150, "w"),
            ("cant",     "Cant",           50, "center"),
            ("nombre",   "Nombre",        380, "w"),
            ("en_album", "Copia en álbum",130, "center"),
        ]:
            self._dt.heading(col, text=head)
            self._dt.column(col, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._dt.yview)
        self._dt.configure(yscrollcommand=vsb.set)
        self._dt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._dt.tag_configure("en_album", foreground="#2ecc71")
        self._dt.bind("<Delete>", self._delete_deck_carta)

    def _refresh_decks(self):
        # Actualiza el selector de decks
        nombres = db.get_deck_names()
        opciones = ["Todos"] + nombres
        self._deck_filtro_menu.configure(values=opciones)
        if self._deck_filtro_var.get() not in opciones:
            self._deck_filtro_var.set("Todos")

        search = self._deck_search.get().strip() or None
        filtro = self._deck_filtro_var.get()
        self._dt.delete(*self._dt.get_children())
        for r in db.get_decks(search=search, deck_nombre=filtro):
            r = dict(r)
            en = r.get("repetidos", 0)
            self._dt.insert("", "end", iid=str(r["id"]),
                tags=("en_album",) if en else (),
                values=(r.get("deck_nombre", ""),
                        r["cant"], r["nombre"],
                        "Sí (copia en álbum)" if en else "No"))

    def _delete_deck_carta(self, _=None):
        sel = self._dt.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", "¿Quitar esta carta del deck?"):
            db.delete_deck_carta(int(sel[0]))
            self._refresh_decks()

    def _renombrar_deck(self):
        filtro = self._deck_filtro_var.get()
        if filtro == "Todos":
            messagebox.showinfo("Info", "Selecciona un deck específico para renombrarlo.")
            return
        dlg = ctk.CTkToplevel(self)
        dlg.title("Renombrar deck")
        dlg.geometry("340x130")
        dlg.grab_set()
        dlg.resizable(False, False)
        ctk.CTkLabel(dlg, text=f"Nuevo nombre para '{filtro}':").pack(pady=(16, 4), padx=16, anchor="w")
        var = ctk.StringVar(value=filtro)
        entry = ctk.CTkEntry(dlg, textvariable=var, width=300)
        entry.pack(padx=16)
        entry.focus()
        entry.select_range(0, "end")

        def confirmar(_=None):
            nuevo = var.get().strip()
            if not nuevo or nuevo == filtro:
                dlg.destroy()
                return
            db.rename_deck(filtro, nuevo)
            self._deck_filtro_var.set(nuevo)
            dlg.destroy()
            self._refresh_decks()

        entry.bind("<Return>", confirmar)
        ctk.CTkButton(dlg, text="Guardar", command=confirmar).pack(pady=10)

    def _eliminar_deck(self):
        filtro = self._deck_filtro_var.get()
        if filtro == "Todos":
            messagebox.showinfo("Info", "Selecciona un deck específico para eliminarlo.")
            return
        if messagebox.askyesno("Eliminar deck",
                               f"¿Eliminar el deck '{filtro}' y todas sus cartas?"):
            db.delete_deck(filtro)
            self._deck_filtro_var.set("Todos")
            self._refresh_decks()

    # ── Pendientes ────────────────────────────────────────────────────────────

    def _build_pendientes(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        self._frames["pendientes"] = f

        form = ctk.CTkFrame(f)
        form.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(form, text="Nombre:").grid(row=0, column=0, padx=8, pady=8)
        self._pend_nombre = ctk.CTkEntry(form, placeholder_text="Nombre de la carta")
        self._pend_nombre.grid(row=0, column=1, padx=4, pady=8, sticky="ew")

        ctk.CTkLabel(form, text="Tipo:").grid(row=0, column=2, padx=8)
        self._pend_tipo = ctk.StringVar(value="Monstruos")
        ctk.CTkOptionMenu(form, variable=self._pend_tipo,
                          values=TIPOS, width=130).grid(row=0, column=3, padx=4)

        ctk.CTkLabel(form, text="Notas:").grid(row=0, column=4, padx=8)
        self._pend_notas = ctk.CTkEntry(form, placeholder_text="Opcional")
        self._pend_notas.grid(row=0, column=5, padx=4, sticky="ew")
        form.grid_columnconfigure(5, weight=1)

        ctk.CTkButton(form, text="+ Añadir", width=90,
                      command=self._add_pendiente).grid(row=0, column=6, padx=10)

        ctk.CTkLabel(f,
            text="Las cartas pendientes aparecen en la búsqueda general de Cartas.",
            text_color="gray60", font=ctk.CTkFont(size=14)
        ).grid(row=1, column=0, sticky="w", padx=8, pady=(0, 2))

        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)
        f.grid_rowconfigure(2, weight=1)

        cols = ("nombre", "tipo", "notas")
        self._pt = ttk.Treeview(tf, columns=cols, show="headings")
        for col, head, width in [("nombre","Nombre",320),("tipo","Tipo",120),("notas","Notas",350)]:
            self._pt.heading(col, text=head)
            self._pt.column(col, width=width, anchor="w")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._pt.yview)
        self._pt.configure(yscrollcommand=vsb.set)
        self._pt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._pt.bind("<Delete>", self._delete_pendiente)

    def _add_pendiente(self):
        nombre = self._pend_nombre.get().strip()
        if not nombre:
            messagebox.showwarning("Error", "Nombre requerido")
            return
        db.add_pendiente(nombre, self._pend_tipo.get(),
                         self._pend_notas.get().strip() or None)
        self._pend_nombre.delete(0, "end")
        self._pend_notas.delete(0, "end")
        self._refresh_pendientes()

    def _refresh_pendientes(self):
        self._pt.delete(*self._pt.get_children())
        for r in db.get_pendientes():
            r = dict(r)
            self._pt.insert("", "end", iid=str(r["id"]),
                values=(r["nombre"], r.get("tipo",""), r.get("notas","") or ""))

    def _delete_pendiente(self, _=None):
        sel = self._pt.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar esta carta pendiente?"):
            db.delete_pendiente(int(sel[0]))
            self._refresh_pendientes()

    # ── Álbumes ───────────────────────────────────────────────────────────────

    def _build_albumes(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)
        self._frames["albumes"] = f

        # Formulario añadir
        form = ctk.CTkFrame(f)
        form.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(form, text="Nombre álbum:").grid(row=0, column=0, padx=8, pady=8)
        self._alb_nombre = ctk.CTkEntry(form, placeholder_text="Ej: Album 1")
        self._alb_nombre.grid(row=0, column=1, padx=6, pady=8, sticky="ew")

        ctk.CTkLabel(form, text="Páginas:").grid(row=0, column=2, padx=8)
        self._alb_paginas = ctk.CTkEntry(form, width=70, placeholder_text="9")
        self._alb_paginas.grid(row=0, column=3, padx=6)

        ctk.CTkButton(form, text="+ Añadir álbum", width=120,
                      command=self._add_album).grid(row=0, column=4, padx=10)

        # Tabla
        tf = ctk.CTkFrame(f, fg_color="transparent")
        tf.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ("nombre", "paginas", "ubicaciones")
        self._alb_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for col, head, width in [
            ("nombre",     "Álbum",       250),
            ("paginas",    "Páginas",      90),
            ("ubicaciones","Ubicaciones generadas", 300),
        ]:
            self._alb_tree.heading(col, text=head)
            self._alb_tree.column(col, width=width, anchor="w" if col != "paginas" else "center")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self._alb_tree.yview)
        self._alb_tree.configure(yscrollcommand=vsb.set)
        self._alb_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._alb_tree.bind("<Delete>", self._delete_album)
        self._alb_tree.bind("<Double-1>", self._edit_album_dialog)

        ctk.CTkLabel(f,
            text="Doble clic para editar · Supr para eliminar · Las ubicaciones se generan automáticamente",
            text_color="gray60", font=ctk.CTkFont(size=12)
        ).grid(row=2, column=0, sticky="w", padx=8, pady=(0, 4))

    def _refresh_albumes(self):
        self._alb_tree.delete(*self._alb_tree.get_children())
        for a in db.get_albums():
            ejemplo = f"{a['nombre']}.pag.1  →  {a['nombre']}.pag.{a['paginas']}"
            self._alb_tree.insert("", "end", iid=str(a["id"]),
                values=(a["nombre"], a["paginas"], ejemplo))

    def _add_album(self):
        nombre = self._alb_nombre.get().strip()
        try:
            paginas = int(self._alb_paginas.get().strip() or 1)
        except ValueError:
            paginas = 1
        if not nombre:
            messagebox.showwarning("Error", "Nombre requerido")
            return
        db.add_album(nombre, paginas)
        self._alb_nombre.delete(0, "end")
        self._alb_paginas.delete(0, "end")
        self._refresh_albumes()
        self._refresh_dialog_opts()

    def _delete_album(self, _=None):
        sel = self._alb_tree.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar este álbum?"):
            db.delete_album(int(sel[0]))
            self._refresh_albumes()

    def _edit_album_dialog(self, _=None):
        sel = self._alb_tree.selection()
        if not sel:
            return
        album_id = int(sel[0])
        vals = self._alb_tree.item(sel[0], "values")
        nombre_actual, paginas_actual = vals[0], vals[1]

        dlg = ctk.CTkToplevel(self)
        dlg.title("Editar álbum")
        dlg.geometry("360x160")
        dlg.grab_set()
        dlg.resizable(False, False)

        ctk.CTkLabel(dlg, text="Nombre:").grid(row=0, column=0, padx=14, pady=12)
        e_nombre = ctk.CTkEntry(dlg, width=200)
        e_nombre.insert(0, nombre_actual)
        e_nombre.grid(row=0, column=1, padx=8)

        ctk.CTkLabel(dlg, text="Páginas:").grid(row=1, column=0, padx=14, pady=6)
        e_pags = ctk.CTkEntry(dlg, width=80)
        e_pags.insert(0, paginas_actual)
        e_pags.grid(row=1, column=1, padx=8, sticky="w")

        def guardar():
            try:
                pags = int(e_pags.get().strip() or 1)
            except ValueError:
                pags = 1
            db.update_album(album_id, e_nombre.get().strip(), pags)
            dlg.destroy()
            self._refresh_albumes()
            self._refresh_dialog_opts()

        ctk.CTkButton(dlg, text="Guardar", command=guardar).grid(
            row=2, column=0, columnspan=2, pady=14)

    # ── Estadísticas ──────────────────────────────────────────────────────────

    def _build_stats(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure((0, 1, 2), weight=1)
        f.grid_rowconfigure(1, weight=1)
        self._frames["stats"] = f

        self._stat_lbls = {}
        for i, (key, label, color) in enumerate([
            ("total",     "Total colección",  "#f0c040"),
            ("owned",     "Cartas tengo",      "#2ecc71"),
            ("restantes", "Me faltan",         "#e74c3c"),
        ]):
            card = ctk.CTkFrame(f, corner_radius=12)
            card.grid(row=0, column=i, padx=12, pady=16, sticky="nsew")
            ctk.CTkLabel(card, text=label,
                         font=ctk.CTkFont(size=14)).pack(pady=(18, 4))
            lbl = ctk.CTkLabel(card, text="—",
                               font=ctk.CTkFont(size=46, weight="bold"),
                               text_color=color)
            lbl.pack(pady=(0, 18))
            self._stat_lbls[key] = lbl

        tf = ctk.CTkFrame(f)
        tf.grid(row=1, column=0, columnspan=3, padx=8, pady=8, sticky="nsew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        cols = ("tipo", "distintas", "unidades", "vacios")
        self._stree = ttk.Treeview(tf, columns=cols, show="headings")
        for col, head, width in [
            ("tipo",      "Tipo",            200),
            ("distintas", "Cartas distintas", 150),
            ("unidades",  "Total unidades",   150),
            ("vacios",    "Huecos vacíos",    130),
        ]:
            self._stree.heading(col, text=head)
            self._stree.column(col, width=width, anchor="center")
        self._stree.pack(fill="both", expand=True, padx=6, pady=6)

    def _refresh_stats(self):
        stats = db.get_stats()
        self._stat_lbls["total"].configure(text=f"{stats['total']:,}")
        self._stat_lbls["owned"].configure(text=f"{stats['owned']:,}")
        self._stat_lbls["restantes"].configure(text=f"{stats['restantes']:,}")
        self._stree.delete(*self._stree.get_children())
        for tipo, distintas, total, vacios in stats["por_tipo"]:
            self._stree.insert("", "end",
                values=(tipo, distintas - vacios, total, vacios))

    # ── Gráficas ──────────────────────────────────────────────────────────────

    def _build_graficas(self):
        f = ctk.CTkFrame(self._main, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(0, weight=1)
        self._frames["graficas"] = f
        self._graf_canvas = None

    def _refresh_graficas(self):
        stats = db.get_stats()
        datos = [(t, d - v, u, v) for t, d, u, v in stats["por_tipo"] if (d - v) > 0]
        if not datos:
            return

        if self._graf_canvas:
            self._graf_canvas.get_tk_widget().destroy()
            plt.close("all")

        tipos  = [d[0] for d in datos]
        distintas = [d[1] for d in datos]
        unidades  = [d[2] for d in datos]
        colors = [TIPO_COLORS.get(t, "#95a5a6") for t in tipos]

        fig, axes = plt.subplots(1, 3, figsize=(15, 5), facecolor="#1e1e2e")

        # Bar: cartas distintas
        ax = axes[0]
        bars = ax.bar(tipos, distintas, color=colors, edgecolor="none")
        ax.set_title("Cartas distintas por tipo", color="white", fontsize=14, pad=10)
        ax.set_facecolor("#2b2b2b")
        ax.tick_params(colors="white")
        ax.xaxis.set_tick_params(rotation=35)
        for b, v in zip(bars, distintas):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.3,
                    str(v), ha="center", va="bottom", color="white", fontsize=10)
        ax.spines[:].set_color("#444")

        # Pie: distribución
        ax2 = axes[1]
        wedges, _, autotexts = ax2.pie(
            distintas, labels=tipos, colors=colors,
            autopct="%1.1f%%", startangle=90,
            textprops={"color": "white", "fontsize": 8})
        for at in autotexts:
            at.set_fontsize(7)
        ax2.set_title("Distribución por tipo", color="white", fontsize=14, pad=10)
        ax2.set_facecolor("#2b2b2b")

        # Bar: total unidades
        ax3 = axes[2]
        bars3 = ax3.bar(tipos, unidades, color=colors, edgecolor="none")
        ax3.set_title("Total unidades por tipo", color="white", fontsize=14, pad=10)
        ax3.set_facecolor("#2b2b2b")
        ax3.tick_params(colors="white")
        ax3.xaxis.set_tick_params(rotation=35)
        for b, v in zip(bars3, unidades):
            ax3.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.3,
                     str(v), ha="center", va="bottom", color="white", fontsize=10)
        ax3.spines[:].set_color("#444")

        fig.patch.set_facecolor("#1e1e2e")
        fig.tight_layout(pad=2.5)

        canvas = FigureCanvasTkAgg(fig, master=self._frames["graficas"])
        canvas.draw()
        canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self._graf_canvas = canvas

    # ── Refresh global ────────────────────────────────────────────────────────

    def _refresh_all(self):
        self._refresh_cartas()
        self._refresh_repetidos()
        self._refresh_decks()
        self._refresh_pendientes()
        self._refresh_albumes()
        self._refresh_stats()
        self._refresh_graficas()
        self._refresh_global_search()

    # ── Import ────────────────────────────────────────────────────────────────

    def _check_import(self):
        conn = db.get_conn()
        count = conn.execute("SELECT COUNT(*) FROM cartas").fetchone()[0]
        conn.close()
        if count == 0:
            self._reimport()

    def _reimport(self):
        if not messagebox.askyesno(
            "⚠ Reimportar desde Excel",
            "Esto BORRARÁ y recreará:\n"
            "  • Todas las cartas\n"
            "  • Todos los decks\n\n"
            "Se CONSERVARÁ:\n"
            "  • Álbumes\n"
            "  • Cartas pendientes\n"
            "  • Configuración\n\n"
            "Las ediciones manuales (ubicación, cant2…)\n"
            "que no estén en el Excel se perderán.\n\n"
            "¿Continuar?"):
            return

        path = EXCEL_PATH if os.path.exists(EXCEL_PATH) else ""
        if not path:
            path = filedialog.askopenfilename(
                title="Seleccionar Excel Yugioh",
                filetypes=[("Excel", "*.xlsx *.xls")])
            if not path:
                return

        pw = ctk.CTkToplevel(self)
        pw.title("Importando…")
        pw.geometry("340x120")
        pw.grab_set()
        pw.resizable(False, False)
        lbl = ctk.CTkLabel(pw, text="Leyendo Excel…",
                           font=ctk.CTkFont(size=14))
        lbl.pack(pady=(20, 8))
        bar = ctk.CTkProgressBar(pw, width=280)
        bar.pack()
        bar.set(0)

        done_flag = threading.Event()

        def do():
            def cb(i, total):
                bar.set(i / total)
            importer.import_excel(path, progress_cb=cb)
            done_flag.set()

        threading.Thread(target=do, daemon=True).start()

        def poll():
            if done_flag.is_set():
                pw.destroy()
                self._refresh_all()
            else:
                self.after(150, poll)

        self.after(150, poll)

    def _exportar_excel(self):
        path = filedialog.asksaveasfilename(
            title="Exportar Excel",
            initialfile="Biblioteca_Yugioh_export.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self._run_excel_op(lambda: exporter.export_excel(path),
                           f"Exportado en:\n{path}")

    def _actualizar_excel(self):
        if not os.path.exists(EXCEL_PATH):
            messagebox.showwarning("Error", "No se encuentra el Excel original.")
            return

        if not messagebox.askyesno(
            "Actualizar Excel original",
            "Esto sobreescribirá el archivo original:\n"
            f"{EXCEL_PATH}\n\n"
            "Se creará una copia de seguridad automáticamente.\n\n"
            "¿Continuar?",
        ):
            return

        def op():
            backup = exporter.update_excel(EXCEL_PATH)
            return backup

        self._run_excel_op(op, mostrar_backup=True)

    def _importar_bd(self):
        path = filedialog.askopenfilename(
            title="Seleccionar base de datos",
            filetypes=[("SQLite DB", "*.db"), ("Todos", "*.*")])
        if not path:
            return
        if not messagebox.askyesno(
            "Importar BD",
            "Esto reemplazará la base de datos actual con la seleccionada.\n\n"
            "¿Continuar?"):
            return
        import shutil
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = db.DB_PATH + f".backup_{ts}"
        if os.path.exists(db.DB_PATH):
            shutil.copy2(db.DB_PATH, backup)
        shutil.copy2(path, db.DB_PATH)
        self._refresh_all()
        messagebox.showinfo("Listo", f"Base de datos cargada.\nBackup anterior en:\n{backup}")

    def _run_excel_op(self, fn, msg=None, mostrar_backup=False):
        pw = ctk.CTkToplevel(self)
        pw.title("Procesando…")
        pw.geometry("320x90")
        pw.grab_set()
        pw.resizable(False, False)
        ctk.CTkLabel(pw, text="Escribiendo Excel…",
                     font=ctk.CTkFont(size=14)).pack(pady=(18, 6))
        bar = ctk.CTkProgressBar(pw, width=260, mode="indeterminate")
        bar.pack()
        bar.start()

        result_box = [None]
        done_flag = threading.Event()

        def do():
            result_box[0] = fn()
            done_flag.set()

        threading.Thread(target=do, daemon=True).start()

        def poll():
            if done_flag.is_set():
                bar.stop()
                pw.destroy()
                if mostrar_backup and result_box[0]:
                    messagebox.showinfo("Listo",
                        f"Excel actualizado.\n\nBackup guardado en:\n{result_box[0]}")
                elif msg:
                    messagebox.showinfo("Listo", msg)
            else:
                self.after(150, poll)

        self.after(150, poll)


if __name__ == "__main__":
    import sys
    app = App()
    app.protocol("WM_DELETE_WINDOW", lambda: (app.destroy(), sys.exit(0)))
    app.mainloop()
