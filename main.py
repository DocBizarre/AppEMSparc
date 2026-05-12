"""
EMS – Emeraude Moteurs Systèmes
Outil de suivi des interventions et demandes clients
Version 1.8 – Tableau de bord modulable, urgence, classifications, heure Paris
"""

import json
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import date, datetime
from pathlib import Path

import database as db
import mailer
import csv_importer
from bon_generator import sauvegarder_bon, ouvrir_fichier

# ─── Palette ──────────────────────────────────────────────────────────────────
C = {
    # ─── Fonds ────────────────────────────────────────────────────────────────
    "bg":         "#f5f7fa",   # gris très clair (au lieu de bleu pâle)
    "bg_alt":     "#eaeef3",   # gris légèrement plus marqué
    "surface":    "#ffffff",   # cartes/inputs
    # ─── Identité EMS ─────────────────────────────────────────────────────────
    "header":     "#002b5c",   # bleu marine EMS (plus profond que #003366)
    "header_alt": "#003d7a",   # bleu intermédiaire
    "nav_sel":    "#0056b3",   # bleu sélection nav
    "accent":     "#c62828",   # rouge EMS (du logo) — pour accents
    # ─── Texte ────────────────────────────────────────────────────────────────
    "text":       "#1a2332",   # presque noir
    "text_muted": "#6b7785",   # gris moyen
    "text_light": "#9ba5b1",   # gris clair
    # ─── Boutons ──────────────────────────────────────────────────────────────
    "btn":        "#0056b3",   # bleu primaire
    "btn_hover":  "#003d80",
    "btn2":       "#1e7e3e",   # vert action positive
    "btn2_hover": "#155a2c",
    "btn3":       "#6b7785",   # gris pour annuler/neutre
    "btn3_hover": "#525c66",
    "danger":     "#c62828",
    "danger_hover":"#a32020",
    "warn":       "#e67e22",   # orange chaud
    # ─── Tableaux ─────────────────────────────────────────────────────────────
    "row_even":   "#f7f9fc",   # bandes très discrètes
    "row_odd":    "#ffffff",
    "row_hover":  "#e3eaf3",
    "border":     "#d8dee5",   # bordures douces
    # ─── États / urgence ──────────────────────────────────────────────────────
    "urg_normale":  "#6b7785",
    "urg_urgente":  "#e67e22",
    "urg_critique": "#c62828",
    "ec":           "#f59e0b",  # En cours - ambre
    "ec_bg":        "#fff7ed",
    "clos":         "#10b981",  # Clos - vert
    "clos_bg":      "#ecfdf5",
    "fact":         "#3b82f6",  # Facturé - bleu vif
    "fact_bg":      "#eff6ff",
}
STATUTS  = ["En cours", "Clos", "Facturé"]
URGENCES = ["Normale", "Urgente", "Critique"]

# Polices unifiées (Tkinter accepte des tuples ou un nom de famille)
F = {
    "title":      ("Segoe UI", 16, "bold"),
    "subtitle":   ("Segoe UI", 10),
    "h1":         ("Segoe UI", 14, "bold"),
    "h2":         ("Segoe UI", 12, "bold"),
    "h3":         ("Segoe UI", 10, "bold"),
    "body":       ("Segoe UI", 10),
    "body_bold":  ("Segoe UI", 10, "bold"),
    "small":      ("Segoe UI", 9),
    "small_bold": ("Segoe UI", 9, "bold"),
    "tiny":       ("Segoe UI", 8),
    "mono":       ("Consolas", 10),
    "stat_value": ("Segoe UI", 26, "bold"),
    "stat_label": ("Segoe UI", 9),
}


# ─── Helpers ──────────────────────────────────────────────────────────────────
def row_get(row, key, default=""):
    if row is None:
        return default
    try:
        v = row[key]
        return v if v is not None else default
    except (KeyError, IndexError):
        return default


def mk_header(parent, title, subtitle=""):
    """Header de page : bandeau bleu marine avec titre + sous-titre optionnel."""
    bar = tk.Frame(parent, bg=C["header"], height=64)
    bar.pack(fill="x")
    bar.pack_propagate(False)
    # Petit liseré rouge à gauche pour rappeler le logo EMS
    accent = tk.Frame(bar, bg=C["accent"], width=4)
    accent.pack(side="left", fill="y")
    inner = tk.Frame(bar, bg=C["header"])
    inner.pack(side="left", fill="both", expand=True, padx=20)
    tk.Label(inner, text=title, font=F["title"],
             bg=C["header"], fg="white").pack(side="top", anchor="w", pady=(10, 0))
    if subtitle:
        tk.Label(inner, text=subtitle, font=F["subtitle"],
                 bg=C["header"], fg="#aac4e8").pack(side="top", anchor="w")


def mk_btn(parent, text, cmd, color=None, hover=None, **kw):
    """
    Bouton stylisé avec effet hover.
    color : couleur de fond (défaut: C["btn"]). hover : couleur survol (déduite si None).
    """
    bg = color or C["btn"]
    # Déduire la couleur hover si non précisée
    hover_map = {
        C["btn"]:    C["btn_hover"],
        C["btn2"]:   C["btn2_hover"],
        C["btn3"]:   C["btn3_hover"],
        C["danger"]: C["danger_hover"],
    }
    hover_bg = hover or hover_map.get(bg) or bg

    btn = tk.Button(parent, text=text, command=cmd,
                    bg=bg, fg="white",
                    font=F["body_bold"], relief="flat",
                    bd=0, padx=14, pady=6, cursor="hand2",
                    activebackground=hover_bg, activeforeground="white",
                    **kw)
    # Effet hover via bindings
    def _on_enter(_e): btn.configure(bg=hover_bg)
    def _on_leave(_e): btn.configure(bg=bg)
    btn.bind("<Enter>", _on_enter)
    btn.bind("<Leave>", _on_leave)
    return btn


def mk_tree(parent, cols, col_defs, height=18):
    """col_defs : liste de tuples (cid, lbl, width, [anchor])."""
    # Container avec bordure douce
    wrapper = tk.Frame(parent, bg=C["border"], bd=0)
    frame = tk.Frame(wrapper, bg=C["surface"])
    frame.pack(fill="both", expand=True, padx=1, pady=1)

    vsb = ttk.Scrollbar(frame, orient="vertical")
    hsb = ttk.Scrollbar(frame, orient="horizontal")
    tree = ttk.Treeview(frame, columns=cols, show="headings",
                        height=height, yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                        style="EMS.Treeview")
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    # Tags pour urgence (fond + couleur de texte)
    tree.tag_configure("even", background=C["row_even"])
    tree.tag_configure("odd",  background=C["row_odd"])
    tree.tag_configure("urg_critique", background="#fef2f2", foreground=C["urg_critique"])
    tree.tag_configure("urg_urgente",  background="#fff7ed", foreground=C["urg_urgente"])

    for cd in col_defs:
        if len(cd) == 4:
            cid, lbl, w, anchor = cd
        else:
            cid, lbl, w = cd
            anchor = "w"
        tree.heading(cid, text=lbl, anchor=anchor)
        tree.column(cid, width=w, minwidth=50, anchor=anchor)
    return wrapper, tree


def fill_tree(tree, rows, urgences=None):
    """rows: liste de tuples. urgences: liste optionnelle de la même longueur."""
    tree.delete(*tree.get_children())
    for i, r in enumerate(rows):
        tags = ["even" if i % 2 == 0 else "odd"]
        if urgences:
            u = urgences[i]
            if u == "Critique": tags.append("urg_critique")
            elif u == "Urgente": tags.append("urg_urgente")
        tree.insert("", "end", iid=str(i), values=r, tags=tuple(tags))


def section_bar(parent, title):
    """Titre de section avec underline coloré (plus discret que bandeau bleu plein)."""
    wrapper = tk.Frame(parent, bg=C["bg"])
    wrapper.pack(fill="x", pady=(14, 4))
    tk.Label(wrapper, text=title.upper(),
             bg=C["bg"], fg=C["header"],
             font=("Segoe UI", 10, "bold")).pack(side="left", padx=4)
    # Ligne fine accent rouge sous le titre
    line = tk.Frame(parent, bg=C["accent"], height=2)
    line.pack(fill="x", padx=2)


def notif_icon(client_notifie, tech_notifie):
    """Icônes claires pour la colonne Notif. : ✓ pour notifié, • pour non."""
    c = "✓" if client_notifie else "○"
    t = "✓" if tech_notifie else "○"
    return f"📧 {c} / 🔧 {t}"


# ─── Champ DATE avec masque JJ/MM/AAAA ────────────────────────────────────────
class DateEntry(ttk.Entry):
    def __init__(self, master, textvariable=None, width=46, **kw):
        self.var = textvariable or tk.StringVar()
        super().__init__(master, textvariable=self.var, width=width, **kw)
        self._lock = False
        self.var.trace_add("write", self._on_write)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<KeyPress>", self._on_key)

    def _on_key(self, ev):
        if ev.char and not ev.char.isdigit() and ev.char not in ("/", "\b", "\x7f", ""):
            return "break"

    def _on_write(self, *_):
        if self._lock:
            return
        s = self.var.get()
        digits = "".join(ch for ch in s if ch.isdigit())[:8]
        if len(digits) <= 2:
            out = digits
        elif len(digits) <= 4:
            out = digits[:2] + "/" + digits[2:]
        else:
            out = digits[:2] + "/" + digits[2:4] + "/" + digits[4:]
        if out != s:
            self._lock = True
            self.var.set(out)
            self._lock = False
        try:
            self.configure(foreground="black")
        except tk.TclError:
            pass

    def _on_focus_out(self, _ev=None):
        s = self.var.get().strip()
        if not s:
            self.configure(foreground="black")
            return
        self.configure(foreground="black" if self.is_valid(s) else C["danger"])

    @staticmethod
    def is_valid(s):
        if not s or len(s) != 10:
            return False
        try:
            datetime.strptime(s, "%d/%m/%Y")
            return True
        except ValueError:
            return False


# ─── Champ EMAIL avec validation NON BLOQUANTE ────────────────────────────────
class EmailEntry(ttk.Entry):
    """
    Entry avec validation visuelle (rouge si format douteux) mais NON BLOQUANTE.
    L'utilisateur peut toujours saisir ce qu'il veut.
    """
    def __init__(self, master, textvariable=None, width=46, **kw):
        self.var = textvariable or tk.StringVar()
        super().__init__(master, textvariable=self.var, width=width, **kw)
        self.bind("<FocusOut>", self._check)
        self.bind("<KeyRelease>", lambda _ev: self.configure(foreground="black"))

    def _check(self, _ev=None):
        s = self.var.get().strip()
        if not s or db.email_looks_valid(s):
            self.configure(foreground="black")
        else:
            self.configure(foreground=C["warn"])  # orange, pas rouge


# ─── Combobox avec recherche textuelle ────────────────────────────────────────
class SearchableCombobox(ttk.Combobox):
    """
    Combobox dont la liste se filtre au fur et à mesure de la saisie.
    - all_values : liste complète des choix possibles
    - L'utilisateur peut taper pour filtrer (insensible casse/accents)
    - La liste déroulante affiche uniquement les correspondances
    - Si focus perdu sans valeur valide : restaure la dernière sélection valide

    Compatible drop-in avec ttk.Combobox : on peut utiliser .get(), .set() normalement.
    """
    def __init__(self, master, textvariable=None, values=None, width=44,
                 case_sensitive=False, **kw):
        # Important : NE PAS mettre state="readonly", on veut autoriser la saisie
        kw.pop("state", None)
        self._all_values    = list(values or [])
        self._case_sensitive = case_sensitive
        self._last_valid    = ""
        self._suppress      = False
        self.var = textvariable or tk.StringVar()
        super().__init__(master, textvariable=self.var, values=self._all_values,
                         width=width, **kw)
        self.var.trace_add("write", self._on_write)
        self.bind("<<ComboboxSelected>>", self._on_select)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<KeyRelease>", self._on_keyrelease)

    # API publique
    def set_values(self, values):
        """Met à jour la liste complète des choix."""
        self._all_values = list(values or [])
        self["values"] = self._all_values

    def get_all_values(self):
        return list(self._all_values)

    # Override de Combobox.set() pour mémoriser la sélection valide
    def set(self, value):
        self._suppress = True
        super().set(value)
        self._suppress = False
        if value in self._all_values:
            self._last_valid = value
        # Reset le filtre : on remet toutes les valeurs visibles
        self["values"] = self._all_values

    # Internes
    @staticmethod
    def _norm(s):
        """Normalise pour comparaison : lower + suppression accents communs."""
        if s is None:
            return ""
        s = str(s).lower()
        # Translit ASCII basique pour les accents français
        trans = str.maketrans("àâäéèêëîïôöùûüç", "aaaeeeeiioouuuc")
        return s.translate(trans)

    def _filter(self, query):
        """Retourne les valeurs contenant query (mode insensible casse/accents)."""
        if not query:
            return list(self._all_values)
        q = query if self._case_sensitive else self._norm(query)
        if self._case_sensitive:
            return [v for v in self._all_values if q in v]
        return [v for v in self._all_values if q in self._norm(v)]

    def _on_write(self, *_):
        if self._suppress:
            return
        # Quand l'utilisateur tape, on filtre la liste sans toucher au champ
        query = self.var.get()
        filtered = self._filter(query)
        self["values"] = filtered if filtered else self._all_values

    def _on_keyrelease(self, ev):
        # Ne pas réagir aux flèches/Tab (laissent ttk faire son boulot)
        if ev.keysym in ("Up", "Down", "Left", "Right", "Tab",
                          "Return", "Escape", "Shift_L", "Shift_R",
                          "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return
        # Sur Backspace ou caractère imprimable : ouvrir la dropdown automatiquement
        if ev.keysym == "BackSpace" or (ev.char and ev.char.isprintable()):
            # Toujours afficher la liste filtrée à jour
            try:
                # Si le champ est non vide, force l'ouverture de la liste
                if self.var.get():
                    self.event_generate("<Down>")
                    self.icursor("end")
            except tk.TclError:
                pass

    def _on_select(self, _ev=None):
        v = self.var.get()
        if v in self._all_values:
            self._last_valid = v
        # Reset filtre après sélection
        self["values"] = self._all_values

    def _on_focus_out(self, _ev=None):
        # Si la valeur n'est pas dans la liste, restaurer la dernière valide
        # (sauf si on autorise les valeurs libres, ce qui n'est pas notre cas ici)
        v = self.var.get()
        if v and v not in self._all_values:
            # Coloration warning, mais ne réécrit pas (l'utilisateur peut vouloir l'ajouter)
            try:
                self.configure(foreground=C["warn"])
            except tk.TclError:
                pass
        else:
            try:
                self.configure(foreground="black")
            except tk.TclError:
                pass
        # Reset le filtre quand on quitte
        self["values"] = self._all_values


# ══════════════════════════════════════════════════════════════════════════════
# FENÊTRE PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════
class AppEMS(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EMS – Gestion des Interventions")
        self.geometry("1340x830")
        self.minsize(960, 620)
        self.configure(bg=C["bg"])
        self._init_ttk_styles()
        db.init_db()
        self._build()
        self.show("dashboard")

    def _init_ttk_styles(self):
        """Configure les styles ttk : Treeview, Combobox, Entry, Notebook."""
        s = ttk.Style(self)
        # Utiliser le thème 'clam' pour avoir plus de contrôle sur les couleurs
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        # ── Treeview ─────────────────────────────────────────────────────────
        s.configure("EMS.Treeview",
                    background=C["surface"],
                    foreground=C["text"],
                    fieldbackground=C["surface"],
                    borderwidth=0,
                    rowheight=28,
                    font=F["body"])
        s.configure("EMS.Treeview.Heading",
                    background=C["bg_alt"],
                    foreground=C["header"],
                    relief="flat",
                    borderwidth=0,
                    font=F["small_bold"],
                    padding=(8, 6))
        s.map("EMS.Treeview.Heading",
              background=[("active", C["bg_alt"])])
        s.map("EMS.Treeview",
              background=[("selected", C["nav_sel"])],
              foreground=[("selected", "white")])
        # Compat : tags personnalisés gardent priorité, donc fond reste appliqué

        # Fallback : style Treeview par défaut aussi
        s.configure("Treeview",
                    background=C["surface"],
                    foreground=C["text"],
                    fieldbackground=C["surface"],
                    rowheight=28,
                    font=F["body"])
        s.configure("Treeview.Heading",
                    background=C["bg_alt"],
                    foreground=C["header"],
                    font=F["small_bold"],
                    padding=(8, 6))

        # ── Entry / Combobox ─────────────────────────────────────────────────
        s.configure("TEntry",
                    fieldbackground=C["surface"],
                    foreground=C["text"],
                    bordercolor=C["border"],
                    lightcolor=C["border"],
                    darkcolor=C["border"],
                    padding=4)
        s.configure("TCombobox",
                    fieldbackground=C["surface"],
                    background=C["surface"],
                    foreground=C["text"],
                    bordercolor=C["border"],
                    arrowcolor=C["header"],
                    padding=4)
        s.map("TCombobox",
              fieldbackground=[("readonly", C["surface"])],
              foreground=[("readonly", C["text"])])

        # ── Scrollbar ────────────────────────────────────────────────────────
        s.configure("Vertical.TScrollbar",
                    background=C["bg_alt"],
                    troughcolor=C["bg"],
                    borderwidth=0,
                    arrowcolor=C["text_muted"])
        s.configure("Horizontal.TScrollbar",
                    background=C["bg_alt"],
                    troughcolor=C["bg"],
                    borderwidth=0,
                    arrowcolor=C["text_muted"])

        # ── Notebook (onglets) ───────────────────────────────────────────────
        s.configure("TNotebook",
                    background=C["bg"],
                    borderwidth=0)
        s.configure("TNotebook.Tab",
                    background=C["bg_alt"],
                    foreground=C["text_muted"],
                    padding=(16, 8),
                    font=F["body_bold"],
                    borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", C["surface"])],
              foreground=[("selected", C["header"])])

        # ── Séparateur ───────────────────────────────────────────────────────
        s.configure("TSeparator", background=C["border"])

        # ── Police par défaut pour les Labels Tk ─────────────────────────────
        self.option_add("*Label.Font", F["body"])
        self.option_add("*Frame.Background", C["bg"])

    def _build(self):
        # Sidebar
        self.sidebar = tk.Frame(self, bg=C["header"], width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # ── Zone logo (fond légèrement plus sombre pour distinguer) ──────────
        logo_zone = tk.Frame(self.sidebar, bg=C["header"])
        logo_zone.pack(fill="x", pady=(16, 8))

        self._logo_img = None
        logo_loaded = False
        try:
            from pathlib import Path as _P
            logo_file = _P(__file__).parent / "assets" / "logo_ems.png"
            if logo_file.is_file():
                self._logo_img = tk.PhotoImage(file=str(logo_file))
                logo_loaded = True
        except (tk.TclError, OSError):
            pass
        if not logo_loaded:
            try:
                from logo_data import LOGO_EMS_B64
                if LOGO_EMS_B64:
                    self._logo_img = tk.PhotoImage(data=LOGO_EMS_B64)
                    logo_loaded = True
            except (ImportError, tk.TclError):
                pass

        if logo_loaded and self._logo_img is not None:
            w = self._logo_img.width()
            max_w = 150
            if w > max_w:
                factor = max(1, w // max_w)
                try:
                    self._logo_img = self._logo_img.subsample(factor, factor)
                except tk.TclError:
                    pass
            tk.Label(logo_zone, image=self._logo_img,
                     bg=C["header"]).pack(pady=(4, 2))
        else:
            tk.Label(logo_zone, text="EMS", font=("Segoe UI", 28, "bold"),
                     bg=C["header"], fg="white").pack(pady=(8, 0))
        tk.Label(logo_zone, text="Emeraude Moteurs Systèmes",
                 font=F["tiny"], bg=C["header"], fg="#aac4e8",
                 justify="center").pack(pady=(2, 0))

        # Liseré rouge accent sous le logo
        tk.Frame(self.sidebar, bg=C["accent"], height=2).pack(fill="x", padx=14, pady=(8, 10))

        # ── Recherche garage (encadré subtil) ────────────────────────────────
        garage = tk.Frame(self.sidebar, bg=C["header"])
        garage.pack(fill="x", padx=14, pady=(0, 12))
        tk.Label(garage, text="RECHERCHE N° SÉRIE", bg=C["header"], fg="#7fa6cc",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 4))
        self.garage_var = tk.StringVar()
        e = ttk.Entry(garage, textvariable=self.garage_var, width=22)
        e.pack(fill="x", pady=(0, 4))
        e.bind("<Return>", lambda _ev: self._garage_search())
        mk_btn(garage, "🔍 Trouver moteur", self._garage_search).pack(fill="x")

        # Séparateur fin
        tk.Frame(self.sidebar, bg="#1a447a", height=1).pack(fill="x", padx=14, pady=(4, 8))

        # ── Navigation principale ────────────────────────────────────────────
        self._nav_btns = {}
        self._nav_widgets = {}
        for key, icon, lbl in [
            ("dashboard",     "📊", "Tableau de bord"),
            ("interventions", "🔧", "Interventions"),
            ("clients",       "👥", "Clients"),
            ("moteurs",       "⚙",  "Moteurs"),
            ("techniciens",   "🔨", "Techniciens"),
            ("nouveau",       "➕", "Nouveau bon"),
        ]:
            row = tk.Frame(self.sidebar, bg=C["header"], cursor="hand2")
            row.pack(fill="x", pady=0)
            # Indicateur de sélection (barre verticale rouge à gauche) — masquée par défaut
            sel_bar = tk.Frame(row, bg=C["header"], width=3)
            sel_bar.pack(side="left", fill="y")
            icon_lbl = tk.Label(row, text=icon, font=("Segoe UI", 13),
                                 bg=C["header"], fg="white",
                                 width=3, anchor="center")
            icon_lbl.pack(side="left", padx=(8, 0), pady=8)
            text_lbl = tk.Label(row, text=lbl, font=("Segoe UI", 11),
                                 bg=C["header"], fg="white", anchor="w")
            text_lbl.pack(side="left", fill="x", expand=True, padx=(6, 8), pady=8)
            # Click handler
            def _make_click(k):
                return lambda _ev=None: self.show(k)
            cb = _make_click(key)
            for w in (row, sel_bar, icon_lbl, text_lbl):
                w.bind("<Button-1>", cb)
            # Hover : éclaircit légèrement
            def _make_hover(widgets, k):
                def enter(_e):
                    # ne pas appliquer si déjà sélectionné
                    cur_key = getattr(self, "_current_nav", None)
                    if cur_key != k:
                        for w in widgets:
                            w.configure(bg=C["header_alt"])
                def leave(_e):
                    cur_key = getattr(self, "_current_nav", None)
                    if cur_key != k:
                        for w in widgets:
                            w.configure(bg=C["header"])
                return enter, leave
            widgets_to_hover = [row, icon_lbl, text_lbl]
            enter, leave = _make_hover(widgets_to_hover, key)
            for w in widgets_to_hover:
                w.bind("<Enter>", enter)
                w.bind("<Leave>", leave)
            self._nav_btns[key] = row
            self._nav_widgets[key] = (sel_bar, icon_lbl, text_lbl)

        # ── Outils en bas ────────────────────────────────────────────────────
        tk.Frame(self.sidebar, bg="#1a447a", height=1).pack(fill="x", padx=14, pady=(10, 6))

        types_btn = tk.Frame(self.sidebar, bg=C["header"], cursor="hand2")
        types_btn.pack(fill="x")
        types_icon = tk.Label(types_btn, text="⚙", font=("Segoe UI", 11),
                               bg=C["header"], fg="#aac4e8", width=3, anchor="center")
        types_icon.pack(side="left", padx=(8, 0), pady=6)
        types_lbl = tk.Label(types_btn, text="Types d'intervention",
                              font=F["small"], bg=C["header"], fg="#aac4e8", anchor="w")
        types_lbl.pack(side="left", fill="x", expand=True, padx=(6, 8), pady=6)
        for w in (types_btn, types_icon, types_lbl):
            w.bind("<Button-1>", lambda _e: TypesDialog(self, self))

        # Version en bas
        tk.Label(self.sidebar, text="v1.8 · Heure de Paris", font=F["tiny"],
                 bg=C["header"], fg="#6699cc").pack(side="bottom", pady=10)

        # Zone principale
        self.main = tk.Frame(self, bg=C["bg"])
        self.main.pack(side="left", fill="both", expand=True)

        self.frames = {}
        for FrameCls, key in [
            (DashboardFrame,     "dashboard"),
            (InterventionsFrame, "interventions"),
            (ClientsFrame,       "clients"),
            (MoteursFrame,       "moteurs"),
            (TechniciensFrame,   "techniciens"),
            (NouveauFrame,       "nouveau"),
        ]:
            f = FrameCls(self.main, self)
            self.frames[key] = f
            f.place(relx=0, rely=0, relwidth=1, relheight=1)

    def show(self, key):
        self._current_nav = key  # pour gérer le hover
        for k, b in self._nav_btns.items():
            is_active = (k == key)
            bg = C["nav_sel"] if is_active else C["header"]
            b.configure(bg=bg)
            if k in self._nav_widgets:
                sel_bar, icon_lbl, text_lbl = self._nav_widgets[k]
                # Barre rouge à gauche uniquement pour l'item actif
                sel_bar.configure(bg=C["accent"] if is_active else bg)
                icon_lbl.configure(bg=bg)
                text_lbl.configure(bg=bg, font=("Segoe UI", 11, "bold") if is_active else ("Segoe UI", 11))
        self.frames[key].lift()
        if hasattr(self.frames[key], "refresh"):
            self.frames[key].refresh()

    def _garage_search(self):
        ns = self.garage_var.get().strip()
        if not ns:
            messagebox.showinfo("Recherche", "Saisissez un N° de série.")
            return
        m = db.find_moteur_by_serie(ns)
        if not m:
            self.show("moteurs")
            self.frames["moteurs"].search_var.set(ns)
            return
        self.show("interventions")
        self.frames["interventions"].search_var.set(ns)
        self.frames["interventions"].statut_var.set("Tous")
        self.garage_var.set("")


# ══════════════════════════════════════════════════════════════════════════════
# WIDGETS DU TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════════════════════
# Chaque widget est un Frame avec une méthode refresh().
# Le DashboardFrame instancie/cache les widgets selon la config utilisateur.

# Catalogue des widgets disponibles
WIDGET_CATALOG = [
    ("stats_cards",        "📊 Cartes statistiques",    "Cartes synthétiques (En cours, Clos, etc.)"),
    ("urgentes",           "⚡ Interventions urgentes", "Liste des bons Urgente / Critique en cours"),
    ("activite_recente",   "🕒 Activité récente",        "Derniers bons modifiés"),
    ("garantie_expirante", "⏰ Garanties expirantes",    "Moteurs dont la garantie expire bientôt"),
    ("par_technicien",     "🛠️ Charge par technicien",   "Nombre d'interventions par technicien"),
    ("par_type",           "🔧 Répartition par type",    "Statistiques par type d'intervention"),
    ("non_notifies",       "📧 Non notifiés",            "Bons en cours sans notification envoyée"),
    ("classifications",    "🏷️ Par classification",     "Garantie / Facturable / Interne"),
]

# Catalogue des cartes statistiques
CARD_CATALOG = [
    # (clé, label affiché, couleur d'accent latéral, couleur du chiffre)
    ("En cours",     "En cours",      C["ec"],          C["ec"]),
    ("Clos",         "Clos",          C["clos"],        C["clos"]),
    ("Facturé",      "Facturé",       C["fact"],        C["fact"]),
    ("Total",        "Total",         C["header"],      C["header"]),
    ("Urgentes",     "Urgentes",      C["urg_urgente"], C["urg_urgente"]),
    ("Critiques",    "Critiques",     C["urg_critique"],C["urg_critique"]),
    ("Garantie",     "Sous garantie", "#0d47a1",        "#0d47a1"),
    ("Facturables",  "Facturables",   "#827717",        "#827717"),
    ("Internes",     "Internes",      "#6a1b9a",        "#6a1b9a"),
    ("Non notifiés", "Non notifiés",  C["accent"],      C["accent"]),
    ("Clients",      "Clients",       "#1b5e20",        "#1b5e20"),
    ("Moteurs",      "Moteurs",       "#e65100",        "#e65100"),
    ("Tech.",        "Techniciens",   "#311b92",        "#311b92"),
]


class StatsCardsWidget(tk.Frame):
    """Cartes statistiques (style 'card' avec barre latérale colorée)."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        self.cards_f = tk.Frame(self, bg=C["bg"])
        self.cards_f.pack(fill="x")

    def refresh(self):
        for w in self.cards_f.winfo_children():
            w.destroy()
        stats = db.get_stats()
        key_map = {
            "En cours": stats["En cours"], "Clos": stats["Clos"],
            "Facturé": stats["Facturé"], "Total": stats["Total"],
            "Urgentes": stats["Urgentes"], "Critiques": stats["Critiques"],
            "Garantie": stats["Garantie"], "Facturables": stats["Facturables"],
            "Internes": stats["Internes"], "Non notifiés": stats["Non notifiés"],
            "Clients": stats["clients"], "Moteurs": stats["moteurs"],
            "Tech.": stats["techniciens"],
        }
        active = db.get_dashboard_cards()
        for key, label, accent_color, value_color in CARD_CATALOG:
            if key not in active:
                continue
            n = key_map.get(key, 0)
            # Container avec bordure douce (simule l'ombre)
            outer = tk.Frame(self.cards_f, bg=C["border"])
            outer.pack(side="left", padx=5, pady=4)
            # Carte : barre accent à gauche + contenu blanc à droite
            card = tk.Frame(outer, bg=C["surface"])
            card.pack(padx=1, pady=1)
            # Barre accent verticale (3px)
            tk.Frame(card, bg=accent_color, width=3).pack(side="left", fill="y")
            # Contenu
            content = tk.Frame(card, bg=C["surface"])
            content.pack(side="left", padx=14, pady=8)
            tk.Label(content, text=str(n), font=F["stat_value"],
                     bg=C["surface"], fg=value_color).pack(anchor="w")
            tk.Label(content, text=label, font=F["stat_label"],
                     bg=C["surface"], fg=C["text_muted"]).pack(anchor="w")


class UrgentesWidget(tk.Frame):
    """Liste des interventions urgentes (Urgente + Critique en cours)."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="⚡ Interventions urgentes en cours",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["danger"]).pack(anchor="w")
        cols = ("urg","num_bon","client","navire","tech","date")
        col_defs = [("urg","⚡",70),("num_bon","N° Bon",115),("client","Client",170),
                    ("navire","Navire",140),("tech","Tech.",110),("date","Date",90)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))
        self._cache = []
        self.tree.bind("<Double-1>", self._open)

    def refresh(self):
        invs = db.get_interventions_urgentes(limit=8)
        self._cache = list(invs)
        rows  = []
        urgs  = []
        for r in self._cache:
            rows.append((r["urgence"], r["num_bon"], r["client_nom"] or "",
                         r["navire"] or "", r["technicien"], r["date_creation"]))
            urgs.append(r["urgence"])
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("—", "(aucune)", "", "", "", ""))
        else:
            fill_tree(self.tree, rows, urgences=urgs)

    def _open(self, _ev):
        sel = self.tree.selection()
        if not sel or not self._cache: return
        try:
            r = self._cache[int(sel[0])]
            BonDialog(self, self.app, inv_id=r["id"], on_save=self.app.frames["dashboard"].refresh)
        except (ValueError, IndexError):
            pass


class ActiviteRecenteWidget(tk.Frame):
    """Derniers bons modifiés."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="🕒 Activité récente",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("num_bon","client","statut","tech","modif")
        col_defs = [("num_bon","N° Bon",115),("client","Client",180),
                    ("statut","Statut",80),("tech","Tech.",110),
                    ("modif","Modifié (Paris)",120)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))
        self._cache = []
        self.tree.bind("<Double-1>", self._open)

    def refresh(self):
        invs = db.get_activite_recente(limit=8)
        self._cache = list(invs)
        rows = [(r["num_bon"], r["client_nom"] or "",
                 r["statut"], r["technicien"],
                 db.fmt_paris_short(r["updated_at"]))
                for r in self._cache]
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("(aucune)", "", "", "", ""))
        else:
            fill_tree(self.tree, rows)

    def _open(self, _ev):
        sel = self.tree.selection()
        if not sel or not self._cache: return
        try:
            r = self._cache[int(sel[0])]
            BonDialog(self, self.app, inv_id=r["id"], on_save=self.app.frames["dashboard"].refresh)
        except (ValueError, IndexError):
            pass


class GarantieExpiranteWidget(tk.Frame):
    """Moteurs dont la garantie expire dans <= 90 jours."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        head = tk.Frame(self, bg=C["bg"])
        head.pack(fill="x")
        tk.Label(head, text="⏰ Garanties expirant dans 90 jours",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["warn"]).pack(side="left")
        cols = ("ns","client","navire","fin","jours")
        col_defs = [("ns","N° Série",130),("client","Client",170),
                    ("navire","Navire/Site",140),("fin","Mise svc",95),
                    ("jours","Jours restants",110, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))

    def refresh(self):
        items = db.get_moteurs_garantie_expirante(jours_max=90)
        rows = []
        for it in items:
            m = it["moteur"]
            rows.append((m["num_serie"], m["client_nom"] or "",
                         row_get(m, "navire"), m["date_mise_service"],
                         f"{it['jours_restants']} j"))
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("—", "(aucune garantie expirant bientôt)", "", "", ""))
        else:
            fill_tree(self.tree, rows)


class ParTechnicienWidget(tk.Frame):
    """Charge de travail par technicien."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="🛠️ Charge par technicien",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("tech","ec","clos","fact","total")
        col_defs = [("tech","Technicien",200),
                    ("ec","En cours",80, "center"),
                    ("clos","Clos",70, "center"),
                    ("fact","Facturé",80, "center"),
                    ("total","Total",70, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))

    def refresh(self):
        rows = db.get_stats_par_technicien()
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("(aucun technicien)", "", "", "", ""))
        else:
            data = [(r["technicien"], r["en_cours"], r["clos"], r["facture"], r["total"])
                    for r in rows]
            fill_tree(self.tree, data)


class ParTypeWidget(tk.Frame):
    """Répartition par type d'intervention."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="🔧 Répartition par type d'intervention",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("type","ec","total")
        col_defs = [("type","Type",260),
                    ("ec","En cours",80, "center"),
                    ("total","Total",80, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))

    def refresh(self):
        rows = db.get_stats_par_type()
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("(aucune intervention)", "", ""))
        else:
            data = [(r["type"], r["en_cours"], r["total"]) for r in rows]
            fill_tree(self.tree, data)


class NonNotifiesWidget(tk.Frame):
    """Bons en cours sans notification envoyée."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="📧 Bons en cours non notifiés",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["warn"]).pack(anchor="w")
        cols = ("num_bon","client","cli","tech","tec")
        col_defs = [("num_bon","N° Bon",115),("client","Client",200),
                    ("cli","Client notifié",110, "center"),
                    ("tech","Technicien",130),
                    ("tec","Tech. notifié",110, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
        tf.pack(fill="x", pady=(2, 0))
        self._cache = []
        self.tree.bind("<Double-1>", self._open)

    def refresh(self):
        invs = db.get_non_notifies(limit=8)
        self._cache = list(invs)
        rows = []
        for r in self._cache:
            rows.append((
                r["num_bon"], r["client_nom"] or "",
                "✓" if row_get(r, "client_notifie") else "✗",
                r["technicien"],
                "✓" if row_get(r, "tech_notifie") else "✗",
            ))
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("—", "(tous notifiés !)", "", "", ""))
        else:
            fill_tree(self.tree, rows)

    def _open(self, _ev):
        sel = self.tree.selection()
        if not sel or not self._cache: return
        try:
            r = self._cache[int(sel[0])]
            BonDialog(self, self.app, inv_id=r["id"], on_save=self.app.frames["dashboard"].refresh)
        except (ValueError, IndexError):
            pass


class ClassificationsWidget(tk.Frame):
    """Mini-vue Garantie / Facturable / Interne."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="🏷️ Répartition par classification",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        self.row = tk.Frame(self, bg=C["bg"])
        self.row.pack(fill="x", pady=(2, 0))

    def refresh(self):
        for w in self.row.winfo_children():
            w.destroy()
        s = db.get_stats()
        for label, val, bg, fg in [
            ("Sous garantie",  s["Garantie"],     "#e3f2fd", "#0d47a1"),
            ("Facturables",    s["Facturables"],  "#fff9c4", "#827717"),
            ("Internes",       s["Internes"],     "#f3e5f5", "#4a148c"),
        ]:
            f = tk.Frame(self.row, bg=bg, relief="solid", bd=1)
            f.pack(side="left", padx=6, pady=2, ipadx=14, ipady=6, fill="x", expand=True)
            tk.Label(f, text=str(val), font=("Arial", 22, "bold"), bg=bg, fg=fg).pack()
            tk.Label(f, text=label, font=("Arial", 9), bg=bg, fg=fg).pack()


# Mapping clé → classe widget
WIDGET_CLASSES = {
    "stats_cards":        StatsCardsWidget,
    "urgentes":           UrgentesWidget,
    "activite_recente":   ActiviteRecenteWidget,
    "garantie_expirante": GarantieExpiranteWidget,
    "par_technicien":     ParTechnicienWidget,
    "par_type":           ParTypeWidget,
    "non_notifies":       NonNotifiesWidget,
    "classifications":    ClassificationsWidget,
}


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG : CONFIGURATION DU TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════════════════════
class DashboardConfigDialog(tk.Toplevel):
    """Permet à l'utilisateur de choisir quels widgets afficher et dans quel ordre."""

    def __init__(self, parent, app, on_save=None):
        super().__init__(parent)
        self.app = app
        self.on_save = on_save
        self.title("Configurer le tableau de bord")
        self.geometry("720x600")
        self.configure(bg=C["bg"])
        self.grab_set()

        tk.Label(self, text="⚙️ Configuration du tableau de bord",
                 font=("Arial", 13, "bold"), bg=C["bg"], fg=C["header"]).pack(pady=(14, 4))
        tk.Label(self,
                 text="Cochez les widgets et les cartes à afficher. Vous pouvez réorganiser les widgets.",
                 bg=C["bg"], font=("Arial", 9), fg="#666").pack()

        # Boutons EN BAS — packés AVANT le notebook pour qu'ils restent visibles
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(side="bottom", pady=10)
        mk_btn(bf, "💾 Enregistrer", self._save).pack(side="left", padx=8)
        mk_btn(bf, "↺ Réinitialiser", self._reset, color="#888").pack(side="left", padx=8)
        mk_btn(bf, "Annuler", self.destroy, color="#888").pack(side="left", padx=8)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=10)

        # Onglet 1 : widgets
        wf = tk.Frame(nb, bg=C["bg"])
        nb.add(wf, text="  Widgets  ")
        self._build_widgets_panel(wf)

        # Onglet 2 : cartes statistiques
        cf = tk.Frame(nb, bg=C["bg"])
        nb.add(cf, text="  Cartes statistiques  ")
        self._build_cards_panel(cf)

    def _build_widgets_panel(self, parent):
        active = db.get_dashboard_widgets()
        # Réordonner le catalogue : les actifs en premier, dans l'ordre choisi
        ordered_keys = [k for k in active if k in WIDGET_CLASSES]
        for k, _, _ in WIDGET_CATALOG:
            if k not in ordered_keys:
                ordered_keys.append(k)

        self.widget_vars = {}
        self.widget_frames = {}
        self.widget_order = list(ordered_keys)

        tk.Label(parent, text="Cochez les widgets à afficher. Utilisez ▲▼ pour réorganiser.",
                 bg=C["bg"], font=("Arial", 9, "italic"), fg="#666").pack(anchor="w", padx=6, pady=(6, 4))

        self.list_frame = tk.Frame(parent, bg=C["bg"])
        self.list_frame.pack(fill="both", expand=True, padx=6, pady=4)

        self._render_widgets_list(active)

    def _render_widgets_list(self, active):
        for w in self.list_frame.winfo_children():
            w.destroy()
        catalog = {k: (lbl, desc) for k, lbl, desc in WIDGET_CATALOG}
        for idx, key in enumerate(self.widget_order):
            if key not in catalog:
                continue
            lbl, desc = catalog[key]
            row = tk.Frame(self.list_frame, bg=C["bg"], relief="solid", bd=1)
            row.pack(fill="x", pady=2)
            var = tk.IntVar(value=1 if key in active else 0)
            self.widget_vars[key] = var

            tk.Checkbutton(row, variable=var, bg=C["bg"]).pack(side="left", padx=6)
            text_f = tk.Frame(row, bg=C["bg"])
            text_f.pack(side="left", fill="x", expand=True, padx=4, pady=4)
            tk.Label(text_f, text=lbl, font=("Arial", 10, "bold"), bg=C["bg"],
                     anchor="w").pack(anchor="w")
            tk.Label(text_f, text=desc, font=("Arial", 9), bg=C["bg"], fg="#666",
                     anchor="w").pack(anchor="w")

            tk.Button(row, text="▲", bg="#ddd", relief="flat", cursor="hand2",
                      width=2, command=lambda i=idx: self._move_up(i)
                      ).pack(side="right", padx=2)
            tk.Button(row, text="▼", bg="#ddd", relief="flat", cursor="hand2",
                      width=2, command=lambda i=idx: self._move_down(i)
                      ).pack(side="right", padx=2)

    def _move_up(self, idx):
        if idx <= 0: return
        # Sauver les coches actuelles avant re-render
        active = [k for k, v in self.widget_vars.items() if v.get()]
        self.widget_order[idx-1], self.widget_order[idx] = \
            self.widget_order[idx], self.widget_order[idx-1]
        self._render_widgets_list(active)

    def _move_down(self, idx):
        if idx >= len(self.widget_order) - 1: return
        active = [k for k, v in self.widget_vars.items() if v.get()]
        self.widget_order[idx+1], self.widget_order[idx] = \
            self.widget_order[idx], self.widget_order[idx+1]
        self._render_widgets_list(active)

    def _build_cards_panel(self, parent):
        active = db.get_dashboard_cards()
        self.card_vars = {}

        tk.Label(parent, text="Cochez les cartes statistiques à afficher en haut du tableau de bord :",
                 bg=C["bg"], font=("Arial", 9, "italic"), fg="#666").pack(anchor="w", padx=6, pady=(6, 4))

        # Grille 2 colonnes
        grid = tk.Frame(parent, bg=C["bg"])
        grid.pack(fill="both", expand=True, padx=6, pady=4)
        for i, (key, label, accent, _value_color) in enumerate(CARD_CATALOG):
            var = tk.IntVar(value=1 if key in active else 0)
            self.card_vars[key] = var
            # Container avec barre accent à gauche pour rappeler le style des cartes
            row = tk.Frame(grid, bg=C["surface"], relief="solid", bd=1)
            row.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=3)
            tk.Frame(row, bg=accent, width=3).pack(side="left", fill="y")
            tk.Checkbutton(row, text=f"  {label}", variable=var,
                            bg=C["surface"], fg=C["text"], font=F["body_bold"],
                            anchor="w", padx=8, pady=6,
                            selectcolor="white", activebackground=C["surface"]
                            ).pack(side="left", fill="x", expand=True)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

    def _save(self):
        # Widgets : ordre + actifs
        active_widgets = [k for k in self.widget_order if self.widget_vars[k].get()]
        db.set_dashboard_widgets(active_widgets)
        # Cartes
        active_cards = [k for k, v in self.card_vars.items() if v.get()]
        db.set_dashboard_cards(active_cards)
        if self.on_save: self.on_save()
        self.destroy()

    def _reset(self):
        if messagebox.askyesno("Réinitialiser",
                "Restaurer les widgets et cartes par défaut ?"):
            db.set_dashboard_widgets([
                "stats_cards", "urgentes", "activite_recente",
                "garantie_expirante", "par_technicien", "non_notifies"])
            db.set_dashboard_cards([
                "En cours", "Clos", "Facturé", "Total",
                "Urgentes", "Clients", "Moteurs", "Tech."])
            if self.on_save: self.on_save()
            self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# TABLEAU DE BORD MODULABLE
# ══════════════════════════════════════════════════════════════════════════════
class DashboardFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app

        # Header avec bouton config
        bar = tk.Frame(self, bg=C["header"], height=52)
        bar.pack(fill="x"); bar.pack_propagate(False)
        tk.Label(bar, text="Tableau de bord", font=("Arial", 15, "bold"),
                 bg=C["header"], fg="white").pack(side="left", padx=20, pady=6)
        tk.Label(bar, text="   Vue modulable & configurable", font=("Arial", 9),
                 bg=C["header"], fg="#aac4e8").pack(side="left")
        mk_btn(bar, "⚙️ Configurer", self._open_config).pack(side="right", padx=12, pady=10)
        mk_btn(bar, "🔄 Rafraîchir", self.refresh, color="#888").pack(side="right", padx=4, pady=10)

        # Zone scrollable pour les widgets
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=C["bg"])
        win_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                         lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                          lambda e: self.canvas.itemconfig(win_id, width=e.width))

        def _on_wheel(ev):
            try: self.canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            except tk.TclError: pass
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", _on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        self._widgets = {}  # key → widget instance

    def _open_config(self):
        DashboardConfigDialog(self, self.app, on_save=self._rebuild)

    def _rebuild(self):
        """Reconstruit les widgets selon la config actuelle."""
        # Détruire tous les widgets actuels
        for w in self.inner.winfo_children():
            w.destroy()
        self._widgets.clear()
        # Recréer dans l'ordre voulu
        active = db.get_dashboard_widgets()
        for key in active:
            cls = WIDGET_CLASSES.get(key)
            if not cls:
                continue
            wrap = tk.Frame(self.inner, bg=C["bg"])
            wrap.pack(fill="x", padx=20, pady=(8, 4))
            w = cls(wrap, self.app)
            w.pack(fill="x")
            self._widgets[key] = w
        self.refresh()

    def refresh(self):
        # Si pas encore construit, construire
        if not self._widgets:
            self._rebuild()
            return
        # Sinon refresh chaque widget
        for w in self._widgets.values():
            try:
                w.refresh()
            except Exception as e:
                print(f"Erreur refresh widget: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# INTERVENTIONS
# ══════════════════════════════════════════════════════════════════════════════
class InterventionsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Interventions", "   Historique complet")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 Recherche :", bg=C["bg"], font=("Arial",10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=24).pack(side="left", padx=4)

        tk.Label(sf, text="  Statut :", bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(12,4))
        self.statut_var = tk.StringVar(value="Tous")
        ttk.Combobox(sf, textvariable=self.statut_var,
                     values=["Tous"]+STATUTS, width=12, state="readonly").pack(side="left")
        self.statut_var.trace_add("write", lambda *_: self.refresh())

        tk.Label(sf, text="  Urgence :", bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(12,4))
        self.urgence_var = tk.StringVar(value="Toutes")
        ttk.Combobox(sf, textvariable=self.urgence_var,
                     values=["Toutes"]+URGENCES, width=12, state="readonly").pack(side="left")
        self.urgence_var.trace_add("write", lambda *_: self.refresh())

        cols = ("urg","num_bon","date","client","navire","num_serie","type","statut","notif","updated")
        col_defs = [
            ("urg","⚡",70),
            ("num_bon","N° Bon",110), ("date","Date",90),
            ("client","Client",155), ("navire","Navire/Site",125),
            ("num_serie","N° Série",115), ("type","Type",115),
            ("statut","Statut",80),
            ("notif","Notif.",115, "center"),
            ("updated","Modifié (Paris)",115),
        ]
        tf, self.tree = mk_tree(self, cols, col_defs, height=22)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier",        self._modifier).pack(side="left", padx=3)
        mk_btn(af, "📄 Générer bon",     self._generer).pack(side="left", padx=3)
        mk_btn(af, "📁 Dossier",         self._dossier).pack(side="left", padx=3)
        mk_btn(af, "📧 Prévenir client", self._mail_client, color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "📧 Prévenir tech.",  self._mail_tech,   color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "🗑️ Supprimer",       self._supprimer, color=C["danger"]).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        invs = db.get_interventions(statut=self.statut_var.get(),
                                     urgence=self.urgence_var.get(),
                                     search=self.search_var.get())
        self._cache = list(invs)
        rows = []
        urgs = []
        for r in self._cache:
            rows.append((
                r["urgence"], r["num_bon"], r["date_creation"], r["client_nom"] or "",
                r["navire"] or "", r["num_serie"] or "",
                r["type_intervention"], r["statut"],
                notif_icon(row_get(r, "client_notifie"), row_get(r, "tech_notifie")),
                db.fmt_paris_short(r["updated_at"]),
            ))
            urgs.append(r["urgence"])
        fill_tree(self.tree, rows, urgences=urgs)

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez une intervention.")
            return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r:
            BonDialog(self, self.app, inv_id=r["id"], on_save=self.refresh)

    def _generer(self):
        r = self._sel()
        if not r: return
        path = sauvegarder_bon(db.get_intervention(inv_id=r["id"]))
        ouvrir_fichier(path)

    def _dossier(self):
        r = self._sel()
        if not r: return
        d = Path(__file__).parent / "dossiers" / r["num_bon"]
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    def _mail_client(self):
        r = self._sel()
        if not r: return
        inv = db.get_intervention(inv_id=r["id"])
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        if not client or not row_get(client, "email"):
            messagebox.showwarning("Email manquant",
                "Aucun email renseigné pour ce client.")
            return
        path = sauvegarder_bon(inv)
        mailer.email_client(inv, client, moteur, str(path))
        db.mark_notifie(r["id"], "client")
        self.refresh()

    def _mail_tech(self):
        r = self._sel()
        if not r: return
        inv = db.get_intervention(inv_id=r["id"])
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        tech = db.get_technicien_by_nom(row_get(inv, "technicien"))
        email = row_get(tech, "email")
        path = sauvegarder_bon(inv)
        mailer.email_technicien(inv, client, moteur, email, str(path))
        db.mark_notifie(r["id"], "tech")
        self.refresh()

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno("Supprimer", f"Supprimer {r['num_bon']} ?"):
            db.delete_intervention(r["id"])
            self.refresh()


# ══════════════════════════════════════════════════════════════════════════════
# CLIENTS
# ══════════════════════════════════════════════════════════════════════════════
class ClientsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Clients", "   Gestion du fichier clients")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 Recherche :", bg=C["bg"], font=("Arial",10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=30).pack(side="left", padx=4)
        mk_btn(sf, "➕ Nouveau client",
               lambda: ClientDialog(self, self.app, on_save=self.refresh)).pack(side="right")

        cols = ("nom","contact","email","telephone","adresse")
        col_defs = [("nom","Nom",200),("contact","Contact",140),
                    ("email","Email",180),("telephone","Téléphone",110),
                    ("adresse","Adresse",260)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=22)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=4)
        mk_btn(af, "🗑️ Supprimer", self._supprimer, color=C["danger"]).pack(side="left", padx=4)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        cs = db.get_clients(search=self.search_var.get())
        self._cache = list(cs)
        fill_tree(self.tree, [(c["nom"],c["contact"],c["email"],
                               c["telephone"],c["adresse"]) for c in self._cache])

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un client."); return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r: ClientDialog(self, self.app, dict(r), on_save=self.refresh)

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno("Supprimer", f"Supprimer '{r['nom']}' ?"):
            db.delete_client(r["id"]); self.refresh()


# ══════════════════════════════════════════════════════════════════════════════
# MOTEURS
# ══════════════════════════════════════════════════════════════════════════════
class MoteursFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Moteurs", "   Recherche par N° de série")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 N° série / Navire / Machine :",
                 bg=C["bg"], font=("Arial",10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=30).pack(side="left", padx=4)
        # Boutons à droite
        mk_btn(sf, "➕ Nouveau moteur",
               lambda: MoteurDialog(self, self.app, on_save=self.refresh)).pack(side="right")
        mk_btn(sf, "📥 Importer CSV",
               lambda: ImportCSVDialog(self, self.app, on_save=self.refresh),
               color=C["btn2"]).pack(side="right", padx=(0, 6))

        cols = ("num_serie","client","navire","marque","machine","type_moteur","date","garantie","statut_g")
        col_defs = [("num_serie","N° Série",125),("client","Client",140),
                    ("navire","Navire/Site",120),
                    ("marque","Marque",90),
                    ("machine","Machine",100),
                    ("type_moteur","Type moteur",115),("date","Mise en service",105),
                    ("garantie","Mois",50, "center"),("statut_g","Garantie",115)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=22)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=4)
        mk_btn(af, "📋 Voir interventions", self._voir_inv).pack(side="left", padx=4)
        mk_btn(af, "🗑️ Supprimer", self._supprimer, color=C["danger"]).pack(side="left", padx=4)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        ms = db.get_moteurs(search=self.search_var.get())
        self._cache = list(ms)
        rows = []
        for m in self._cache:
            statut, jours = db.garantie_status(m["date_mise_service"], m["duree_garantie"])
            if statut == "Active":   gtxt = f"Active · {jours}j"
            elif statut == "Expirée": gtxt = f"Expirée · -{jours}j"
            else: gtxt = "—"
            rows.append((m["num_serie"], m["client_nom"] or "",
                         row_get(m, "navire"),
                         row_get(m, "marque"),
                         row_get(m, "machine"),
                         row_get(m, "type_moteur"),
                         m["date_mise_service"], m["duree_garantie"], gtxt))
        fill_tree(self.tree, rows)

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un moteur."); return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r: MoteurDialog(self, self.app, dict(r), on_save=self.refresh)

    def _voir_inv(self):
        r = self._sel()
        if not r: return
        self.app.show("interventions")
        self.app.frames["interventions"].search_var.set(r["num_serie"])
        self.app.frames["interventions"].statut_var.set("Tous")

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno("Supprimer", f"Supprimer '{r['num_serie']}' ?"):
            db.delete_moteur(r["id"]); self.refresh()


# ══════════════════════════════════════════════════════════════════════════════
# TECHNICIENS
# ══════════════════════════════════════════════════════════════════════════════
class TechniciensFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Techniciens", "   Annuaire interne EMS")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        mk_btn(sf, "➕ Nouveau technicien",
               lambda: TechnicienDialog(self, self.app, on_save=self.refresh)).pack(side="right")

        # Alignement amélioré : colonnes cadrées à gauche avec padding correct
        cols = ("nom","email","tel","interventions")
        col_defs = [
            ("nom",           "Nom",              240, "w"),
            ("email",         "Email",            260, "w"),
            ("tel",           "Téléphone",        140, "w"),
            ("interventions", "Interventions",     90, "center"),
        ]
        tf, self.tree = mk_tree(self, cols, col_defs, height=22)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=4)
        mk_btn(af, "🗑️ Supprimer", self._supprimer, color=C["danger"]).pack(side="left", padx=4)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        ts = db.get_techniciens()
        self._cache = list(ts)
        # Compter les interventions de chaque tech
        tech_stats = {r["technicien"]: r["total"] for r in db.get_stats_par_technicien()}
        rows = [(t["nom"], t["email"], t["telephone"], tech_stats.get(t["nom"], 0))
                for t in self._cache]
        fill_tree(self.tree, rows)

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un technicien."); return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r: TechnicienDialog(self, self.app, dict(r), on_save=self.refresh)

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno("Supprimer", f"Supprimer '{r['nom']}' ?"):
            db.delete_technicien(r["id"]); self.refresh()


# ══════════════════════════════════════════════════════════════════════════════
# NOUVEAU
# ══════════════════════════════════════════════════════════════════════════════
class NouveauFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Nouveau bon d'intervention")
        inner = tk.Frame(self, bg=C["bg"])
        inner.pack(expand=True)
        tk.Label(inner, text="➕", font=("Arial",48), bg=C["bg"], fg=C["header"]).pack(pady=(40,10))
        tk.Label(inner, text="Créer un nouveau bon d'intervention",
                 font=("Arial",14), bg=C["bg"]).pack()
        mk_btn(inner, "Ouvrir le formulaire",
               lambda: BonDialog(self, self.app,
                                 on_save=lambda: self.app.show("dashboard"))
               ).pack(pady=20, ipadx=20, ipady=8)

    def refresh(self): pass


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET TABLEAU MATÉRIELS
# ══════════════════════════════════════════════════════════════════════════════
class MaterielsTable(tk.Frame):
    """Tableau éditable pour les matériels utilisés."""
    COLS = [("qte", "Quantité", 8),
            ("ref", "Référence", 18),
            ("designation", "Désignation", 40)]

    def __init__(self, master, **kw):
        super().__init__(master, bg=C["bg"], **kw)
        self._rows = []
        self.head = tk.Frame(self, bg=C["header"])
        self.head.pack(fill="x", pady=(0, 1))
        for key, lbl, w in self.COLS:
            tk.Label(self.head, text=lbl, bg=C["header"], fg="white",
                     font=("Arial", 9, "bold"),
                     width=w, anchor="w", padx=4).pack(side="left", padx=1)
        tk.Label(self.head, text="", bg=C["header"], width=4).pack(side="left", padx=1)
        self.body = tk.Frame(self, bg=C["bg"])
        self.body.pack(fill="x")
        self.bf = tk.Frame(self, bg=C["bg"])
        self.bf.pack(fill="x", pady=(4, 0))
        mk_btn(self.bf, "➕ Ajouter une ligne", self.add_row, color=C["btn2"]).pack(side="left")
        for _ in range(3): self.add_row()

    def add_row(self, qte="", ref="", designation=""):
        rf = tk.Frame(self.body, bg=C["bg"]); rf.pack(fill="x", pady=1)
        v_qte = tk.StringVar(value=qte); v_ref = tk.StringVar(value=ref)
        v_des = tk.StringVar(value=designation)
        ttk.Entry(rf, textvariable=v_qte, width=8).pack(side="left", padx=1)
        ttk.Entry(rf, textvariable=v_ref, width=18).pack(side="left", padx=1)
        ttk.Entry(rf, textvariable=v_des, width=40).pack(side="left", padx=1, fill="x", expand=True)
        row = {"qte": v_qte, "ref": v_ref, "designation": v_des, "frame": rf}
        tk.Button(rf, text="✕", bg="#ddd", fg=C["danger"],
                  font=("Arial", 9, "bold"), relief="flat", cursor="hand2",
                  padx=4, pady=0, command=lambda r=row: self._remove(r)
                  ).pack(side="left", padx=2)
        self._rows.append(row)

    def _remove(self, row):
        if len(self._rows) <= 1:
            row["qte"].set(""); row["ref"].set(""); row["designation"].set("")
            return
        row["frame"].destroy()
        self._rows.remove(row)

    def load(self, items):
        for r in self._rows: r["frame"].destroy()
        self._rows.clear()
        if not items:
            for _ in range(3): self.add_row()
        else:
            for item in items:
                self.add_row(qte=str(item.get("qte","")),
                             ref=str(item.get("ref","")),
                             designation=str(item.get("designation","")))

    def to_list(self):
        out = []
        for r in self._rows:
            qte  = r["qte"].get().strip()
            ref  = r["ref"].get().strip()
            desg = r["designation"].get().strip()
            if qte or ref or desg:
                out.append({"qte": qte, "ref": ref, "designation": desg})
        return out


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET DÉPLACEMENTS
# ══════════════════════════════════════════════════════════════════════════════
class DeplacementsTable(tk.Frame):
    LEFT = [
        ("trajet_aller",       "Temps de trajet aller"),
        ("heure_debut_matin",  "Heure début intervention matin"),
        ("heure_fin_matin",    "Heure fin intervention matin"),
        ("heure_debut_apres",  "Heure début intervention après-midi"),
        ("heure_fin_apres",    "Heure fin intervention après-midi"),
        ("trajet_retour",      "Temps de trajet retour"),
    ]
    RIGHT = [
        ("frais_repas",        "Frais de repas"),
        ("frais_hotel",        "Frais d'hôtel"),
        ("frais_peages",       "Frais de péages"),
        ("temps_preparation",  "Temps de préparation"),
        ("temps_rangement",    "Temps de rangement"),
        ("",                   ""),
    ]

    def __init__(self, master, **kw):
        super().__init__(master, bg=C["bg"], **kw)
        self.vars = {}
        for i in range(len(self.LEFT)):
            kl, ll = self.LEFT[i]
            kr, lr = self.RIGHT[i]
            row = tk.Frame(self, bg=C["bg"]); row.pack(fill="x", pady=1)
            if kl:
                tk.Label(row, text=ll, bg=C["bg"], font=("Arial", 9),
                         width=30, anchor="w").pack(side="left", padx=(0, 4))
                self.vars[kl] = tk.StringVar()
                ttk.Entry(row, textvariable=self.vars[kl], width=14).pack(side="left", padx=(0, 12))
            if kr:
                tk.Label(row, text=lr, bg=C["bg"], font=("Arial", 9),
                         width=24, anchor="w").pack(side="left", padx=(8, 4))
                self.vars[kr] = tk.StringVar()
                ttk.Entry(row, textvariable=self.vars[kr], width=14).pack(side="left")

    def load(self, data):
        if not data: data = {}
        for k, var in self.vars.items():
            var.set(str(data.get(k, "")))

    def to_dict(self):
        return {k: v.get().strip() for k, v in self.vars.items() if v.get().strip()}


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG CLIENT
# ══════════════════════════════════════════════════════════════════════════════
class ClientDialog(tk.Toplevel):
    def __init__(self, parent, app, client=None, on_save=None):
        super().__init__(parent)
        self.app = app; self.client = client; self.on_save = on_save
        self.title("Nouveau client" if not client else f"Modifier – {client['nom']}")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.v = {k: tk.StringVar(value=client.get(k,"") if client else "")
                  for k in ["nom","contact","email","telephone","adresse"]}
        f = tk.Frame(self, bg=C["bg"])
        f.pack(padx=24, pady=18)
        for i,(key,lbl) in enumerate([("nom","Nom *"),("contact","Contact"),
                                       ("email","Email"),("telephone","Téléphone"),
                                       ("adresse","Adresse")]):
            tk.Label(f, text=lbl, bg=C["bg"], font=("Arial",10),
                     anchor="e", width=14).grid(row=i, column=0, sticky="e", padx=(0,6), pady=4)
            if key == "email":
                EmailEntry(f, textvariable=self.v[key], width=36).grid(row=i, column=1, pady=4)
            else:
                ttk.Entry(f, textvariable=self.v[key], width=36).grid(row=i, column=1, pady=4)
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(pady=12)
        mk_btn(bf,"💾 Enregistrer",self._save).pack(side="left",padx=8)
        mk_btn(bf,"Annuler",self.destroy,color="#888").pack(side="left",padx=8)

    def _save(self):
        nom = self.v["nom"].get().strip()
        if not nom:
            messagebox.showwarning("Champ requis","Le nom est obligatoire.")
            return
        # Email : avertissement informatif si format douteux, mais pas bloquant
        email = self.v["email"].get().strip()
        if email and not db.email_looks_valid(email):
            if not messagebox.askyesno("Format email douteux",
                    f"L'email '{email}' semble mal formé.\n\n"
                    "Voulez-vous quand même l'enregistrer tel quel ?"):
                return
        db.upsert_client({k:self.v[k].get().strip() for k in self.v},
                         client_id=self.client["id"] if self.client else None)
        if self.on_save: self.on_save()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG MOTEUR
# ══════════════════════════════════════════════════════════════════════════════
class MoteurDialog(tk.Toplevel):
    FIELDS = [
        ("num_serie",         "N° Série *"),
        ("navire",            "Navire / Site"),
        ("machine",           "Machine"),
        ("type_moteur",       "Type moteur / inverseur"),
        ("marque",            "Marque"),
        ("ref_constructeur",  "Réf. Constructeur"),
        ("cylindree",         "Cylindrée"),
        ("famille",           "Famille"),
        ("application",       "Application"),
        ("typologie",         "Typologie"),
        ("collection",        "Collection"),
        ("code_affaire",      "Code Affaire"),
        ("date_mise_service", "Mise en service"),
        ("duree_garantie",    "Garantie (mois)"),
    ]

    def __init__(self, parent, app, moteur=None, on_save=None):
        super().__init__(parent)
        self.app = app; self.moteur = moteur; self.on_save = on_save
        self.title("Nouveau moteur" if not moteur else f"Modifier – {moteur['num_serie']}")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.grab_set()
        self._clients = list(db.get_clients())
        self.client_var = tk.StringVar()
        if moteur and moteur.get("client_id"):
            c = next((x for x in self._clients if x["id"]==moteur["client_id"]),None)
            if c: self.client_var.set(c["nom"])
        f = tk.Frame(self, bg=C["bg"])
        f.pack(padx=24, pady=18)
        tk.Label(f, text="Client *", bg=C["bg"], font=("Arial",10),
                 anchor="e", width=22).grid(row=0, column=0, sticky="e", padx=(0,6), pady=4)
        ttk.Combobox(f, textvariable=self.client_var,
                     values=[c["nom"] for c in self._clients],
                     width=33, state="readonly").grid(row=0, column=1, pady=4)
        self.v = {k: tk.StringVar(value=(moteur.get(k,"") if moteur else ""))
                  for k, _ in self.FIELDS}
        for i,(key,lbl) in enumerate(self.FIELDS, start=1):
            tk.Label(f, text=lbl, bg=C["bg"], font=("Arial",10),
                     anchor="e", width=22).grid(row=i, column=0, sticky="e", padx=(0,6), pady=4)
            if key == "date_mise_service":
                DateEntry(f, textvariable=self.v[key], width=32).grid(row=i, column=1, pady=4)
            else:
                ttk.Entry(f, textvariable=self.v[key], width=34).grid(row=i, column=1, pady=4)
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(pady=12)
        mk_btn(bf,"💾 Enregistrer",self._save).pack(side="left",padx=8)
        mk_btn(bf,"Annuler",self.destroy,color="#888").pack(side="left",padx=8)

    def _save(self):
        ns = self.v["num_serie"].get().strip()
        cn = self.client_var.get()
        if not ns or not cn:
            messagebox.showwarning("Champs requis","N° de série et client sont obligatoires.")
            return
        d_svc = self.v["date_mise_service"].get().strip()
        if d_svc and not DateEntry.is_valid(d_svc):
            messagebox.showwarning("Date invalide",
                f"Date de mise en service invalide : '{d_svc}'\nFormat attendu : JJ/MM/AAAA")
            return
        client = next((c for c in self._clients if c["nom"]==cn),None)
        db.upsert_moteur({"client_id":client["id"] if client else "",
                          **{k:self.v[k].get().strip() for k in self.v}},
                         moteur_id=self.moteur["id"] if self.moteur else None)
        if self.on_save: self.on_save()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG TECHNICIEN
# ══════════════════════════════════════════════════════════════════════════════
class TechnicienDialog(tk.Toplevel):
    def __init__(self, parent, app, tech=None, on_save=None):
        super().__init__(parent)
        self.app = app; self.tech = tech; self.on_save = on_save
        self.title("Nouveau technicien" if not tech else f"Modifier – {tech['nom']}")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.v = {k: tk.StringVar(value=tech.get(k,"") if tech else "")
                  for k in ["nom","email","telephone"]}
        f = tk.Frame(self, bg=C["bg"]); f.pack(padx=24, pady=18)
        for i,(key,lbl) in enumerate([("nom","Nom *"),("email","Email"),("telephone","Téléphone")]):
            tk.Label(f, text=lbl, bg=C["bg"], font=("Arial",10),
                     anchor="e", width=14).grid(row=i, column=0, sticky="e", padx=(0,6), pady=4)
            if key == "email":
                EmailEntry(f, textvariable=self.v[key], width=36).grid(row=i, column=1, pady=4)
            else:
                ttk.Entry(f, textvariable=self.v[key], width=36).grid(row=i, column=1, pady=4)
        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=12)
        mk_btn(bf,"💾 Enregistrer",self._save).pack(side="left",padx=8)
        mk_btn(bf,"Annuler",self.destroy,color="#888").pack(side="left",padx=8)

    def _save(self):
        nom = self.v["nom"].get().strip()
        if not nom: messagebox.showwarning("Champ requis","Le nom est obligatoire."); return
        email = self.v["email"].get().strip()
        if email and not db.email_looks_valid(email):
            if not messagebox.askyesno("Format email douteux",
                    f"L'email '{email}' semble mal formé.\n\n"
                    "Voulez-vous quand même l'enregistrer ?"):
                return
        db.upsert_technicien({k:self.v[k].get().strip() for k in self.v},
                             tech_id=self.tech["id"] if self.tech else None)
        if self.on_save: self.on_save()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG TYPES D'INTERVENTION
# ══════════════════════════════════════════════════════════════════════════════
class TypesDialog(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Types d'intervention")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.geometry("400x460")
        self.grab_set()

        tk.Label(self, text="Types d'intervention configurables",
                 bg=C["bg"], font=("Arial",11,"bold"), fg=C["header"]).pack(pady=(14,4))
        tk.Label(self,
                 text="Les 4 types officiels (Entretien, Dépannage, Diagnostic,\n"
                      "Garantie) apparaissent comme cases à cocher sur le bon.",
                 bg=C["bg"], font=("Arial",9), fg="#666", justify="center").pack()

        lf = tk.Frame(self, bg=C["bg"]); lf.pack(fill="both", expand=True, padx=16, pady=10)
        self.lst = tk.Listbox(lf, font=("Arial",10), height=12,
                              selectbackground=C["btn"], selectforeground="white")
        self.lst.pack(fill="both", expand=True)

        af = tk.Frame(self, bg=C["bg"]); af.pack(fill="x", padx=16, pady=(0,6))
        self.input_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.input_var, width=28).pack(side="left", padx=(0,4))
        mk_btn(af, "➕ Ajouter", self._add).pack(side="left")

        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=10)
        mk_btn(bf, "✏️ Renommer", self._rename).pack(side="left", padx=4)
        mk_btn(bf, "🗑️ Supprimer", self._del, color=C["danger"]).pack(side="left", padx=4)
        mk_btn(bf, "Fermer", self.destroy, color="#888").pack(side="left", padx=4)
        self._refresh()

    def _refresh(self):
        self.lst.delete(0, "end")
        for t in db.get_types_intervention():
            self.lst.insert("end", t)

    def _add(self):
        v = self.input_var.get().strip()
        if not v: return
        if not db.add_type_intervention(v):
            messagebox.showwarning("Doublon", f"Le type '{v}' existe déjà.")
            return
        self.input_var.set("")
        self._refresh()

    def _selected(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un type.")
            return None
        return self.lst.get(sel[0])

    def _rename(self):
        cur = self._selected()
        if not cur: return
        new = simpledialog.askstring("Renommer", f"Nouveau nom pour '{cur}' :",
                                      parent=self, initialvalue=cur)
        if new and new.strip() and new != cur:
            if db.update_type_intervention(cur, new.strip()):
                self._refresh()
            else:
                messagebox.showwarning("Erreur", "Le nouveau nom existe déjà.")

    def _del(self):
        cur = self._selected()
        if not cur: return
        if messagebox.askyesno("Supprimer",
                f"Supprimer le type '{cur}' ?\n\n(Les bons existants conservent leur libellé.)"):
            db.delete_type_intervention(cur)
            self._refresh()


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG BON D'INTERVENTION (formulaire complet)
# ══════════════════════════════════════════════════════════════════════════════
class BonDialog(tk.Toplevel):
    """Formulaire création / modification bon, conforme au modèle EMS."""

    MOTEUR_INFO_FIELDS = [
        ("navire",            "Navire / Site"),
        ("machine",           "Machine"),
        ("type_moteur",       "Type moteur / inverseur"),
        ("date_mise_service", "Mise en service"),
        ("duree_garantie",    "Garantie (mois)"),
    ]

    def __init__(self, parent, app, inv_id=None, on_save=None):
        super().__init__(parent)
        self.app     = app
        self.inv_id  = inv_id
        self.on_save = on_save
        self.is_edit = inv_id is not None
        self.title("Modifier bon" if self.is_edit else "Nouveau bon d'intervention")
        self.geometry("980x830")
        self.configure(bg=C["bg"])
        self.grab_set()

        self._clients     = list(db.get_clients())
        self._all_moteurs = list(db.get_moteurs())
        self._techniciens = list(db.get_techniciens())

        self._client_by_id  = {c["id"]:  c for c in self._clients}
        self._client_by_nom = {c["nom"]: c for c in self._clients}
        self._moteur_by_id  = {m["id"]:  m for m in self._all_moteurs}
        self._moteur_by_ns  = {m["num_serie"]: m for m in self._all_moteurs}

        # Cases à cocher : options + classifications
        self.chk = {k: tk.IntVar(value=0) for k in [
            "outil_diagnostic", "memoriser_avant", "memoriser_apres",
            "photos_avant", "photos_apres", "pour_information", "preconisation",
            "garantie_intervention", "facturable", "interne",
        ]}

        self._build()

        if self.is_edit:
            self.after(50, self._load)

    def _build(self):
        canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        p = tk.Frame(canvas, bg=C["bg"])
        win_id = canvas.create_window((0, 0), window=p, anchor="nw")
        p.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))

        def _on_wheel(ev):
            try: canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            except tk.TclError: pass
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        self.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ── CLIENT ────────────────────────────────────────────────────────────
        section_bar(p, "CLIENT")
        row = tk.Frame(p, bg=C["bg"])
        row.pack(fill="x", padx=20, pady=8)
        tk.Label(row, text="Client *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.client_var = tk.StringVar()
        self.client_combo = SearchableCombobox(row, textvariable=self.client_var,
            values=[c["nom"] for c in self._clients], width=44)
        self.client_combo.pack(side="left")
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_selected)
        mk_btn(row, "+", lambda: ClientDialog(self, self.app, on_save=self._reload_refs),
               color=C["btn2"]).pack(side="left", padx=(8,0))

        self.lieu_var       = tk.StringVar()
        self.signataire_var = tk.StringVar()
        self.email_sig_var  = tk.StringVar()
        for lbl, var, cls in [("Lieu de l'intervention", self.lieu_var,       ttk.Entry),
                                ("Nom du signataire",      self.signataire_var, ttk.Entry),
                                ("Courriel du signataire", self.email_sig_var,  EmailEntry)]:
            r = tk.Frame(p, bg=C["bg"])
            r.pack(fill="x", padx=20, pady=2)
            tk.Label(r, text=lbl, bg=C["bg"], font=("Arial",10),
                     width=24, anchor="e").pack(side="left", padx=(0,8))
            cls(r, textvariable=var, width=46).pack(side="left")

        # ── ÉQUIPEMENT ────────────────────────────────────────────────────────
        section_bar(p, "ÉQUIPEMENT")
        row2 = tk.Frame(p, bg=C["bg"])
        row2.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(row2, text="N° Série *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.moteur_var = tk.StringVar()
        self.moteur_combo = SearchableCombobox(row2, textvariable=self.moteur_var,
            values=[m["num_serie"] for m in self._all_moteurs], width=44)
        self.moteur_combo.pack(side="left")
        self.moteur_combo.bind("<<ComboboxSelected>>", self._on_moteur_selected)
        mk_btn(row2, "+", lambda: MoteurDialog(self, self.app, on_save=self._reload_refs),
               color=C["btn2"]).pack(side="left", padx=(8,0))

        self._info_lbls = {}
        for key, txt in self.MOTEUR_INFO_FIELDS:
            r = tk.Frame(p, bg=C["bg"])
            r.pack(fill="x", padx=20, pady=2)
            tk.Label(r, text=txt, bg=C["bg"], font=("Arial",10),
                     width=24, anchor="e").pack(side="left", padx=(0,8))
            lw = tk.Label(r, text="", bg="#eef2f7", font=("Arial",10),
                          anchor="w", relief="groove", width=46, padx=4)
            lw.pack(side="left", ipady=3)
            self._info_lbls[key] = lw

        rh = tk.Frame(p, bg=C["bg"]); rh.pack(fill="x", padx=20, pady=2)
        tk.Label(rh, text="Nb heures de fonctionnement", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.nb_heures_var = tk.StringVar()
        ttk.Entry(rh, textvariable=self.nb_heures_var, width=46).pack(side="left")

        rg = tk.Frame(p, bg=C["bg"]); rg.pack(fill="x", padx=20, pady=2)
        tk.Label(rg, text="Statut garantie", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.garantie_lbl = tk.Label(rg, text="—", bg="#eef2f7", font=("Arial",10,"bold"),
                                      anchor="w", relief="groove", width=46, padx=4)
        self.garantie_lbl.pack(side="left", ipady=3)

        # ── DÉTAILS INTERVENTION ──────────────────────────────────────────────
        section_bar(p, "DÉTAILS INTERVENTION")
        self.type_var    = tk.StringVar()
        self.tech_var    = tk.StringVar()
        self.date_var    = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.statut_var  = tk.StringVar(value="En cours")
        self.urgence_var = tk.StringVar(value="Normale")

        # Type d'intervention
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Type d'intervention *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        ttk.Combobox(r, textvariable=self.type_var, values=db.get_types_intervention(),
                     width=44, state="readonly").pack(side="left")
        mk_btn(r, "⚙", lambda: TypesDialog(self, self.app), color=C["btn2"]).pack(side="left", padx=(8,0))

        # Urgence
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Urgence *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        urg_combo = ttk.Combobox(r, textvariable=self.urgence_var, values=URGENCES,
                                  width=44, state="readonly")
        urg_combo.pack(side="left")
        # Couleur du label selon urgence
        self.urg_indicator = tk.Label(r, text="", bg=C["bg"], font=("Arial", 14, "bold"))
        self.urg_indicator.pack(side="left", padx=(8, 0))
        self.urgence_var.trace_add("write", lambda *_: self._update_urg_indicator())
        self._update_urg_indicator()

        # Classifications (3 cases indépendantes)
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Classifications", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        cls_frame = tk.Frame(r, bg=C["bg"])
        cls_frame.pack(side="left")
        tk.Checkbutton(cls_frame, text="Garantie", variable=self.chk["garantie_intervention"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 12))
        tk.Checkbutton(cls_frame, text="Facturable", variable=self.chk["facturable"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 12))
        tk.Checkbutton(cls_frame, text="Interne", variable=self.chk["interne"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left")

        # Technicien
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Technicien *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.tech_combo = SearchableCombobox(r, textvariable=self.tech_var,
            values=[t["nom"] for t in self._techniciens], width=44)
        self.tech_combo.pack(side="left")
        mk_btn(r, "+", lambda: TechnicienDialog(self, self.app, on_save=self._reload_refs),
               color=C["btn2"]).pack(side="left", padx=(8,0))

        # Date intervention
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Date intervention", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        DateEntry(r, textvariable=self.date_var, width=44).pack(side="left")
        tk.Label(r, text="  (JJ/MM/AAAA)", bg=C["bg"], fg="#888",
                 font=("Arial", 8)).pack(side="left")

        # Statut
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Statut", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        ttk.Combobox(r, textvariable=self.statut_var, values=STATUTS,
                     width=44, state="readonly").pack(side="left")

        # ── DEMANDE DU CLIENT ─────────────────────────────────────────────────
        section_bar(p, "DEMANDE DU CLIENT *")
        self.txt_demande = tk.Text(p, height=4, font=("Arial",10), wrap="word",
                                    relief="solid", bd=1, padx=6, pady=4)
        self.txt_demande.pack(fill="x", padx=20, pady=(4,4))

        # ── OPTIONS ───────────────────────────────────────────────────────────
        op = tk.Frame(p, bg=C["bg"])
        op.pack(fill="x", padx=20, pady=4)
        tk.Checkbutton(op, text="Utilisation de l'Outil de diagnostic",
                       variable=self.chk["outil_diagnostic"], bg=C["bg"],
                       font=("Arial",10)).pack(side="left", padx=(0, 16))

        op2 = tk.Frame(p, bg=C["bg"]); op2.pack(fill="x", padx=20, pady=2)
        tk.Label(op2, text="Mémoriser les données :", bg=C["bg"],
                 font=("Arial",10,"bold")).pack(side="left", padx=(0, 8))
        tk.Checkbutton(op2, text="avant", variable=self.chk["memoriser_avant"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 8))
        tk.Checkbutton(op2, text="après", variable=self.chk["memoriser_apres"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 24))
        tk.Label(op2, text="PHOTOS :", bg=C["bg"],
                 font=("Arial",10,"bold")).pack(side="left", padx=(0, 8))
        tk.Checkbutton(op2, text="avant", variable=self.chk["photos_avant"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 8))
        tk.Checkbutton(op2, text="après", variable=self.chk["photos_apres"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left")

        # ── CONSTAT ───────────────────────────────────────────────────────────
        section_bar(p, "CONSTAT AVANT INTERVENTION")
        self.txt_constat = tk.Text(p, height=5, font=("Arial",10), wrap="word",
                                    relief="solid", bd=1, padx=6, pady=4)
        self.txt_constat.pack(fill="x", padx=20, pady=(4,4))

        # ── TRAVAUX ───────────────────────────────────────────────────────────
        section_bar(p, "TRAVAUX RÉALISÉS")
        self.txt_travaux = tk.Text(p, height=5, font=("Arial",10), wrap="word",
                                    relief="solid", bd=1, padx=6, pady=4)
        self.txt_travaux.pack(fill="x", padx=20, pady=(4,4))

        # ── MATÉRIELS ─────────────────────────────────────────────────────────
        section_bar(p, "MATÉRIELS UTILISÉS (Quantité / Référence / Désignation)")
        self.materiels_tbl = MaterielsTable(p)
        self.materiels_tbl.pack(fill="x", padx=20, pady=(4,6))

        info_f = tk.Frame(p, bg=C["bg"])
        info_f.pack(fill="x", padx=20, pady=(2, 4))
        tk.Checkbutton(info_f, text="Pour information", variable=self.chk["pour_information"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 16))
        tk.Checkbutton(info_f, text="Préconisation", variable=self.chk["preconisation"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left")
        self.txt_info = tk.Text(p, height=3, font=("Arial",10), wrap="word",
                                 relief="solid", bd=1, padx=6, pady=4)
        self.txt_info.pack(fill="x", padx=20, pady=(2, 6))

        # ── DÉPLACEMENTS ──────────────────────────────────────────────────────
        section_bar(p, "DÉPLACEMENTS – Temps & frais")
        self.depl_tbl = DeplacementsTable(p)
        self.depl_tbl.pack(fill="x", padx=20, pady=(4, 8))

        # ── DOSSIER ───────────────────────────────────────────────────────────
        section_bar(p, "DOSSIER – FICHIERS & PHOTOS")
        df = tk.Frame(p, bg=C["bg"])
        df.pack(fill="x", padx=20, pady=(4,6))
        sub = tk.Frame(df, bg=C["bg"]); sub.pack(fill="x")
        mk_btn(sub, "📎 Ajouter fichier(s)", self._ajouter_fichiers, color=C["btn2"]).pack(side="left", padx=(0,4))
        mk_btn(sub, "📁 Ouvrir dossier", self._ouvrir_dossier).pack(side="left", padx=4)
        mk_btn(sub, "🔄 Rafraîchir liste", self._refresh_files).pack(side="left", padx=4)
        self._files_lbl = tk.Label(df, text="(Bon non encore enregistré)",
                                    bg=C["bg"], fg="#666", font=("Arial",9), justify="left",
                                    anchor="w")
        self._files_lbl.pack(fill="x", pady=(6,0))

        # ── BOUTONS ───────────────────────────────────────────────────────────
        bf = tk.Frame(p, bg=C["bg"])
        bf.pack(pady=14)
        mk_btn(bf, "💾 Enregistrer",            lambda: self._save(generer=False)).pack(side="left", padx=6)
        mk_btn(bf, "📄 Enregistrer + Générer",  lambda: self._save(generer=True)).pack(side="left", padx=6)
        mk_btn(bf, "📧 Prévenir client",        self._mail_client_inline, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "📧 Prévenir technicien",    self._mail_tech_inline,   color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "Annuler", self.destroy, color="#888").pack(side="left", padx=6)

    def _update_urg_indicator(self):
        u = self.urgence_var.get()
        if u == "Critique":
            self.urg_indicator.config(text="⚡ Critique", fg=C["urg_critique"])
        elif u == "Urgente":
            self.urg_indicator.config(text="⚡ Urgente", fg=C["urg_urgente"])
        else:
            self.urg_indicator.config(text="●  Normale", fg=C["urg_normale"])

    # ─── Référentiels ─────────────────────────────────────────────────────────
    def _reload_refs(self):
        old_client = self.client_var.get()
        old_moteur = self.moteur_var.get()
        old_tech   = self.tech_var.get()

        self._clients     = list(db.get_clients())
        self._all_moteurs = list(db.get_moteurs())
        self._techniciens = list(db.get_techniciens())
        self._client_by_id  = {c["id"]:  c for c in self._clients}
        self._client_by_nom = {c["nom"]: c for c in self._clients}
        self._moteur_by_id  = {m["id"]:  m for m in self._all_moteurs}
        self._moteur_by_ns  = {m["num_serie"]: m for m in self._all_moteurs}

        self.client_combo.set_values([c["nom"] for c in self._clients])
        self.tech_combo.set_values([t["nom"] for t in self._techniciens])
        cli = self._client_by_nom.get(old_client)
        filtrés = ([m for m in self._all_moteurs if m["client_id"] == cli["id"]]
                   if cli else self._all_moteurs)
        self.moteur_combo.set_values([m["num_serie"] for m in filtrés])
        self.client_var.set(old_client)
        self.moteur_var.set(old_moteur)
        self.tech_var.set(old_tech)

    # ─── Événements ───────────────────────────────────────────────────────────
    def _on_client_selected(self, *_):
        client = self._client_by_nom.get(self.client_var.get())
        if client:
            if not self.signataire_var.get():
                self.signataire_var.set(row_get(client, "contact"))
            if not self.email_sig_var.get():
                self.email_sig_var.set(row_get(client, "email"))
        filtrés = ([m for m in self._all_moteurs if m["client_id"] == client["id"]]
                   if client else self._all_moteurs)
        self.moteur_combo.set_values([m["num_serie"] for m in filtrés])
        self.moteur_var.set("")
        for lw in self._info_lbls.values():
            lw.config(text="")
        self.garantie_lbl.config(text="—", fg="black")

    def _on_moteur_selected(self, *_):
        m = self._moteur_by_ns.get(self.moteur_var.get())
        if m:
            self._set_moteur_info(m)
            if not self.client_var.get():
                c = self._client_by_id.get(m["client_id"])
                if c:
                    self.client_combo.set(c["nom"])
                    if not self.signataire_var.get():
                        self.signataire_var.set(row_get(c, "contact"))
                    if not self.email_sig_var.get():
                        self.email_sig_var.set(row_get(c, "email"))

    def _set_moteur_info(self, m):
        for key, lw in self._info_lbls.items():
            lw.config(text=str(row_get(m, key, "")))
        statut, jours = db.garantie_status(row_get(m, "date_mise_service"),
                                           row_get(m, "duree_garantie"))
        if statut == "Active":
            self.garantie_lbl.config(text=f"✅ Active – {jours} jour(s) restant(s)", fg="#0f5132")
            # Auto-cocher classification Garantie si garantie active
            if not self.chk["garantie_intervention"].get():
                self.chk["garantie_intervention"].set(1)
        elif statut == "Expirée":
            self.garantie_lbl.config(text=f"❌ Expirée depuis {jours} jour(s)", fg=C["danger"])
        else:
            self.garantie_lbl.config(text="—  (date ou durée non renseignée)", fg="#666")

    # ─── Fichiers ─────────────────────────────────────────────────────────────
    def _current_dossier(self):
        if not self.is_edit or not self.inv_id:
            return None
        inv = db.get_intervention(inv_id=self.inv_id)
        if not inv: return None
        return Path(__file__).parent / "dossiers" / inv["num_bon"]

    def _refresh_files(self):
        d = self._current_dossier()
        if not d:
            self._files_lbl.config(text="(Bon non encore enregistré – les fichiers seront classés "
                                         "dans dossiers/BON-AAAA-XXXX/ après sauvegarde)")
            return
        d.mkdir(parents=True, exist_ok=True)
        files = sorted([f.name for f in d.iterdir() if f.is_file()])
        if not files:
            self._files_lbl.config(text=f"📁 {d.name} (vide)")
        else:
            txt = f"📁 {d.name} – {len(files)} fichier(s) :\n  • " + "\n  • ".join(files[:8])
            if len(files) > 8:
                txt += f"\n  ... et {len(files)-8} autre(s)"
            self._files_lbl.config(text=txt)

    def _ajouter_fichiers(self):
        d = self._current_dossier()
        if not d:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon. Le dossier sera créé automatiquement.")
            return
        files = filedialog.askopenfilenames(
            title="Sélectionner fichier(s) ou photo(s)",
            filetypes=[("Tous", "*.*"),
                       ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.heic"),
                       ("PDF",    "*.pdf"),
                       ("Documents", "*.doc *.docx *.xls *.xlsx *.txt")])
        if not files: return
        d.mkdir(parents=True, exist_ok=True)
        n = 0
        for f in files:
            try:
                shutil.copy2(f, d / Path(f).name); n += 1
            except (OSError, shutil.SameFileError) as e:
                messagebox.showwarning("Copie", f"Erreur sur {Path(f).name} :\n{e}")
        if n:
            messagebox.showinfo("Import", f"{n} fichier(s) ajouté(s) dans {d.name}.")
        self._refresh_files()

    def _ouvrir_dossier(self):
        d = self._current_dossier()
        if not d:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon pour créer le dossier.")
            return
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    # ─── Notifications ────────────────────────────────────────────────────────
    def _mail_client_inline(self):
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon avant de notifier le client.")
            return
        inv = db.get_intervention(inv_id=self.inv_id)
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        if not client or not row_get(client, "email"):
            messagebox.showwarning("Email manquant",
                "Aucun email renseigné pour ce client.")
            return
        path = sauvegarder_bon(inv)
        mailer.email_client(inv, client, moteur, str(path))
        db.mark_notifie(self.inv_id, "client")

    def _mail_tech_inline(self):
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon avant de notifier le technicien.")
            return
        inv = db.get_intervention(inv_id=self.inv_id)
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        tech = db.get_technicien_by_nom(row_get(inv, "technicien"))
        email = row_get(tech, "email")
        path = sauvegarder_bon(inv)
        mailer.email_technicien(inv, client, moteur, email, str(path))
        db.mark_notifie(self.inv_id, "tech")

    # ─── Chargement ───────────────────────────────────────────────────────────
    def _load(self):
        inv = db.get_intervention(inv_id=self.inv_id)
        if not inv: return

        c = self._client_by_id.get(row_get(inv, "client_id"))
        if c:
            self.client_var.set(c["nom"])
        self.lieu_var.set(row_get(inv, "lieu_intervention"))
        self.signataire_var.set(row_get(inv, "nom_signataire"))
        self.email_sig_var.set(row_get(inv, "email_signataire"))

        if c:
            filtrés = [m for m in self._all_moteurs if m["client_id"] == c["id"]]
        else:
            filtrés = self._all_moteurs
        self.moteur_combo["values"] = [m["num_serie"] for m in filtrés]
        m = self._moteur_by_id.get(row_get(inv, "moteur_id"))
        if m:
            self.moteur_var.set(m["num_serie"])
            self._set_moteur_info(m)
        self.nb_heures_var.set(row_get(inv, "nb_heures_fct"))

        self.type_var.set(row_get(inv, "type_intervention"))
        self.tech_var.set(row_get(inv, "technicien"))
        self.date_var.set(row_get(inv, "date_creation"))
        self.statut_var.set(row_get(inv, "statut", "En cours"))
        self.urgence_var.set(row_get(inv, "urgence", "Normale") or "Normale")

        # Cases à cocher (toutes : options + classifications)
        for k in self.chk:
            self.chk[k].set(int(row_get(inv, k, 0)) or 0)

        self.txt_demande.delete("1.0", "end")
        self.txt_demande.insert("1.0", row_get(inv, "demande_client") or row_get(inv, "description"))
        self.txt_constat.delete("1.0", "end")
        self.txt_constat.insert("1.0", row_get(inv, "constat"))
        self.txt_travaux.delete("1.0", "end")
        self.txt_travaux.insert("1.0", row_get(inv, "travaux"))
        self.txt_info.delete("1.0", "end")
        self.txt_info.insert("1.0", row_get(inv, "informations"))

        try:
            mats = json.loads(row_get(inv, "materiels_json", "[]") or "[]")
        except (json.JSONDecodeError, TypeError):
            mats = []
        legacy = row_get(inv, "pieces")
        if not mats and legacy:
            mats = [{"qte": "", "ref": "", "designation": legacy}]
        self.materiels_tbl.load(mats)

        try:
            depl = json.loads(row_get(inv, "deplacements_json", "{}") or "{}")
        except (json.JSONDecodeError, TypeError):
            depl = {}
        self.depl_tbl.load(depl)

        self._refresh_files()

    # ─── Sauvegarde ───────────────────────────────────────────────────────────
    def _save(self, generer=False):
        cn   = self.client_var.get().strip()
        ns   = self.moteur_var.get().strip()
        ti   = self.type_var.get().strip()
        tech = self.tech_var.get().strip()
        date_str = self.date_var.get().strip()
        demande  = self.txt_demande.get("1.0", "end").strip()

        if not all([cn, ns, ti, tech, demande]):
            messagebox.showwarning("Champs manquants",
                "Client, N° Série, Type, Technicien et Demande du client sont obligatoires.")
            return

        if date_str and not DateEntry.is_valid(date_str):
            messagebox.showwarning("Date invalide",
                f"La date '{date_str}' est invalide.\nFormat attendu : JJ/MM/AAAA")
            return

        # Email signataire : warning informatif si format douteux
        email_sig = self.email_sig_var.get().strip()
        if email_sig and not db.email_looks_valid(email_sig):
            if not messagebox.askyesno("Format email douteux",
                    f"L'email signataire '{email_sig}' semble mal formé.\n\n"
                    "Voulez-vous quand même enregistrer ?"):
                return

        client = self._client_by_nom.get(cn)
        moteur = self._moteur_by_ns.get(ns)

        if tech and not db.get_technicien_by_nom(tech):
            if messagebox.askyesno("Nouveau technicien",
                    f"'{tech}' n'est pas dans l'annuaire.\nL'ajouter maintenant (sans email) ?"):
                db.upsert_technicien({"nom": tech, "email": "", "telephone": ""})

        materiels    = self.materiels_tbl.to_list()
        deplacements = self.depl_tbl.to_dict()

        data = {
            "client_id":         client["id"] if client else "",
            "moteur_id":         moteur["id"] if moteur else "",
            "type_intervention": ti,
            "urgence":           self.urgence_var.get(),
            "technicien":        tech,
            "date_creation":     date_str,
            "statut":            self.statut_var.get(),

            "lieu_intervention": self.lieu_var.get().strip(),
            "nom_signataire":    self.signataire_var.get().strip(),
            "email_signataire":  email_sig,
            "nb_heures_fct":     self.nb_heures_var.get().strip(),

            **{k: int(v.get()) for k, v in self.chk.items()},

            "demande_client":    demande,
            "constat":           self.txt_constat.get("1.0", "end").strip(),
            "travaux":           self.txt_travaux.get("1.0", "end").strip(),
            "informations":      self.txt_info.get("1.0", "end").strip(),

            "description":       demande,
            "pieces":            "",

            "materiels_json":    json.dumps(materiels, ensure_ascii=False),
            "deplacements_json": json.dumps(deplacements, ensure_ascii=False),
        }

        if self.is_edit:
            db.update_intervention(self.inv_id, data)
            num_bon = db.get_intervention(inv_id=self.inv_id)["num_bon"]
        else:
            iid, num_bon = db.create_intervention(data)
            self.inv_id = iid
            self.is_edit = True

        if self.on_save:
            self.on_save()

        if generer:
            inv = db.get_intervention(inv_id=self.inv_id)
            path = sauvegarder_bon(inv)
            messagebox.showinfo("Bon généré",
                f"✅ {num_bon}\nEnregistré dans :\n{path}\n\nOuverture dans le navigateur.")
            ouvrir_fichier(path)
            self.destroy()
        else:
            self._refresh_files()
            messagebox.showinfo("Enregistré",
                f"✅ Bon {num_bon} enregistré.")


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG IMPORT CSV (parc moteurs & clients)
# ══════════════════════════════════════════════════════════════════════════════
class ImportCSVDialog(tk.Toplevel):
    """
    Dialogue d'import CSV avec :
    - choix du fichier
    - aperçu des premières lignes
    - détection automatique des colonnes (modifiable)
    - dry-run pour valider avant import
    - options de dédoublonnage
    """

    # Libellés français pour les champs DB
    FIELD_LABELS = [
        ("nom_client",       "Nom client (Tiers) *"),
        ("type_client",      "Type Client"),
        ("navire",           "Machine / Engin (Navire)"),
        ("num_serie",        "N° de Série *"),
        ("num_moteur",       "N° Moteur"),
        ("cylindree",        "Cylindrée"),
        ("famille",          "Famille"),
        ("ref_constructeur", "Réf. Constructeur"),
        ("application",      "Application"),
        ("type_moteur",      "Type moteur"),
        ("typologie",        "Typologie"),
        ("marque",           "Marque"),
        ("collection",       "Collection"),
        ("code_affaire",     "Code Affaire"),
    ]

    def __init__(self, parent, app, on_save=None):
        super().__init__(parent)
        self.app = app
        self.on_save = on_save
        self.title("Importer le parc depuis un CSV")
        self.geometry("960x720")
        self.configure(bg=C["bg"])
        self.grab_set()

        self.csv_path     = None
        self.headers      = []
        self.preview_rows = []
        self.total_lines  = 0
        self.mapping      = {}
        self.delimiter    = ";"
        self.encoding     = "utf-8"
        self.column_vars  = {}  # field → StringVar (header sélectionné)

        # Boutons EN BAS d'abord (sinon poussés hors fenêtre)
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(side="bottom", pady=10)
        self.btn_simulate = mk_btn(bf, "🔍 Simuler (dry-run)", self._simulate)
        self.btn_simulate.pack(side="left", padx=6)
        self.btn_import = mk_btn(bf, "📥 Lancer l'import", self._import_real, color=C["btn2"])
        self.btn_import.pack(side="left", padx=6)
        mk_btn(bf, "Fermer", self.destroy, color="#888").pack(side="left", padx=6)

        # Désactivés tant qu'aucun fichier n'est chargé
        self.btn_simulate.config(state="disabled")
        self.btn_import.config(state="disabled")

        # Bandeau en-tête
        tk.Label(self, text="📥 Import du parc moteurs & clients",
                 font=("Arial", 13, "bold"), bg=C["bg"], fg=C["header"]).pack(pady=(12, 2))
        tk.Label(self,
                 text="1. Sélectionnez votre fichier CSV   →   2. Vérifiez le mapping des colonnes   "
                      "→   3. Simulez puis importez",
                 bg=C["bg"], font=("Arial", 9), fg="#666").pack(pady=(0, 6))

        # Choix du fichier
        ff = tk.Frame(self, bg=C["bg"])
        ff.pack(fill="x", padx=14, pady=4)
        mk_btn(ff, "📁 Choisir un fichier CSV…", self._choose_file).pack(side="left", padx=(0, 8))
        self.file_lbl = tk.Label(ff, text="(aucun fichier sélectionné)",
                                  bg=C["bg"], font=("Arial", 9), fg="#666", anchor="w")
        self.file_lbl.pack(side="left", fill="x", expand=True)

        # Notebook 2 onglets : Mapping + Aperçu
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=14, pady=8)

        self.mapping_tab = tk.Frame(nb, bg=C["bg"])
        nb.add(self.mapping_tab, text="  Mapping des colonnes  ")
        self.preview_tab = tk.Frame(nb, bg=C["bg"])
        nb.add(self.preview_tab, text="  Aperçu du fichier  ")

        # Options en bas du mapping
        opt_f = tk.Frame(self, bg=C["bg"])
        opt_f.pack(fill="x", padx=14, pady=(0, 4))
        self.skip_empty_var = tk.IntVar(value=0)
        tk.Checkbutton(opt_f,
            text="Ignorer les lignes sans N° de série (recommandé pour les réimports)",
            variable=self.skip_empty_var, bg=C["bg"], font=("Arial", 9)).pack(anchor="w")

        # Placeholders initiaux
        tk.Label(self.mapping_tab,
                 text="Chargez d'abord un fichier CSV pour configurer le mapping.",
                 bg=C["bg"], font=("Arial", 10, "italic"), fg="#888").pack(pady=40)
        tk.Label(self.preview_tab,
                 text="Aperçu disponible après chargement d'un fichier.",
                 bg=C["bg"], font=("Arial", 10, "italic"), fg="#888").pack(pady=40)

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV du parc",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous", "*.*")])
        if not path:
            return
        try:
            hd, prev, total, mapping, delim, enc = csv_importer.read_csv_preview(path)
        except Exception as e:
            messagebox.showerror("Lecture impossible",
                                  f"Impossible de lire le fichier :\n{e}")
            return

        if not hd:
            messagebox.showwarning("Fichier vide", "Le fichier ne contient pas d'en-têtes.")
            return

        self.csv_path     = path
        self.headers      = hd
        self.preview_rows = prev
        self.total_lines  = total
        self.mapping      = mapping
        self.delimiter    = delim
        self.encoding     = enc

        self.file_lbl.config(
            text=f"📄 {Path(path).name} – {total} ligne(s), {len(hd)} colonne(s) – "
                  f"séparateur '{delim}', encodage {enc}")
        self._render_mapping_tab()
        self._render_preview_tab()
        self.btn_simulate.config(state="normal")
        self.btn_import.config(state="normal")

    def _render_mapping_tab(self):
        for w in self.mapping_tab.winfo_children():
            w.destroy()

        tk.Label(self.mapping_tab,
                 text="Pour chaque champ EMS, sélectionnez la colonne CSV correspondante "
                      "(— laisser vide pour ignorer) :",
                 bg=C["bg"], font=("Arial", 9), fg="#444",
                 anchor="w", justify="left").pack(anchor="w", padx=10, pady=(8, 4))

        # Scrollable area
        canvas = tk.Canvas(self.mapping_tab, bg=C["bg"], highlightthickness=0, height=260)
        vsb = ttk.Scrollbar(self.mapping_tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True, padx=8)
        inner = tk.Frame(canvas, bg=C["bg"])
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))

        choices = ["(ignorer)"] + list(self.headers)
        self.column_vars = {}
        for field, label in self.FIELD_LABELS:
            row = tk.Frame(inner, bg=C["bg"])
            row.pack(fill="x", padx=4, pady=2)
            tk.Label(row, text=label, bg=C["bg"], font=("Arial", 10),
                     anchor="e", width=28).pack(side="left", padx=(0, 8))
            idx = self.mapping.get(field)
            current = self.headers[idx] if idx is not None else "(ignorer)"
            var = tk.StringVar(value=current)
            self.column_vars[field] = var
            ttk.Combobox(row, textvariable=var, values=choices,
                         width=42, state="readonly").pack(side="left", fill="x", expand=True)

    def _render_preview_tab(self):
        for w in self.preview_tab.winfo_children():
            w.destroy()

        if not self.preview_rows:
            tk.Label(self.preview_tab, text="(fichier vide)",
                     bg=C["bg"], font=("Arial", 10), fg="#888").pack(pady=20)
            return

        info = tk.Label(self.preview_tab,
                         text=f"Aperçu des {len(self.preview_rows)} premières lignes "
                              f"(sur {self.total_lines} au total) :",
                         bg=C["bg"], font=("Arial", 9, "italic"), fg="#666")
        info.pack(anchor="w", padx=8, pady=(8, 4))

        # Treeview pour les 15 premières lignes
        cols = [f"c{i}" for i in range(len(self.headers))]
        col_defs = [(f"c{i}", h, max(80, min(180, len(h) * 10)), "w")
                    for i, h in enumerate(self.headers)]
        tf, tree = mk_tree(self.preview_tab, cols, col_defs, height=12)
        tf.pack(fill="both", expand=True, padx=8, pady=4)
        fill_tree(tree, [tuple(r + [""] * (len(self.headers) - len(r)))
                          for r in self.preview_rows])

    def _build_mapping_from_ui(self):
        """Reconstruit le dict mapping à partir des combos utilisateur."""
        mapping = {}
        for field, var in self.column_vars.items():
            h = var.get()
            if h == "(ignorer)" or h not in self.headers:
                mapping[field] = None
            else:
                mapping[field] = self.headers.index(h)
        return mapping

    def _simulate(self):
        if not self.csv_path: return
        mapping = self._build_mapping_from_ui()
        if mapping.get("num_serie") is None and mapping.get("nom_client") is None:
            messagebox.showwarning("Mapping incomplet",
                "Vous devez mapper au moins le N° de série OU le nom client.")
            return
        try:
            stats = csv_importer.import_rows(
                self.csv_path, mapping, self.delimiter, self.encoding,
                skip_empty_serie=bool(self.skip_empty_var.get()), dry_run=True)
        except Exception as e:
            messagebox.showerror("Erreur de simulation", str(e))
            return
        self._show_stats("Résultat de la simulation (aucune donnée écrite)", stats, color="#084298")

    def _import_real(self):
        if not self.csv_path: return
        mapping = self._build_mapping_from_ui()
        if mapping.get("num_serie") is None and mapping.get("nom_client") is None:
            messagebox.showwarning("Mapping incomplet",
                "Vous devez mapper au moins le N° de série OU le nom client.")
            return
        if not messagebox.askyesno("Confirmer l'import",
                f"Importer {self.total_lines} ligne(s) dans la base ?\n\n"
                "Les moteurs avec un N° de série existant seront MIS À JOUR.\n"
                "Les clients avec un nom déjà connu seront RÉUTILISÉS."):
            return
        try:
            stats = csv_importer.import_rows(
                self.csv_path, mapping, self.delimiter, self.encoding,
                skip_empty_serie=bool(self.skip_empty_var.get()), dry_run=False)
        except Exception as e:
            messagebox.showerror("Erreur d'import", str(e))
            return
        self._show_stats("Import terminé", stats, color="#0f5132")
        if self.on_save:
            self.on_save()

    def _show_stats(self, title, stats, color="#003366"):
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=C["bg"])
        win.geometry("520x420")
        win.grab_set()
        tk.Label(win, text=title, font=("Arial", 13, "bold"),
                 bg=C["bg"], fg=color).pack(pady=(14, 8))

        f = tk.Frame(win, bg=C["bg"])
        f.pack(padx=20, pady=4, fill="both", expand=True)
        rows = [
            ("Lignes traitées",     stats["lignes_total"]),
            ("Lignes ignorées",     stats["lignes_ignorees"]),
            ("Clients créés",       stats["clients_crees"]),
            ("Clients existants",   stats["clients_existants"]),
            ("Moteurs créés",       stats["moteurs_crees"]),
            ("Moteurs mis à jour",  stats["moteurs_mis_a_jour"]),
            ("Doublons N° série",   stats["moteurs_doublons"]),
            ("Erreurs",             len(stats["erreurs"])),
        ]
        for label, val in rows:
            r = tk.Frame(f, bg=C["bg"])
            r.pack(fill="x", pady=2)
            tk.Label(r, text=label, bg=C["bg"], font=("Arial", 10),
                     anchor="w", width=24).pack(side="left")
            color_v = C["danger"] if (label == "Erreurs" and val) else "#000"
            tk.Label(r, text=str(val), bg=C["bg"], font=("Arial", 11, "bold"),
                     fg=color_v, anchor="w").pack(side="left")

        if stats["erreurs"]:
            tk.Label(f, text="Détail des erreurs :", bg=C["bg"],
                     font=("Arial", 9, "bold"), fg=C["danger"]).pack(anchor="w", pady=(10, 2))
            tx = tk.Text(f, height=6, font=("Arial", 9), wrap="word",
                          relief="solid", bd=1)
            tx.pack(fill="both", expand=True)
            for ln, msg in stats["erreurs"][:20]:
                tx.insert("end", f"  Ligne {ln}: {msg}\n")
            if len(stats["erreurs"]) > 20:
                tx.insert("end", f"  ... et {len(stats['erreurs']) - 20} autres\n")
            tx.config(state="disabled")

        mk_btn(win, "Fermer", win.destroy).pack(pady=10)


# ─── Lancement ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AppEMS()
    app.mainloop()
