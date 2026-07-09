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

from ems_client import api as db
from shared import mailer
from shared import csv_importer
from shared.bon_generator import sauvegarder_bon, ouvrir_fichier, apply_icon


def _check_api_startup():
    """Vérifie la connectivité sans bloquer l'ouverture de l'app."""
    pass  # Le statut est affiché dans la bannière de la sidebar


_check_api_startup()


def _dossiers_root() -> Path:
    """
    Retourne le dossier racine ou sont stockes les sous-dossiers des bons.
    
    Priorite :
      1. Cle 'dossiers_root' dans config.ini (section [files])
      2. Sinon en mode .exe : a cote de l'executable
      3. Sinon en mode dev : a cote de main.py
    
    Le dossier est cree s'il n'existe pas.
    """
    import sys
    from configparser import ConfigParser
    
    # Determiner le dossier de base (exe ou script)
    if getattr(sys, "frozen", False):
        # .exe PyInstaller : dossier de l'executable, pas du _MEI temporaire
        base_default = Path(sys.executable).parent
    else:
        base_default = Path(__file__).resolve().parent
    
    # Chercher config.ini a cote de l'exe/script
    candidats_cfg = [
        base_default / "config.ini",
        Path(__file__).resolve().parent.parent / "config.ini",
    ]
    for cfg_path in candidats_cfg:
        if cfg_path.is_file():
            try:
                cp = ConfigParser()
                cp.read(cfg_path, encoding="utf-8")
                custom = cp.get("files", "dossiers_root", fallback="").strip()
                if custom:
                    p = Path(custom)
                    p.mkdir(parents=True, exist_ok=True)
                    return p
            except Exception:
                pass
            break
    
    # Defaut : sous-dossier 'dossiers' a cote de l'exe/script
    p = base_default / "dossiers"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _dossiers_root_garanties() -> Path:
    """Dossier racine des fiches de garantie (même logique que garantie_generator)."""
    try:
        from garantie_generator import _get_garanties_root
        return _get_garanties_root()
    except Exception:
        p = Path(__file__).resolve().parent / "garanties"
        p.mkdir(parents=True, exist_ok=True)
        return p


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
    "se_row_even": "#dce6f6",  # sous-ensembles — fond pair (bleu EMS atténué)
    "se_row_odd":  "#e8f0fb",  # sous-ensembles — fond impair
    "border":     "#d8dee5",   # bordures douces
    # ─── États / urgence ──────────────────────────────────────────────────────
    "urg_normale":  "#6b7785",
    "urg_urgente":  "#e67e22",
    "urg_critique": "#c62828",
    "ec":           "#f59e0b",  # En cours - ambre
    "ec_bg":        "#fff7ed",
    "afact":        "#6366f1",  # À facturer - indigo (intermédiaire)
    "afact_bg":     "#eef2ff",
    "fact":         "#3b82f6",  # Facturé - bleu vif
    "fact_bg":      "#eff6ff",
    "clos":         "#10b981",  # Clos - vert
    "clos_bg":      "#ecfdf5",
    "prog":         "#0ea5e9",  # Date à programmer - bleu ciel
    "prog_bg":      "#e0f2fe",
}
STATUTS  = ["En cours", "Date à programmer", "À facturer", "Facturé", "Clos"]
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


def mk_tree(parent, cols, col_defs, height=18, show_tree=False, show_hsb=True):
    """col_defs : liste de tuples (cid, lbl, width, [anchor]).
    show_tree=True : conserve la colonne arborescence (#0) avec les flèches
    ▶/▼ d'expansion, pour les vues avec hiérarchie parent/enfant.
    """
    # Container avec bordure douce
    wrapper = tk.Frame(parent, bg=C["border"], bd=0)
    frame = tk.Frame(wrapper, bg=C["surface"])
    frame.pack(fill="both", expand=True, padx=1, pady=1)

    vsb = ttk.Scrollbar(frame, orient="vertical")
    tree = ttk.Treeview(frame, columns=cols,
                        show=("tree headings" if show_tree else "headings"),
                        height=height, yscrollcommand=vsb.set,
                        style="EMS.Treeview")
    if show_hsb:
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        hsb.config(command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
    if show_tree:
        tree.column("#0", width=26, minwidth=26, stretch=False, anchor="center")
        tree.heading("#0", text="")
    vsb.config(command=tree.yview)
    vsb.pack(side="right", fill="y")
    tree.pack(fill="both", expand=True)

    # Tags pour urgence (fond + couleur de texte)
    tree.tag_configure("even", background=C["row_even"])
    tree.tag_configure("odd",  background=C["row_odd"])
    tree.tag_configure("urg_critique", background="#fef2f2", foreground=C["urg_critique"])
    tree.tag_configure("urg_urgente",  background="#fff7ed", foreground=C["urg_urgente"])
    if show_tree:
        tree.tag_configure("se_even", background=C["se_row_even"])
        tree.tag_configure("se_odd",  background=C["se_row_odd"])

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
    """
    Entry avec masque JJ/MM/AAAA.
    Toutes les touches sont interceptées en <KeyPress> afin de gérer la
    position du curseur manuellement — évite le saut de curseur causé par
    var.set() dans un callback trace.
    """

    def __init__(self, master, textvariable=None, width=46, **kw):
        self.var = textvariable or tk.StringVar()
        super().__init__(master, textvariable=self.var, width=width, **kw)
        self._lock = False
        self.var.trace_add("write", self._on_ext_write)   # paste / set externe
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<KeyPress>", self._on_key)

    # ── helpers ───────────────────────────────────────────────────────────────
    @staticmethod
    def _fmt(digits: str) -> str:
        d = digits[:8]
        if len(d) <= 2:   return d
        if len(d) <= 4:   return d[:2] + "/" + d[2:]
        return d[:2] + "/" + d[2:4] + "/" + d[4:]

    @staticmethod
    def _digit_pos(s: str, n: int) -> int:
        """Index juste après le n-ième chiffre de s (pour replacer le curseur)."""
        count = 0
        for i, c in enumerate(s):
            if c.isdigit():
                count += 1
                if count == n:
                    return i + 1
        return len(s)

    def _apply(self, new_s: str, new_pos: int) -> None:
        self._lock = True
        self.var.set(new_s)
        self._lock = False
        self.icursor(new_pos)

    # ── trace : paste ou set() externe (chargement d'un bon) ─────────────────
    def _on_ext_write(self, *_):
        if self._lock:
            return
        s = self.var.get()
        out = self._fmt("".join(c for c in s if c.isdigit()))
        if out == s:
            return
        try:
            nb = sum(1 for c in s[:self.index(tk.INSERT)] if c.isdigit())
        except Exception:
            nb = len("".join(c for c in out if c.isdigit()))
        self._lock = True
        self.var.set(out)
        self._lock = False
        self.after(0, lambda p=self._digit_pos(out, nb): self.icursor(p))

    # ── interception clavier ──────────────────────────────────────────────────
    def _on_key(self, ev):
        key   = ev.keysym
        char  = ev.char

        # Laisser passer : navigation, Tab, Entrée, Ctrl+* (copier/coller/…)
        if key in ("Left", "Right", "Home", "End", "Tab", "Return",
                   "KP_Enter", "Escape"):
            return
        if ev.state & 0x4:     # Ctrl maintenu
            return

        s   = self.var.get()
        pos = self.index(tk.INSERT)

        # Sélection active ?
        sel0 = sel1 = None
        try:
            sel0 = self.index(tk.SEL_FIRST)
            sel1 = self.index(tk.SEL_LAST)
        except tk.TclError:
            pass

        # ── Backspace ─────────────────────────────────────────────────────────
        if key == "BackSpace":
            if sel0 is not None:
                raw = "".join(c for c in s[:sel0] + s[sel1:] if c.isdigit())
                nb  = sum(1 for c in s[:sel0] if c.isdigit())
            elif pos > 0:
                dp = pos - 1
                if dp < len(s) and s[dp] == "/":
                    dp -= 1
                if dp < 0:
                    return "break"
                raw = "".join(c for c in s[:dp] + s[dp+1:] if c.isdigit())
                nb  = sum(1 for c in s[:dp] if c.isdigit())
            else:
                return "break"
            out = self._fmt(raw)
            self._apply(out, self._digit_pos(out, nb))
            return "break"

        # ── Delete ────────────────────────────────────────────────────────────
        if key == "Delete":
            if sel0 is not None:
                raw = "".join(c for c in s[:sel0] + s[sel1:] if c.isdigit())
                nb  = sum(1 for c in s[:sel0] if c.isdigit())
            elif pos < len(s):
                dp = pos
                if s[dp] == "/":
                    dp += 1
                if dp >= len(s):
                    return "break"
                raw = "".join(c for c in s[:dp] + s[dp+1:] if c.isdigit())
                nb  = sum(1 for c in s[:pos] if c.isdigit())
            else:
                return "break"
            out = self._fmt(raw)
            self._apply(out, self._digit_pos(out, nb))
            return "break"

        # ── Chiffres uniquement ───────────────────────────────────────────────
        if not char or not char.isdigit():
            return "break"

        # Déjà 8 chiffres et pas de sélection → refuser
        if sel0 is None and sum(1 for c in s if c.isdigit()) >= 8:
            return "break"

        # Sauter par-dessus un / si le curseur est juste devant
        if sel0 is None and pos < len(s) and s[pos] == "/":
            pos += 1

        if sel0 is not None:
            before, after = s[:sel0], s[sel1:]
        else:
            before = s[:pos]
            after  = s[pos+1:] if pos < len(s) else ""

        raw = "".join(c for c in before + char + after if c.isdigit())[:8]
        out = self._fmt(raw)
        nb  = sum(1 for c in before if c.isdigit()) + 1
        self._apply(out, self._digit_pos(out, nb))
        return "break"

    # ── validation visuelle à la perte du focus ───────────────────────────────
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
class SearchableCombobox(tk.Frame):
    """
    Combobox avec recherche textuelle, robuste au scroll.

    Implémentation : Entry + bouton flèche + Listbox affichée EN FLUX
    (dans le même conteneur, pas une Toplevel flottante). La liste se
    place donc naturellement dans le formulaire et suit le scroll sans
    artefact ni calcul de position.

    - Saisie normale (curseur, sélection natifs)
    - Filtrage instantané, tri par pertinence (préfixe puis contenu)
    - ↓/↑ pour naviguer, Entrée pour valider, Échap pour fermer
    - Clic sur une entrée pour la sélectionner
    - Émet l'événement virtuel <<ComboboxSelected>> à la sélection

    API compatible drop-in avec ttk.Combobox :
      .get() / .set(value) / .set_values(values)
      .bind("<<ComboboxSelected>>", handler)
      .configure(foreground=...)
    """
    def __init__(self, master, textvariable=None, values=None, width=44,
                 case_sensitive=False, max_visible=8, allow_free_text=False,
                 min_chars_to_open=0, **kw):
        for k in ("state", "values", "textvariable", "validate", "validatecommand"):
            kw.pop(k, None)
        super().__init__(master, bg=master.cget("bg") if "bg" not in kw else kw.get("bg"))

        self._all_values      = list(values or [])
        self._filtered        = list(self._all_values)
        self._case_sensitive  = case_sensitive
        self._max_visible     = max_visible
        self._allow_free_text = allow_free_text
        self._min_chars_to_open = min_chars_to_open
        self.var             = textvariable or tk.StringVar()
        self._suppress_trace = False
        self._is_open        = False

        # ── Ligne 1 : Entry + flèche ─────────────────────────────────────────
        top = tk.Frame(self, bg=self.cget("bg"))
        top.pack(fill="x")
        self.entry = ttk.Entry(top, textvariable=self.var, width=width)
        self.entry.pack(side="left", fill="x", expand=True)
        self.arrow = tk.Label(top, text="▾", bg="#e8eaed", fg="#444",
                               font=("Segoe UI", 9, "bold"),
                               cursor="hand2", padx=6, pady=1, bd=1, relief="solid")
        self.arrow.pack(side="left", fill="y")
        self.arrow.bind("<Button-1>", lambda _e: self.toggle_dropdown())

        # ── Ligne 2 : zone liste (masquée par défaut) ────────────────────────
        # Conteneur réservé : on le pack/forget pour afficher/cacher la liste.
        self._list_holder = tk.Frame(self, bg="#9ba5b1")
        # Pas packé maintenant ; on le packe à l'ouverture.
        self._list_inner = tk.Frame(self._list_holder, bg="white")
        self._list_inner.pack(fill="both", expand=True, padx=1, pady=1)
        self._listbox = tk.Listbox(self._list_inner, font=("Segoe UI", 10),
                                    bd=0, highlightthickness=0,
                                    selectbackground=C["nav_sel"],
                                    selectforeground="white",
                                    activestyle="none",
                                    exportselection=False,
                                    height=1)
        self._listbox.pack(side="left", fill="both", expand=True)
        self._scrollbar = ttk.Scrollbar(self._list_inner, orient="vertical",
                                         command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=self._scrollbar.set)
        # scrollbar packée seulement si nécessaire

        # ── Bindings ─────────────────────────────────────────────────────────
        self.var.trace_add("write", self._on_text_change)
        self.entry.bind("<Down>",     self._on_arrow_down)
        self.entry.bind("<Up>",       self._on_arrow_up)
        self.entry.bind("<Return>",   self._on_enter)
        self.entry.bind("<Escape>",   self._on_escape)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self._listbox.bind("<ButtonRelease-1>", self._on_listbox_click)
        self._listbox.bind("<Return>",          self._on_listbox_enter)
        self._listbox.bind("<Double-Button-1>", self._on_listbox_enter)
        self._listbox.bind("<Escape>",
                            lambda _e: (self._close_dropdown(), self.entry.focus_set()))
        # Capture la molette dans la listbox et empeche la propagation a la dialog
        def _wheel_listbox(ev):
            self._listbox.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            return "break"
        self._listbox.bind("<MouseWheel>", _wheel_listbox)
        self._list_holder.bind("<MouseWheel>", _wheel_listbox)
        self._list_inner.bind("<MouseWheel>", _wheel_listbox)
    # ─── API publique ─────────────────────────────────────────────────────────
    def get(self):
        return self.var.get()

    def set(self, value):
        self._suppress_trace = True
        self.var.set(value)
        self._suppress_trace = False
        self._filtered = list(self._all_values)
        self._close_dropdown()

    def set_values(self, values):
        self._all_values = list(values or [])
        self._filtered   = list(self._all_values)
        if self._is_open:
            self._refresh_listbox()

    def get_all_values(self):
        return list(self._all_values)

    def configure(self, **kw):
        fg = kw.pop("foreground", kw.pop("fg", None))
        if fg is not None:
            try:
                self.entry.configure(foreground=fg)
            except tk.TclError:
                pass
        if kw:
            super().configure(**kw)
    config = configure

    # ─── Filtrage ─────────────────────────────────────────────────────────────
    @staticmethod
    def _norm(s):
        if s is None:
            return ""
        s = str(s).lower()
        trans = str.maketrans("àâäéèêëîïôöùûüç", "aaaeeeeiioouuuc")
        return s.translate(trans)

    def _filter(self, query):
        if not query:
            return list(self._all_values)
        if self._case_sensitive:
            q = query
            starts = [v for v in self._all_values if v.startswith(q)]
            contains = [v for v in self._all_values
                        if q in v and not v.startswith(q)]
            return starts + contains
        q = self._norm(query)
        starts, contains = [], []
        for v in self._all_values:
            nv = self._norm(v)
            if nv.startswith(q):
                starts.append(v)
            elif q in nv:
                contains.append(v)
        return starts + contains

    # ─── Événements de saisie ─────────────────────────────────────────────────
    def _on_text_change(self, *_):
        if self._suppress_trace:
            return
        query = self.var.get()
        self._filtered = self._filter(query)
        enough = len(query) >= max(1, self._min_chars_to_open)
        if enough and (self._filtered or not self._allow_free_text):
            self._open_dropdown()
            self._refresh_listbox()
        else:
            self._close_dropdown()

    def _on_arrow_down(self, _ev=None):
        if not self._is_open:
            self._filtered = self._filter(self.var.get())
            self._open_dropdown()
            self._refresh_listbox()
        if self._listbox.size() > 0:
            self._listbox.focus_set()
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(0)
            self._listbox.activate(0)
            self._listbox.see(0)
        return "break"

    def _on_arrow_up(self, _ev=None):
        if self._is_open and self._listbox.size() > 0:
            self._listbox.focus_set()
            last = self._listbox.size() - 1
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(last)
            self._listbox.activate(last)
            self._listbox.see(last)
        return "break"

    def _on_enter(self, _ev=None):
        if self._is_open and self._filtered:
            # Sélectionne la 1ère suggestion
            self._select_value(self._filtered[0])
        return "break"

    def _on_escape(self, _ev=None):
        self._close_dropdown()
        return "break"

    def _on_focus_out(self, _ev=None):
        # Fermeture différée : laisse le temps à un clic listbox d'aboutir
        self.after(120, self._maybe_close_on_focus_out)

    def _maybe_close_on_focus_out(self):
        try:
            focused = self.focus_get()
        except (tk.TclError, KeyError):
            focused = None
        # Si le focus est dans la listbox, on ne ferme pas
        w = focused
        try:
            while w is not None:
                if w is self._listbox or w is self.entry:
                    return
                w = w.master
        except (tk.TclError, AttributeError):
            pass
        self._close_dropdown()
        # Coloration informative si valeur hors liste (uniquement pour les champs à sélection obligatoire)
        v = self.var.get().strip()
        try:
            if v and v not in self._all_values and not self._allow_free_text:
                self.entry.configure(foreground=C["warn"])
            else:
                self.entry.configure(foreground="black")
        except tk.TclError:
            pass

    # ─── Sélection dans la liste ──────────────────────────────────────────────
    def _on_listbox_click(self, ev):
        try:
            idx = self._listbox.nearest(ev.y)
            if 0 <= idx < len(self._filtered):
                self._select_value(self._filtered[idx])
        except (tk.TclError, IndexError):
            pass
        return "break"

    def _on_listbox_enter(self, _ev=None):
        sel = self._listbox.curselection()
        if sel and sel[0] < len(self._filtered):
            self._select_value(self._filtered[sel[0]])
        return "break"

    def _select_value(self, value):
        self._suppress_trace = True
        self.var.set(value)
        self._suppress_trace = False
        self._close_dropdown()
        try:
            self.entry.icursor("end")
            self.entry.configure(foreground="black")
            self.entry.focus_set()
        except tk.TclError:
            pass
        try:
            self.event_generate("<<ComboboxSelected>>")
        except tk.TclError:
            pass

    # ─── Ouverture / fermeture (in-flow, pas de Toplevel) ─────────────────────
    def toggle_dropdown(self):
        if self._is_open:
            self._close_dropdown()
        else:
            q = self.var.get()
            if self._min_chars_to_open and len(q) < self._min_chars_to_open:
                self.entry.focus_set()
                return
            self._filtered = self._filter(q)
            self._open_dropdown()
            self._refresh_listbox()
            self.entry.focus_set()

    def _open_dropdown(self):
        if self._is_open:
            return
        self._list_holder.pack(fill="x", pady=(2, 0))
        self._is_open = True

    def _close_dropdown(self):
        if not self._is_open:
            return
        try:
            self._list_holder.pack_forget()
        except tk.TclError:
            pass
        self._is_open = False

    def _refresh_listbox(self):
        if not self._is_open:
            return
        try:
            self._listbox.delete(0, "end")
            if not self._filtered:
                self._listbox.insert("end", "  (aucune correspondance)")
                self._listbox.itemconfigure(0, foreground="#9ba5b1")
                self._listbox.configure(height=1)
                self._scrollbar.pack_forget()
            else:
                for v in self._filtered:
                    self._listbox.insert("end", "  " + v)
                n = len(self._filtered)
                visible = min(self._max_visible, n)
                self._listbox.configure(height=visible)
                # Scrollbar seulement si la liste dépasse la zone visible
                if n > visible:
                    self._scrollbar.pack(side="right", fill="y")
                else:
                    self._scrollbar.pack_forget()
        except tk.TclError:
            pass


# ─── Sélecteur multi-techniciens (chips ajoutables) ───────────────────────────
class TechniciensPicker(tk.Frame):
    """
    Widget pour sélectionner UN OU PLUSIEURS techniciens sur une intervention.

    Composé d'une SearchableCombobox pour choisir un nom + bouton "+ Ajouter",
    et d'une zone de "chips" en dessous : chaque chip représente un technicien
    sélectionné, avec un bouton ✕ pour le retirer.

    API :
      .get_names() → liste de noms
      .set_names(list) → définit la sélection
      .get() / .set(str) → version compat (chaîne CSV)
      .set_available(list) → met à jour la liste des techniciens disponibles
    """
    def __init__(self, master, available=None, on_add_new=None, on_change=None, **kw):
        super().__init__(master, bg=master.cget("bg") if "bg" not in kw else kw.get("bg"))
        self._available = list(available or [])
        self._selected  = []
        self._on_add_new = on_add_new  # callback : ouvre TechnicienDialog
        self._on_change_cb = on_change  # callback(names) quand la sélection change

        # Ligne 1 : combobox + boutons
        row1 = tk.Frame(self, bg=self.cget("bg"))
        row1.pack(fill="x")
        self._combo_var = tk.StringVar()
        self._combo = SearchableCombobox(row1, textvariable=self._combo_var,
                                          values=self._available, width=38)
        self._combo.pack(side="left", fill="x", expand=True)
        self._combo.bind("<<ComboboxSelected>>", self._on_combo_selected)
        # Entrée valide une saisie (ajoute si le nom existe ou propose création)
        self._combo.entry.bind("<Return>", self._on_combo_enter, add="+")
        mk_btn(row1, "+", self._add_from_combo, color=C["btn2"]).pack(side="left", padx=(6, 0))
        if on_add_new is not None:
            mk_btn(row1, "Nouveau", self._on_add_new_clicked, color=C["btn2"]
                   ).pack(side="left", padx=(4, 0))

        # Ligne 2 : zone des chips
        self._chips_frame = tk.Frame(self, bg=self.cget("bg"))
        self._chips_frame.pack(fill="x", pady=(6, 0))

        self._render_chips()

    # ─── API publique ────────────────────────────────────────────────────────
    def get_names(self):
        return list(self._selected)

    def set_names(self, names):
        self._selected = [n for n in (names or []) if n and n.strip()]
        # Dédoublonner en conservant l'ordre
        seen = set()
        out = []
        for n in self._selected:
            if n not in seen:
                seen.add(n)
                out.append(n)
        self._selected = out
        self._render_chips()

    def get(self):
        return db.format_techniciens(self._selected)

    def set(self, s):
        self.set_names(db.parse_techniciens(s))

    def set_available(self, available):
        self._available = list(available or [])
        self._combo.set_values(self._available)

    def configure(self, **kw):
        fg = kw.pop("foreground", kw.pop("fg", None))
        if fg is not None:
            self._combo.configure(foreground=fg)
        if kw:
            super().configure(**kw)
    config = configure

    # ─── Internes ────────────────────────────────────────────────────────────
    def _on_combo_selected(self, _ev=None):
        # Quand l'utilisateur valide via la dropdown : on ajoute directement
        self._add_from_combo()

    def _on_combo_enter(self, _ev=None):
        # Entrée dans le champ : tenter d'ajouter (même comportement que +)
        self._add_from_combo()
        return "break"

    def _add_from_combo(self):
        name = self._combo_var.get().strip()
        if not name:
            return
        # Résolution tolérante :
        # 1. Match exact dans l'annuaire → OK
        # 2. Sinon, si la saisie matche UN SEUL technicien (insensible
        #    casse/accents, préfixe ou contenu) → on prend celui-là
        resolved = None
        if name in self._available:
            resolved = name
        else:
            def _norm(s):
                s = str(s).lower()
                return s.translate(str.maketrans(
                    "àâäéèêëîïôöùûüç", "aaaeeeeiioouuuc"))
            q = _norm(name)
            matches = [v for v in self._available if q in _norm(v)]
            if len(matches) == 1:
                resolved = matches[0]
            elif len(matches) > 1:
                # Plusieurs candidats : tenter un match préfixe exact
                pref = [v for v in matches if _norm(v).startswith(q)]
                if len(pref) == 1:
                    resolved = pref[0]

        if resolved is None:
            messagebox.showwarning(
                "Technicien inconnu",
                f"'{name}' ne correspond à aucun technicien de l'annuaire.\n\n"
                "Choisissez un nom dans la liste déroulante, ou cliquez sur "
                "« Nouveau » pour créer ce technicien.")
            return
        if resolved in self._selected:
            self._combo_var.set("")
            return
        self._selected.append(resolved)
        self._combo_var.set("")
        self._render_chips()
        self._fire_change()

    def _on_add_new_clicked(self):
        # Délègue à la fonction passée (qui ouvre TechnicienDialog
        # avec un callback pour rafraîchir la liste et auto-sélectionner)
        if self._on_add_new is not None:
            self._on_add_new(self._after_new_tech)

    def _after_new_tech(self, new_name=None):
        """Appelé par le dialogue de création de tech, avec le nom créé."""
        # Recharger la liste depuis la DB
        self._available = [t["nom"] for t in db.get_techniciens()]
        self._combo.set_values(self._available)
        if new_name and new_name not in self._selected and new_name in self._available:
            self._selected.append(new_name)
            self._render_chips()
            self._fire_change()

    def _remove(self, name):
        if name in self._selected:
            self._selected.remove(name)
            self._render_chips()
            self._fire_change()

    def _fire_change(self):
        if self._on_change_cb is not None:
            self._on_change_cb(list(self._selected))

    def _render_chips(self):
        for w in self._chips_frame.winfo_children():
            w.destroy()
        if not self._selected:
            tk.Label(self._chips_frame,
                     text="Aucun technicien · choisissez dans la liste ci-dessus",
                     bg=self.cget("bg"), fg=C["text_muted"],
                     font=("Segoe UI", 9, "italic")).pack(anchor="w")
            return
        # Une ligne de chips qui wrappent (via grid auto)
        for i, name in enumerate(self._selected):
            chip = tk.Frame(self._chips_frame, bg=C["nav_sel"], bd=0)
            chip.pack(side="left", padx=(0, 5), pady=2)
            tk.Label(chip, text=name, bg=C["nav_sel"], fg="white",
                     font=("Segoe UI", 9, "bold"),
                     padx=8, pady=3).pack(side="left")
            close = tk.Label(chip, text=" ✕ ", bg=C["nav_sel"], fg="#cce0f5",
                              font=("Segoe UI", 9, "bold"),
                              cursor="hand2", padx=2, pady=3)
            close.pack(side="left")
            close.bind("<Button-1>", lambda _e, n=name: self._remove(n))
            close.bind("<Enter>", lambda _e, w=close: w.configure(fg="white"))
            close.bind("<Leave>", lambda _e, w=close: w.configure(fg="#cce0f5"))


# ══════════════════════════════════════════════════════════════════════════════
# FENÊTRE PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════
class AppEMS(tk.Tk):
    # Onglets visibles selon le mode de l'application
    NAV_PARELEMENTS = {
        "bons": ["dashboard", "interventions", "import_export", "nouveau"],
        "parc": ["clients", "moteurs", "techniciens"],
        "full": ["dashboard", "interventions", "clients",
                 "moteurs", "techniciens", "import_export", "nouveau"],
    }
    TITRES = {
        "bons": "EMS – Bons d'intervention",
        "parc": "EMS – Gestion de parc",
        "full": "EMS – Gestion des Interventions",
    }

    def __init__(self, mode="full"):
        super().__init__()
        self.withdraw()
        apply_icon(self)
        self.mode = mode if mode in self.NAV_PARELEMENTS else "full"
        self._onglets_actifs = self.NAV_PARELEMENTS[self.mode]
        self.title(self.TITRES[self.mode])
        self.geometry("1340x830")
        self.minsize(960, 620)
        self.configure(bg=C["bg"])
        self._init_ttk_styles()
        db.init_db()
        self._build()
        self.update_idletasks()
        self.deiconify()
        # Différer le premier refresh : la fenêtre s'affiche avant tout appel API
        self.after(50, self._initial_show)

    def _initial_show(self):
        try:
            self.show(self._onglets_actifs[0])
        except Exception:
            pass

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

        # ── Zone logo : carte blanche pour que le fond du PNG s'intègre ──────
        # Le logo officiel EMS a un fond blanc, on le pose dans un encart blanc
        # qui devient une "carte d'identité" visuelle et cohérente.
        logo_zone = tk.Frame(self.sidebar, bg=C["header"])
        logo_zone.pack(fill="x", pady=(14, 8))

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
                from shared.logo_data import LOGO_EMS_B64
                if LOGO_EMS_B64:
                    self._logo_img = tk.PhotoImage(data=LOGO_EMS_B64)
                    logo_loaded = True
            except (ImportError, tk.TclError):
                pass

        if logo_loaded and self._logo_img is not None:
            # Cible : logo dans une carte de ~150px de large (sidebar 210 moins marges)
            target_w = 150
            w = self._logo_img.width()
            h = self._logo_img.height()
            if w > target_w:
                # 1. Tenter PIL/Pillow pour un resize de qualité (antialiasing)
                try:
                    from PIL import Image, ImageTk
                    from pathlib import Path as _P
                    src = _P(__file__).parent / "assets" / "logo_ems.png"
                    if src.is_file():
                        im = Image.open(src)
                    else:
                        # Reconstruire depuis le base64 embarqué
                        import base64, io
                        from shared.logo_data import LOGO_EMS_B64
                        im = Image.open(io.BytesIO(base64.b64decode(LOGO_EMS_B64)))
                    ratio = target_w / im.width
                    new_size = (target_w, int(im.height * ratio))
                    im_resized = im.resize(new_size, Image.LANCZOS)
                    self._logo_img = ImageTk.PhotoImage(im_resized)
                except (ImportError, Exception):
                    # 2. Fallback Tkinter natif : subsample (arrondi vers le haut)
                    #    pour s'assurer que la largeur finale est <= target_w
                    import math
                    factor = max(2, math.ceil(w / target_w))
                    try:
                        self._logo_img = self._logo_img.subsample(factor, factor)
                    except tk.TclError:
                        pass

            # Encart blanc qui accueille le logo : bordure douce + padding
            card_outer = tk.Frame(logo_zone, bg=C["accent"])
            card_outer.pack(padx=14, pady=(2, 4))
            card = tk.Frame(card_outer, bg="white")
            card.pack(padx=1, pady=1)
            tk.Frame(card, bg=C["accent"], height=2).pack(fill="x")
            tk.Label(card, image=self._logo_img, bg="white",
                     bd=0).pack(padx=12, pady=(10, 8))
            tk.Label(logo_zone, text="Emeraude Moteurs Systèmes",
                     font=("Segoe UI", 9, "bold"), bg=C["header"], fg="white",
                     justify="center").pack(pady=(4, 0))
            tk.Label(logo_zone, text="L'application motorisée,\nau cœur de vos énergies",
                     font=F["tiny"], bg=C["header"], fg="#aac4e8",
                     justify="center").pack(pady=(0, 2))
        else:
            # Fallback texte uniquement : pas besoin de carte blanche
            tk.Label(logo_zone, text="EMS", font=("Segoe UI", 28, "bold"),
                     bg=C["header"], fg="white").pack(pady=(8, 0))
            tk.Label(logo_zone, text="Emeraude Moteurs Systèmes",
                     font=F["tiny"], bg=C["header"], fg="#aac4e8",
                     justify="center").pack(pady=(2, 0))

        # Liseré rouge accent sous la zone logo
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
            ("import_export", "🔄", "Import / Export"),
            ("nouveau",       "➕", "Nouveau bon"),
        ]:
            if key not in self._onglets_actifs:
                continue
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
                 bg=C["header"], fg="#6699cc").pack(side="bottom", pady=(0, 8))

        # Indicateur de connexion
        self._conn_lbl = tk.Label(self.sidebar, text="● Vérification…",
                                   font=F["tiny"], bg=C["header"], fg="#6699cc")
        self._conn_lbl.pack(side="bottom", pady=(4, 0))
        self.after(500, self._update_conn)

        # Zone principale
        self.main = tk.Frame(self, bg=C["bg"])
        self.main.pack(side="left", fill="both", expand=True)

        self.frames = {}
        for FrameCls, key in [
            (DashboardFrame,      "dashboard"),
            (InterventionsFrame,  "interventions"),
            (ClientsFrame,        "clients"),
            (MoteursFrame,        "moteurs"),
            (TechniciensFrame,    "techniciens"),
            (ImportExportFrame,   "import_export"),
            (NouveauFrame,        "nouveau"),
        ]:
            if key not in self._onglets_actifs:
                continue
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

    def goto(self, key):
        """Affiche un onglet seulement s'il existe dans ce mode."""
        if key in self.frames:
            self.show(key)
            return True
        return False

    def _update_conn(self):
        """Met à jour l'indicateur de connexion en thread de fond (toutes les 20 s)."""
        import threading
        def _bg():
            en_ligne = db.ping(timeout=2.0)
            try:
                self.after(0, lambda: self._conn_lbl.config(
                    text="● En ligne" if en_ligne else "○ Hors ligne",
                    fg="#6fd46f" if en_ligne else "#f0a0a0"))
            except Exception:
                pass
        threading.Thread(target=_bg, daemon=True).start()
        self.after(20000, self._update_conn)

    def refresh_frame(self, key):
        """Rafraîchit un onglet s'il existe (sécurisé multi-mode)."""
        fr = self.frames.get(key)
        if fr is not None and hasattr(fr, "refresh"):
            fr.refresh()

    def _garage_search(self):
        ns = self.garage_var.get().strip()
        if not ns:
            messagebox.showinfo("Recherche", "Saisissez un N° de série.")
            return
        m = db.find_moteur_by_serie(ns)
        if not m:
            if self.goto("moteurs"):
                self.frames["moteurs"].serie_only_var.set(1)
                self.frames["moteurs"].search_var.set(ns)
            else:
                messagebox.showinfo("Recherche", f"Moteur '{ns}' introuvable.")
            return
        if self.goto("interventions"):
            self.frames["interventions"].search_var.set(ns)
            self.frames["interventions"].statut_var.set("Tous")
        elif self.goto("moteurs"):
            self.frames["moteurs"].serie_only_var.set(1)
            self.frames["moteurs"].search_var.set(ns)
        self.garage_var.set("")


# ══════════════════════════════════════════════════════════════════════════════
# WIDGETS DU TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════════════════════
# Chaque widget est un Frame avec une méthode refresh().
# Le DashboardFrame instancie/cache les widgets selon la config utilisateur.

# Catalogue des widgets disponibles
WIDGET_CATALOG = [
    ("stats_cards",        "Cartes statistiques",    "Cartes synthétiques (En cours, Clos, etc.)"),
    ("urgentes",           "Interventions urgentes", "Liste des bons Urgente / Critique en cours"),
    ("a_programmer",       "Bons à programmer",      "Liste des bons au statut 'Date à programmer'"),
    ("a_facturer",         "Bons à facturer",         "Liste des bons au statut 'À facturer'"),
    ("activite_recente",   "Activité récente",        "Derniers bons modifiés"),
    ("garantie_expirante", "Garanties expirantes",    "Moteurs dont la garantie expire bientôt"),
    ("par_technicien",     "Charge par technicien",   "Nombre d'interventions par technicien"),
    ("par_type",           "Répartition par type",    "Statistiques par type d'intervention"),
    ("non_notifies",       "Non notifiés",            "Bons en cours sans notification envoyée"),
    ("classifications",    "Par classification",      "Garantie / Facturable / Interne"),
]

# Catalogue des cartes statistiques
CARD_CATALOG = [
    # (clé, label affiché, couleur d'accent latéral, couleur du chiffre)
    ("En cours",     "En cours",      C["ec"],          C["ec"]),
    ("Date à programmer", "À programmer", C["prog"],    C["prog"]),
    ("À facturer",   "À facturer",    C["afact"],       C["afact"]),
    ("Facturé",      "Facturé",       C["fact"],        C["fact"]),
    ("Clos",         "Clos",          C["clos"],        C["clos"]),
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
            "En cours": stats["En cours"],
            "Date à programmer": stats.get("Date à programmer", 0),
            "À facturer": stats["À facturer"],
            "Facturé": stats["Facturé"], "Clos": stats["Clos"],
            "Total": stats["Total"],
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
        tk.Label(self, text="Interventions urgentes en cours",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["danger"]).pack(anchor="w")
        cols = ("urg","num_bon","client","navire","tech","date")
        col_defs = [("urg","⚡",70),("num_bon","N° Bon",115),("client","Client",170),
                    ("navire","Navire",140),("tech","Tech.",110),("date","Date",90)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
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
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
        except (ValueError, IndexError):
            pass


class AProgrammerWidget(tk.Frame):
    """Liste des bons au statut 'Date à programmer'."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="Bons à programmer",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["prog"]).pack(anchor="w")
        cols = ("urg","num_bon","client","navire","tech","date")
        col_defs = [("urg","⚡",70),("num_bon","N° Bon",115),("client","Client",170),
                    ("navire","Navire",140),("tech","Tech.",110),("date","Date",90)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
        tf.pack(fill="x", pady=(2, 0))
        self._cache = []
        self.tree.bind("<Double-1>", self._open)

    def refresh(self):
        invs = db.get_interventions_a_programmer(limit=8)
        self._cache = list(invs)
        rows  = []
        urgs  = []
        for r in self._cache:
            rows.append((r["urgence"], r["num_bon"], r["client_nom"] or "",
                         r["navire"] or "", r["technicien"], r["date_creation"]))
            urgs.append(r["urgence"])
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("—", "(aucun)", "", "", "", ""))
        else:
            fill_tree(self.tree, rows, urgences=urgs)

    def _open(self, _ev):
        sel = self.tree.selection()
        if not sel or not self._cache: return
        try:
            r = self._cache[int(sel[0])]
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
        except (ValueError, IndexError):
            pass


class AFacturerWidget(tk.Frame):
    """Liste des bons au statut 'À facturer'."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="Bons à facturer",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["afact"]).pack(anchor="w")
        cols = ("urg","num_bon","client","navire","tech","date")
        col_defs = [("urg","⚡",70),("num_bon","N° Bon",115),("client","Client",170),
                    ("navire","Navire",140),("tech","Tech.",110),("date","Date",90)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
        tf.pack(fill="x", pady=(2, 0))
        self._cache = []
        self.tree.bind("<Double-1>", self._open)

    def refresh(self):
        invs = db.get_interventions_a_facturer(limit=20)
        self._cache = list(invs)
        rows  = []
        urgs  = []
        for r in self._cache:
            rows.append((r["urgence"], r["num_bon"], r["client_nom"] or "",
                         r["navire"] or "", r["technicien"], r["date_creation"]))
            urgs.append(r["urgence"])
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("—", "(aucun)", "", "", "", ""))
        else:
            fill_tree(self.tree, rows, urgences=urgs)

    def _open(self, _ev):
        sel = self.tree.selection()
        if not sel or not self._cache: return
        try:
            r = self._cache[int(sel[0])]
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
        except (ValueError, IndexError):
            pass


class ActiviteRecenteWidget(tk.Frame):
    """Derniers bons modifiés."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="Activité récente",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("num_bon","client","statut","tech","modif")
        col_defs = [("num_bon","N° Bon",115),("client","Client",180),
                    ("statut","Statut",80),("tech","Tech.",110),
                    ("modif","Modifié (Paris)",120)]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
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
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
        except (ValueError, IndexError):
            pass


class GarantieExpiranteWidget(tk.Frame):
    """Moteurs dont la garantie expire dans <= 90 jours."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        head = tk.Frame(self, bg=C["bg"])
        head.pack(fill="x")
        tk.Label(head, text="Garanties expirant dans 90 jours",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["warn"]).pack(side="left")
        cols = ("ns","client","navire","fin","jours")
        col_defs = [("ns","N° Série",130),("client","Client",170),
                    ("navire","Navire/Site",140),("fin","Mise svc",95),
                    ("jours","Jours restants",110, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
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
        tk.Label(self, text="Charge par technicien",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("tech","ec","afact","fact","clos","total")
        col_defs = [("tech","Technicien",180),
                    ("ec",   "En cours",   75, "center"),
                    ("afact","À facturer", 80, "center"),
                    ("fact", "Facturé",    75, "center"),
                    ("clos", "Clos",       70, "center"),
                    ("total","Total",      65, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
        tf.pack(fill="x", pady=(2, 0))

    def refresh(self):
        rows = db.get_stats_par_technicien()
        if not rows:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", values=("(aucun technicien)", "", "", "", "", ""))
        else:
            data = [(r["technicien"], r["en_cours"], r["a_facturer"],
                     r["facture"], r["clos"], r["total"])
                    for r in rows]
            fill_tree(self.tree, data)


class ParTypeWidget(tk.Frame):
    """Répartition par type d'intervention."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="Répartition par type d'intervention",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["header"]).pack(anchor="w")
        cols = ("type","ec","total")
        col_defs = [("type","Type",260),
                    ("ec","En cours",80, "center"),
                    ("total","Total",80, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
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
        tk.Label(self, text="Bons en cours non notifiés",
                 font=("Arial", 11, "bold"), bg=C["bg"], fg=C["warn"]).pack(anchor="w")
        cols = ("num_bon","client","cli","tech","tec")
        col_defs = [("num_bon","N° Bon",115),("client","Client",200),
                    ("cli","Client notifié",110, "center"),
                    ("tech","Technicien",130),
                    ("tec","Tech. notifié",110, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5, show_hsb=False)
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
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
        except (ValueError, IndexError):
            pass


class ClassificationsWidget(tk.Frame):
    """Mini-vue Garantie / Facturable / Interne."""
    def __init__(self, master, app):
        super().__init__(master, bg=C["bg"])
        self.app = app
        tk.Label(self, text="Répartition par classification",
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
    "a_programmer":       AProgrammerWidget,
    "a_facturer":         AFacturerWidget,
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

        # Zone scrollable
        scroll_outer = tk.Frame(parent, bg=C["bg"])
        scroll_outer.pack(fill="both", expand=True, padx=6, pady=4)
        self._list_canvas = tk.Canvas(scroll_outer, bg=C["bg"], highlightthickness=0)
        _vsb = ttk.Scrollbar(scroll_outer, orient="vertical", command=self._list_canvas.yview)
        self._list_canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        self._list_canvas.pack(side="left", fill="both", expand=True)
        self.list_frame = tk.Frame(self._list_canvas, bg=C["bg"])
        _win = self._list_canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>",
            lambda e: self._list_canvas.configure(scrollregion=self._list_canvas.bbox("all")))
        self._list_canvas.bind("<Configure>",
            lambda e: self._list_canvas.itemconfig(_win, width=e.width))

        def _on_wheel(ev):
            try: self._list_canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            except tk.TclError: pass
        self._list_canvas.bind("<Enter>", lambda e: self._list_canvas.bind_all("<MouseWheel>", _on_wheel))
        self._list_canvas.bind("<Leave>", lambda e: self._list_canvas.unbind_all("<MouseWheel>"))

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
                "stats_cards", "urgentes", "a_programmer", "activite_recente",
                "par_technicien", "non_notifies"])
            db.set_dashboard_cards([
                "En cours", "Date à programmer", "À facturer", "Facturé", "Clos", "Total",
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
        self.canvas.bind("<MouseWheel>", _on_wheel)
        self._on_wheel = _on_wheel

        self._widgets = {}  # key → widget instance

    def _open_config(self):
        DashboardConfigDialog(self, self.app, on_save=self._rebuild)

    def _rebuild(self):
        """Reconstruit les widgets selon la config actuelle."""
        for w in self.inner.winfo_children():
            w.destroy()
        self._widgets.clear()
        offline = False
        try:
            active = db.get_dashboard_widgets()
        except Exception:
            active = []
            offline = True

        if not active:
            msg = ("Serveur inaccessible — tableau de bord indisponible hors ligne."
                   if offline else
                   "Aucun widget configuré. Cliquez sur ⚙ Configurer.")
            tk.Label(self.inner, text=msg, bg=C["bg"], fg="#888",
                     font=("Arial", 11), justify="center").pack(pady=60)
            return  # évite la récursion refresh() → _rebuild()

        for key in active:
            cls = WIDGET_CLASSES.get(key)
            if not cls:
                continue
            wrap = tk.Frame(self.inner, bg=C["bg"])
            wrap.pack(fill="x", padx=20, pady=(8, 4))
            w = cls(wrap, self.app)
            w.pack(fill="x")
            self._widgets[key] = w
        if self._widgets:
            self._patch_wheels(self.inner)
            self.refresh()

    def _patch_wheels(self, widget):
        """Bind <MouseWheel> sur tous les descendants pour scroller le canvas (et bloquer le scroll natif des Treeview)."""
        on_wheel = self._on_wheel
        for child in widget.winfo_children():
            child.bind("<MouseWheel>", lambda ev, f=on_wheel: (f(ev), "break")[1])
            self._patch_wheels(child)

    def refresh(self):
        if not self._widgets:
            self._rebuild()
            return
        import threading
        widgets = list(self._widgets.values())
        def _bg():
            for w in widgets:
                try:
                    w.refresh()
                except Exception:
                    pass
        threading.Thread(target=_bg, daemon=True).start()


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

        # Bouton export base centrale, en haut a droite (dissocie des actions du bas)
        #mk_btn(sf, "📤 Exporter base centrale", self._exporter_central,
         #      color=C["btn2"]).pack(side="right")

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
        mk_btn(af, "✍ Faire signer",     self._signer, color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "📄 Générer bon",     self._generer).pack(side="left", padx=3)
        mk_btn(af, "📁 Dossier",         self._dossier).pack(side="left", padx=3)
        mk_btn(af, "📧 Prévenir client", self._mail_client,   color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "📧 Prévenir tech.",  self._mail_tech,    color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "✅ Mail clôture",    self._mail_cloture, color=C["btn2"]).pack(side="left", padx=3)
        mk_btn(af, "🗑️ Supprimer",       self._supprimer,   color=C["danger"]).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        import threading
        statut  = self.statut_var.get()
        urgence = self.urgence_var.get()
        search  = self.search_var.get()

        def _bg():
            try:
                invs = list(db.get_interventions(statut=statut,
                                                  urgence=urgence, search=search))
            except Exception:
                invs = None
            try:
                self.after(0, lambda: self._apply_refresh(invs))
            except Exception:
                pass

        threading.Thread(target=_bg, daemon=True).start()

    def _apply_refresh(self, invs):
        if invs is None:
            self.tree.delete(*self.tree.get_children())
            self.tree.insert("", "end", iid="0",
                             values=("", "Hors ligne", "", "Serveur inaccessible",
                                     "", "", "", "", "", ""))
            self._cache = []
            return
        self._cache = invs
        rows, urgs = [], []
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

    def _signer(self):
        r = self._sel()
        if r:
            SignatureDialog(self, self.app, inv_id=r["id"],
                            on_save=self.refresh)

    def _generer(self):
        r = self._sel()
        if not r: return
        inv = db.get_intervention(inv_id=r["id"])
        dossier = _dossiers_root() / inv["num_bon"]

        def _faire_gen(photos):
            if photos is None:
                return
            try:
                path = sauvegarder_bon(inv, photos_annexe=photos, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc))
                return
            ouvrir_fichier(path)

        PhotosAnnexeDialog(self, dossier, on_valider=_faire_gen)
    def _exporter_central(self):
        """Exporte le bon sélectionné vers la base centrale (serveur atelier)."""
        r = self._sel()
        if not r:
            return
        try:
            from ems_client import sync_client
        except ImportError:
            messagebox.showerror("Module manquant",
                "Le module de synchronisation n'est pas disponible.")
            return

        if not sync_client.serveur_central_joignable():
            messagebox.showwarning("Serveur central injoignable",
                "Impossible de joindre la base centrale.\n\n"
                "Vérifiez votre connexion au réseau de l'atelier, "
                "puis réessayez. Le bon reste enregistré localement.")
            return

        res = sync_client.exporter_bon(r["id"])
        status = res["status"]
        if status == sync_client.EXPORT_OK:
            messagebox.showinfo("Export réussi", res["message"])
        elif status == sync_client.EXPORT_CONFLIT:
            if messagebox.askyesno("Conflit détecté",
                    res["message"] + "\n\n"
                    "• OUI : votre version remplacera celle du bureau\n"
                    "• NON : annuler l'export (rien n'est modifié)"):
                res2 = sync_client.exporter_bon(r["id"], force=True)
                if res2["status"] == sync_client.EXPORT_OK:
                    messagebox.showinfo("Export réussi (forcé)", res2["message"])
                else:
                    messagebox.showerror("Échec", res2["message"])
        elif status == sync_client.EXPORT_HORS_LIGNE:
            messagebox.showwarning("Hors ligne", res["message"])
        else:
            messagebox.showerror("Erreur d'export", res["message"])


    def _dossier(self):
        r = self._sel()
        if not r: return
        d = _dossiers_root() / r["num_bon"]
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    def _mail_client(self):
        r = self._sel()
        if not r: return
        inv = db.get_intervention(inv_id=r["id"])
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        email = (row_get(inv, "email_demandeur") or
                 row_get(inv, "email_signataire") or
                 row_get(client, "email") if client else "")
        if not email:
            messagebox.showwarning("Email manquant",
                "Aucun email renseigné (demandeur, signataire ou client).")
            return
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_client(inv, client, moteur, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")
        db.mark_notifie(r["id"], "client")
        self.refresh()

    def _mail_tech(self):
        r = self._sel()
        if not r: return
        inv = db.get_intervention(inv_id=r["id"])
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        # Collecter les emails de TOUS les techniciens assignés
        tech_names = db.parse_techniciens(row_get(inv, "technicien"))
        emails = []
        for n in tech_names:
            t = db.get_technicien_by_nom(n)
            em = row_get(t, "email")
            if em:
                emails.append(em)
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_technicien(inv, client, moteur, emails, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")
        db.mark_notifie(r["id"], "tech")
        self.refresh()

    def _mail_cloture(self):
        r = self._sel()
        if not r: return
        inv    = db.get_intervention(inv_id=r["id"])
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        # Vérification : au moins un destinataire
        if not row_get(inv, "email_demandeur") and not row_get(inv, "email_signataire"):
            messagebox.showwarning("Destinataires manquants",
                "Aucun email demandeur ni signataire renseigné sur ce bon.")
            return
        # Emails des techniciens en CC
        tech_emails = []
        for n in db.parse_techniciens(row_get(inv, "technicien")):
            t = db.get_technicien_by_nom(n)
            em = row_get(t, "email")
            if em:
                tech_emails.append(em)
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_cloture(inv, client, moteur, tech_emails, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")

    def _supprimer(self):
        r = self._sel()
        if not r:
            return
        num_bon = r["num_bon"]
        if not messagebox.askyesno("Supprimer", f"Supprimer le bon {num_bon} ?"):
            return
        # Proposer la suppression du dossier si celui-ci existe
        dossier = _dossiers_root() / num_bon
        if dossier.exists():
            fichiers = list(dossier.rglob("*"))
            nb = sum(1 for f in fichiers if f.is_file())
            suppr_dossier = messagebox.askyesno(
                "Supprimer le dossier ?",
                f"Le dossier {num_bon} contient {nb} fichier(s).\n\n"
                "Voulez-vous aussi supprimer le dossier et tous ses fichiers "
                "(PDFs, photos, etc.) ?\n\n"
                "⚠ Cette action est irréversible.")
            if suppr_dossier:
                import shutil
                shutil.rmtree(dossier, ignore_errors=True)
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
        self.tree.bind("<Double-1>", lambda e: self._voir_moteurs())
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

    def _voir_moteurs(self):
        r = self._sel()
        if r:
            ClientMoteursDialog(self, self.app, dict(r))

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
        mk_header(self, "Moteurs", "   Recherche par N° série, navire, marque, code affaire…")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 Recherche :",
                 bg=C["bg"], font=("Arial",10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=35).pack(side="left", padx=4)
        # Option : limiter la recherche au N° de série uniquement
        self.serie_only_var = tk.IntVar(value=0)
        tk.Checkbutton(sf, text="N° série uniquement", variable=self.serie_only_var,
                       bg=C["bg"], font=("Arial", 9), fg=C["text_muted"],
                       command=self.refresh).pack(side="left", padx=(8, 0))
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
        tf, self.tree = mk_tree(self, cols, col_defs, height=22, show_tree=True)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=4)
        mk_btn(af, "📋 Fiche & historique", self._voir_inv).pack(side="left", padx=4)
        mk_btn(af, "➕ Sous-ensemble", self._se_add,
               color=C["btn2"]).pack(side="left", padx=4)
        self.btn_suppr = mk_btn(af, "🗑️ Supprimer",
                                 self._supprimer, color=C["danger"])
        self.btn_suppr.pack(side="left", padx=4)
        self.sel_info = tk.Label(af, text="", bg=C["bg"],
                                  fg=C["text_muted"], font=F["small"])
        self.sel_info.pack(side="left", padx=(12, 0))
        self.tree.bind("<Double-1>", lambda e: None if self.tree.identify_column(e.x) == "#0" else self._modifier())
        self.tree.bind("<<TreeviewSelect>>", self._on_select_change)
        self.tree.bind("<<TreeviewOpen>>", self._on_tree_open)
        self._cache = []
        self._by_id = {}
        self._loaded_children = set()

    def _row_values(self, m):
        statut, jours = db.garantie_status(m.get("date_mise_service", ""),
                                            m.get("duree_garantie", ""))
        if statut == "Active":   gtxt = f"Active · {jours}j"
        elif statut == "Expirée": gtxt = f"Expirée · -{jours}j"
        else: gtxt = "—"
        return (m.get("num_serie", ""), m.get("client_nom", "") or "",
                row_get(m, "navire"), row_get(m, "marque"), row_get(m, "machine"),
                row_get(m, "type_moteur"), m.get("date_mise_service", ""),
                m.get("duree_garantie", ""), gtxt)

    def refresh(self):
        ms = db.get_moteurs(search=self.search_var.get(),
                             serie_only=bool(self.serie_only_var.get()))
        self._cache = list(ms)
        self._by_id = {m["id"]: m for m in self._cache}
        self._loaded_children = set()
        self.tree.delete(*self.tree.get_children())
        for i, m in enumerate(self._cache):
            tags = ["even" if i % 2 == 0 else "odd"]
            self.tree.insert("", "end", iid=m["id"],
                              values=self._row_values(m), tags=tags)
            if m.get("nb_sous_ensembles", 0) > 0:
                # Nœud factice : juste pour afficher la flèche ▶, remplacé
                # par les vrais sous-ensembles au premier déploiement.
                self.tree.insert(m["id"], "end",
                                  values=("⏳ chargement…", "", "", "", "", "", "", "", ""))

    def _on_tree_open(self, _ev=None):
        iid = self.tree.focus()
        if iid in self._by_id and iid not in self._loaded_children:
            self._load_children(iid)

    def _load_children(self, parent_id):
        self._loaded_children.add(parent_id)
        self.tree.delete(*self.tree.get_children(parent_id))
        try:
            children = list(db.get_sous_ensembles(parent_id))
        except Exception as e:
            children = []
            print(f"[MoteursFrame] sous-ensembles : {e}")
        if not children:
            self.tree.insert(parent_id, "end",
                              values=("(aucun sous-ensemble)", "", "", "", "", "", "", "", ""))
            return
        for j, se in enumerate(children):
            self._by_id[se["id"]] = se
            tags = ["se_even" if j % 2 == 0 else "se_odd"]
            self.tree.insert(parent_id, "end", iid=se["id"],
                              values=self._row_values(se), tags=tags)

    def _on_select_change(self, _ev=None):
        sel = self.tree.selection()
        n = len(sel)
        if n <= 1:
            self.sel_info.config(text="")
            self.btn_suppr.config(text="🗑️ Supprimer")
        else:
            self.sel_info.config(
                text=f"{n} éléments sélectionnés "
                     "(Ctrl/Maj-clic pour ajuster)")
            self.btn_suppr.config(text=f"🗑️ Supprimer ({n})")

    def _se_add(self):
        r = self._sel()
        if not r:
            return
        if r.get("parent_moteur_id"):
            messagebox.showwarning(
                "Sous-ensemble",
                "Impossible d'ajouter un sous-ensemble à un sous-ensemble.\n"
                "Sélectionnez un moteur principal.")
            return
        def _after_save():
            self.refresh()
            if self.tree.exists(r["id"]):
                self.tree.item(r["id"], open=True)
                self._load_children(r["id"])
        MoteurDialog(self, self.app, parent_moteur=r, on_save=_after_save)

    def _sel(self):
        """Premier élément sélectionné (moteur ou sous-ensemble)."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un moteur."); return None
        return self._by_id.get(sel[0])

    def _sel_multi(self):
        """Liste des éléments sélectionnés (Ctrl/Shift click pris en charge).
        Si un moteur et l'un de ses sous-ensembles sont sélectionnés en même
        temps, seul le moteur est conservé (la suppression cascade)."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning(
                "Sélection",
                "Sélectionnez un ou plusieurs moteurs.\n"
                "(Ctrl-clic ou Maj-clic pour en sélectionner plusieurs)")
            return []
        items = [self._by_id[s] for s in sel if s in self._by_id]
        ids = {m["id"] for m in items}
        return [m for m in items if m.get("parent_moteur_id") not in ids]

    def _modifier(self):
        r = self._sel()
        if r: MoteurDialog(self, self.app, dict(r), on_save=self.refresh)

    def _voir_inv(self):
        r = self._sel()
        if not r:
            return
        MoteurFicheDialog(self, self.app, dict(r))

    def _supprimer(self):
        moteurs = self._sel_multi()
        if not moteurs:
            return
        n = len(moteurs)
        if n == 1:
            msg = (f"Supprimer le moteur « {moteurs[0]['num_serie']} » ?\n\n"
                   "Cette action est irréversible.")
        else:
            apercu = "\n".join(f"  • {m['num_serie']}" for m in moteurs[:8])
            if n > 8:
                apercu += f"\n  … et {n - 8} autre(s)"
            msg = (f"⚠ Vous allez supprimer {n} moteurs :\n\n"
                   f"{apercu}\n\n"
                   "Cette action est IRRÉVERSIBLE. Continuer ?")
        if not messagebox.askyesno("Confirmer la suppression", msg,
                                    icon="warning"):
            return
        echecs = []
        for m in moteurs:
            try:
                db.delete_moteur(m["id"])
            except Exception as e:
                echecs.append(f"{m['num_serie']} : {e}")
        self.refresh()
        if echecs:
            messagebox.showerror(
                "Suppression partielle",
                f"{n - len(echecs)}/{n} moteurs supprimés.\n\n"
                "Échecs :\n" + "\n".join(echecs[:10]))
        else:
            messagebox.showinfo(
                "Suppression effectuée",
                f"✅ {n} moteur(s) supprimé(s).")


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
# GARANTIES MOTEURS
# ══════════════════════════════════════════════════════════════════════════════
class GarantiesFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Garanties",
                  "   Suivi des dossiers de garantie moteur")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 Recherche :", bg=C["bg"], font=("Arial", 10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=24).pack(side="left", padx=4)

        tk.Label(sf, text="  Statut :", bg=C["bg"], font=("Arial", 10)).pack(side="left", padx=(12, 4))
        self.statut_var = tk.StringVar(value="Tous")
        self.statut_combo = ttk.Combobox(sf, textvariable=self.statut_var,
                     values=["Tous"] + db.get_statuts_garantie(), width=22,
                     state="readonly")
        self.statut_combo.pack(side="left")
        self.statut_var.trace_add("write", lambda *_: self.refresh())

        tk.Label(sf, text="  Attribution :", bg=C["bg"], font=("Arial", 10)).pack(side="left", padx=(12, 4))
        self.attr_var = tk.StringVar(value="Toutes")
        self.attr_combo = ttk.Combobox(sf, textvariable=self.attr_var,
                     values=["Toutes"] + db.get_attributions_garantie(),
                     width=18, state="readonly")
        self.attr_combo.pack(side="left")
        self.attr_var.trace_add("write", lambda *_: self.refresh())

        mk_btn(sf, "➕ Nouvelle garantie",
               lambda: GarantieDialog(self, self.app, on_save=self.refresh)
               ).pack(side="right")
        mk_btn(sf, "⚙ Statuts",
               lambda: StatutsGarantieDialog(self, self.app,
                                              on_save=self._reload_statuts),
               color=C["btn2"]).pack(side="right", padx=(0, 6))

        cols = ("ems", "constr", "client", "moteur", "attr", "statut", "ouv", "maj")
        col_defs = [
            ("ems",    "N° EMS",          130),
            ("constr", "N° Constructeur", 140),
            ("client", "Client",          140),
            ("moteur", "Moteur",          130),
            ("attr",   "Attribution",     110, "center"),
            ("statut", "Statut",          170),
            ("ouv",    "Ouverture",        90, "center"),
            ("maj",    "Modifié",          95),
        ]
        tf, self.tree = mk_tree(self, cols, col_defs, height=20)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=3)
        mk_btn(af, "📄 Générer fiche", self._fiche).pack(side="left", padx=3)
        mk_btn(af, "📁 Dossier", self._dossier).pack(side="left", padx=3)
        mk_btn(af, "🗑️ Supprimer", self._supprimer,
               color=C["danger"]).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def _reload_statuts(self):
        self.statut_combo["values"] = ["Tous"] + db.get_statuts_garantie()
        self.refresh()

    def refresh(self):
        items = db.get_garanties(statut=self.statut_var.get(),
                                  search=self.search_var.get(),
                                  attribution=self.attr_var.get())
        self._cache = list(items)
        rows = []
        for g in self._cache:
            moteur = g["num_serie"] or ""
            if g["marque"]:
                moteur = f"{g['num_serie']} ({g['marque']})"
            rows.append((
                g["num_ems"], g["num_constructeur"] or "—",
                g["client_nom"] or "", moteur,
                g["attribution"], g["statut"],
                g["date_ouverture"] or "—",
                db.fmt_paris_short(g["updated_at"]),
            ))
        fill_tree(self.tree, rows)

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez une garantie.")
            return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r:
            GarantieDialog(self, self.app, garantie_id=r["id"],
                           on_save=self.refresh)

    def _fiche(self):
        r = self._sel()
        if not r:
            return
        from garantie_generator import sauvegarder_fiche
        g = db.get_garantie(garantie_id=r["id"])
        path = sauvegarder_fiche(g)
        ouvrir_fichier(path)

    def _dossier(self):
        r = self._sel()
        if not r:
            return
        d = Path(__file__).parent / "garanties" / r["num_ems"]
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno(
                "Supprimer",
                f"Supprimer la garantie {r['num_ems']} ?\n\n"
                "(Le dossier sur disque n'est pas supprimé.)"):
            db.delete_garantie(r["id"])
            self.refresh()


class StatutsGarantieDialog(tk.Toplevel):
    """Configuration de la liste des statuts de garantie."""
    def __init__(self, parent, app, on_save=None):
        super().__init__(parent)
        self.app = app
        self.on_save = on_save
        self.title("Statuts de garantie")
        self.geometry("420x480")
        self.configure(bg=C["bg"])
        self.grab_set()

        tk.Label(self, text="Statuts de garantie configurables",
                 bg=C["bg"], font=F["h2"], fg=C["header"]).pack(pady=(14, 4))
        tk.Label(self, text="Ces statuts apparaissent dans le suivi des garanties.",
                 bg=C["bg"], font=F["small"], fg=C["text_muted"]).pack()

        lf = tk.Frame(self, bg=C["bg"])
        lf.pack(fill="both", expand=True, padx=16, pady=10)
        self.lst = tk.Listbox(lf, font=("Segoe UI", 10), height=12,
                              selectbackground=C["nav_sel"],
                              selectforeground="white")
        self.lst.pack(fill="both", expand=True)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=16, pady=(0, 6))
        self.input_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.input_var, width=28).pack(
            side="left", padx=(0, 4))
        mk_btn(af, "➕ Ajouter", self._add).pack(side="left")

        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(pady=10)
        mk_btn(bf, "✏️ Renommer", self._rename).pack(side="left", padx=4)
        mk_btn(bf, "🗑️ Supprimer", self._del,
               color=C["danger"]).pack(side="left", padx=4)
        mk_btn(bf, "Fermer", self._close, color=C["btn3"]).pack(side="left", padx=4)
        self._refresh()

    def _refresh(self):
        self.lst.delete(0, "end")
        for s in db.get_statuts_garantie():
            self.lst.insert("end", s)

    def _add(self):
        v = self.input_var.get().strip()
        if not v:
            return
        if not db.add_statut_garantie(v):
            messagebox.showwarning("Doublon", f"Le statut '{v}' existe déjà.")
            return
        self.input_var.set("")
        self._refresh()

    def _selected(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un statut.")
            return None
        return self.lst.get(sel[0])

    def _rename(self):
        cur = self._selected()
        if not cur:
            return
        new = simpledialog.askstring("Renommer", f"Nouveau nom pour '{cur}' :",
                                      parent=self, initialvalue=cur)
        if new and new.strip() and new != cur:
            if db.update_statut_garantie(cur, new.strip()):
                self._refresh()
            else:
                messagebox.showwarning("Erreur", "Ce nom existe déjà.")

    def _del(self):
        cur = self._selected()
        if not cur:
            return
        if messagebox.askyesno(
                "Supprimer",
                f"Supprimer le statut '{cur}' ?\n\n"
                "(Les garanties existantes gardent leur libellé.)"):
            db.delete_statut_garantie(cur)
            self._refresh()

    def _close(self):
        if self.on_save:
            self.on_save()
        self.destroy()


class GarantieDialog(tk.Toplevel):
    def __init__(self, parent, app, garantie_id=None, on_save=None):
        super().__init__(parent)
        self.app = app
        self.garantie_id = garantie_id
        self.on_save = on_save
        self.is_edit = garantie_id is not None
        self.title("Modifier la garantie" if self.is_edit
                   else "Nouvelle garantie moteur")
        self.geometry("860x840")
        self.minsize(860, 620)
        self.configure(bg=C["bg"])
        self.grab_set()

        self._clients = list(db.get_clients())
        self._moteurs = list(db.get_moteurs())
        self._techniciens = list(db.get_techniciens())
        self._contacts = list(db.get_contacts())
        self._moteur_by_ns = {m["num_serie"]: m for m in self._moteurs}
        self._client_by_id = {c["id"]: c for c in self._clients}
        self._tech_by_nom = {t["nom"]: t for t in self._techniciens}
        self._contact_by_nom = {c["nom"]: c for c in self._contacts}

        mk_header(self, "Garantie moteur",
                  "   Dossier de garantie constructeur / interne")

        # Boutons packés en bas EN PREMIER pour rester visibles au redimensionnement
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(side="bottom", pady=12)
        mk_btn(bf, "💾 Enregistrer", lambda: self._save(False)).pack(
            side="left", padx=6)
        mk_btn(bf, "📄 Enregistrer + Fiche", lambda: self._save(True)).pack(
            side="left", padx=6)
        mk_btn(bf, "📧 Prévenir client",
               self._mail_client_garantie, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "📧 Prévenir responsable",
               self._mail_tech_garantie, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "Annuler", self.destroy, color=C["btn3"]).pack(
            side="left", padx=6)

        _canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        _vsb = ttk.Scrollbar(self, orient="vertical", command=_canvas.yview)
        _canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        _canvas.pack(fill="both", expand=True)
        body = tk.Frame(_canvas, bg=C["bg"])
        _win = _canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>",
                  lambda e: _canvas.configure(scrollregion=_canvas.bbox("all")))
        _canvas.bind("<Configure>",
                     lambda e: _canvas.itemconfig(_win, width=e.width))

        def _on_wheel(ev):
            try: _canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            except tk.TclError: pass
        self.bind("<MouseWheel>", _on_wheel)
        for _cls in ("TFrame", "Frame", "TLabel", "Label",
                     "TCheckbutton", "Checkbutton", "TButton", "Button"):
            self.bind_class(_cls, "<MouseWheel>", _on_wheel, add="+")

        # padding interne du contenu
        body = tk.Frame(body, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=24, pady=12)

        # Moteur (N° série) — détermine le client automatiquement
        tk.Label(body, text="Moteur (N° série) *", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.moteur_var = tk.StringVar()
        self.moteur_combo = SearchableCombobox(
            body, textvariable=self.moteur_var,
            values=[m["num_serie"] for m in self._moteurs], width=44)
        self.moteur_combo.pack(fill="x", pady=(2, 2))
        self.moteur_combo.bind("<<ComboboxSelected>>", self._on_moteur)
        self.moteur_info = tk.Label(body, text="", bg=C["bg"],
                                     fg=C["text_muted"], font=F["small"],
                                     anchor="w")
        self.moteur_info.pack(anchor="w", pady=(0, 10))

        # N° constructeur + N° EMS (lecture seule si edit)
        row = tk.Frame(body, bg=C["bg"])
        row.pack(fill="x", pady=(0, 10))
        nc = tk.Frame(row, bg=C["bg"])
        nc.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Label(nc, text="N° garantie constructeur", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.num_constr_var = tk.StringVar()
        ttk.Entry(nc, textvariable=self.num_constr_var, width=28).pack(
            fill="x", pady=(2, 0))
        ne = tk.Frame(row, bg=C["bg"])
        ne.pack(side="left")
        tk.Label(ne, text="N° garantie EMS", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.num_ems_lbl = tk.Label(
            ne, text="(auto à la création)", bg="#eef2f7",
            font=F["body_bold"], fg=C["header"], anchor="w",
            relief="groove", width=20, padx=6)
        self.num_ems_lbl.pack(pady=(2, 0), ipady=3)

        # Attribution + Statut
        row2 = tk.Frame(body, bg=C["bg"])
        row2.pack(fill="x", pady=(0, 10))
        at = tk.Frame(row2, bg=C["bg"])
        at.pack(side="left", padx=(0, 16))
        tk.Label(at, text="Attribution", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.attr_var = tk.StringVar(value=db.GARANTIE_ATTRIBUTION_DEFAULT)
        self.attr_combo = ttk.Combobox(
            at, textvariable=self.attr_var,
            values=db.get_attributions_garantie(), width=18)
        self.attr_combo.pack(pady=(2, 0))
        st = tk.Frame(row2, bg=C["bg"])
        st.pack(side="left", fill="x", expand=True)
        tk.Label(st, text="Statut", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.statut_var = tk.StringVar(value=db.GARANTIE_STATUT_DEFAULT)
        ttk.Combobox(st, textvariable=self.statut_var,
                     values=db.get_statuts_garantie(), width=30,
                     state="readonly").pack(fill="x", pady=(2, 0))

        # Responsable garantie
        tk.Label(body, text="Responsable garantie", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.responsable_var = tk.StringVar()
        self.responsable_combo = SearchableCombobox(
            body, textvariable=self.responsable_var,
            values=[t["nom"] for t in self._techniciens], width=44)
        self.responsable_combo.pack(fill="x", pady=(2, 10))

        # Demandeur
        tk.Label(body, text="Demandeur", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        _contact_noms = [c["nom"] for c in self._contacts]
        self.demandeur_var = tk.StringVar()
        self.demandeur_combo = SearchableCombobox(
            body, textvariable=self.demandeur_var,
            values=_contact_noms, width=44, allow_free_text=True)
        self.demandeur_combo.pack(fill="x", pady=(2, 4))
        self.demandeur_combo.bind("<<ComboboxSelected>>", self._on_demandeur_selected)
        dem_row = tk.Frame(body, bg=C["bg"])
        dem_row.pack(fill="x", pady=(0, 10))
        self.email_demand_var = tk.StringVar()
        self.tel_demand_var = tk.StringVar()
        for lbl, var in [("Email", self.email_demand_var), ("Tél.", self.tel_demand_var)]:
            f = tk.Frame(dem_row, bg=C["bg"])
            f.pack(side="left", padx=(0, 16))
            tk.Label(f, text=lbl, bg=C["bg"], font=F["body"]).pack(anchor="w")
            ttk.Entry(f, textvariable=var, width=24).pack()

        # Dates + montant
        row3 = tk.Frame(body, bg=C["bg"])
        row3.pack(fill="x", pady=(0, 10))
        do = tk.Frame(row3, bg=C["bg"])
        do.pack(side="left", padx=(0, 12))
        tk.Label(do, text="Date ouverture", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.d_ouv_var = tk.StringVar()
        DateEntry(do, textvariable=self.d_ouv_var, width=14).pack(pady=(2, 0))
        dc = tk.Frame(row3, bg=C["bg"])
        dc.pack(side="left", padx=(0, 12))
        tk.Label(dc, text="Date clôture", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.d_clo_var = tk.StringVar()
        DateEntry(dc, textvariable=self.d_clo_var, width=14).pack(pady=(2, 0))
        mt = tk.Frame(row3, bg=C["bg"])
        mt.pack(side="left")
        tk.Label(mt, text="Montant (€)", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.montant_var = tk.StringVar()
        ttk.Entry(mt, textvariable=self.montant_var, width=12).pack(pady=(2, 0))

        # Description
        tk.Label(body, text="Description du dossier *", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.txt_desc = tk.Text(body, height=5, font=("Segoe UI", 10),
                                 wrap="word", relief="solid", bd=1,
                                 padx=6, pady=4)
        self.txt_desc.pack(fill="x", pady=(2, 10))


        # Intervention liée (optionnel) — recherche multi-champs
        tk.Label(body, text="Intervention liée (optionnel)", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self._interventions = list(db.get_interventions(statut="Tous"))
        # Construire les libelles "BON — moteur — type — client"
        self._inv_labels = {}     # label -> intervention dict
        labels = []
        for inv in self._interventions:
            lab = (f"{inv.get('num_bon','')} — "
                   f"{inv.get('num_serie','') or inv.get('navire','')} — "
                   f"{inv.get('type_intervention','')} — "
                   f"{inv.get('client_nom','')}").strip(" —")
            self._inv_labels[lab] = inv
            labels.append(lab)
        self.inv_lie_var = tk.StringVar()
        self.inv_lie_combo = SearchableCombobox(
            body, textvariable=self.inv_lie_var,
            values=labels, width=44)
        self.inv_lie_combo.pack(fill="x", pady=(2, 4))
        self.inv_lie_combo.bind("<<ComboboxSelected>>", self._on_inv_lie_selected)
        tk.Label(body, text="💡 Tapez un n° de bon, un n° de série, un type "
                            "ou un client pour filtrer",
                 bg=C["bg"], fg=C["text_muted"],
                 font=("Segoe UI", 8, "italic")).pack(anchor="w", pady=(0, 10))

        # Commentaires
        tk.Label(body, text="Commentaires / Suivi", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.txt_comm = tk.Text(body, height=4, font=("Segoe UI", 10),
                                 wrap="word", relief="solid", bd=1,
                                 padx=6, pady=4)
        self.txt_comm.pack(fill="x", pady=(2, 4))

        if self.is_edit:
            self.after(50, self._load)

    def _on_moteur(self, _ev=None):
        m = self._moteur_by_ns.get(self.moteur_var.get())
        if m:
            c = self._client_by_id.get(m["client_id"])
            cl = c["nom"] if c else "—"
            mk = f"{m['marque']} {m['ref_constructeur']}".strip() \
                if m["marque"] else (m["ref_constructeur"] or m["type_moteur"])
            self.moteur_info.config(
                text=f"→ Client : {cl}   |   {mk}   |   Navire : {m['navire'] or '—'}")
            # Pré-remplir l'attribution avec la marque du moteur si l'utilisateur
            # n'a pas encore choisi (une garantie BAUDOUIN concerne un moteur
            # BAUDOUIN). Ne pas écraser une saisie existante en mode édition.
            marque = (m["marque"] or "").strip()
            if marque and self.attr_var.get() in (
                    "", db.GARANTIE_ATTRIBUTION_DEFAULT):
                self.attr_var.set(marque)

    def _on_demandeur_selected(self, *_):
        c = self._contact_by_nom.get(self.demandeur_var.get())
        if c:
            if not self.email_demand_var.get():
                self.email_demand_var.set(c.get("email", ""))
            if not self.tel_demand_var.get():
                self.tel_demand_var.set(c.get("telephone", ""))

    def _on_inv_lie_selected(self, *_):
        inv = self._inv_labels.get(self.inv_lie_var.get())
        if not inv:
            return
        if not self.demandeur_var.get():
            nom = row_get(inv, "nom_demandeur")
            if nom:
                self.demandeur_combo.set(nom)
                if not self.email_demand_var.get():
                    self.email_demand_var.set(row_get(inv, "email_demandeur"))
                if not self.tel_demand_var.get():
                    self.tel_demand_var.set(row_get(inv, "telephone_demandeur"))

    def _load(self):
        g = db.get_garantie(garantie_id=self.garantie_id)
        if not g:
            return
        self.num_ems_lbl.config(text=row_get(g, "num_ems"))
        self.num_constr_var.set(row_get(g, "num_constructeur"))
        m = next((x for x in self._moteurs
                  if x["id"] == row_get(g, "moteur_id")), None)
        if m:
            self.moteur_combo.set(m["num_serie"])
            self._on_moteur()
        self.attr_var.set(row_get(g, "attribution",
                                   db.GARANTIE_ATTRIBUTION_DEFAULT))
        self.statut_var.set(row_get(g, "statut", db.GARANTIE_STATUT_DEFAULT))
        self.responsable_combo.set(row_get(g, "responsable"))
        # Demandeur : depuis la garantie, sinon depuis l'intervention liée
        nom_dem = row_get(g, "nom_demandeur")
        if not nom_dem:
            inv_id_dem = row_get(g, "intervention_id")
            if inv_id_dem:
                inv = db.get_intervention(inv_id=inv_id_dem)
                if inv:
                    nom_dem = row_get(inv, "nom_demandeur")
                    if not self.email_demand_var.get():
                        self.email_demand_var.set(row_get(inv, "email_demandeur"))
                    if not self.tel_demand_var.get():
                        self.tel_demand_var.set(row_get(inv, "telephone_demandeur"))
        self.demandeur_combo.set(nom_dem)
        self.email_demand_var.set(row_get(g, "email_demandeur") or self.email_demand_var.get())
        self.tel_demand_var.set(row_get(g, "telephone_demandeur") or self.tel_demand_var.get())
        self.d_ouv_var.set(row_get(g, "date_ouverture"))
        self.d_clo_var.set(row_get(g, "date_cloture"))
        self.montant_var.set(row_get(g, "montant"))
        # Intervention liee
        inv_id = row_get(g, "intervention_id")
        if inv_id:
            inv = next((i for i in self._interventions if i["id"] == inv_id), None)
            if inv:
                for lab, i in self._inv_labels.items():
                    if i["id"] == inv_id:
                        self.inv_lie_combo.set(lab)
                        break
        self.txt_desc.delete("1.0", "end")
        self.txt_desc.insert("1.0", row_get(g, "description"))
        self.txt_comm.delete("1.0", "end")
        self.txt_comm.insert("1.0", row_get(g, "commentaires"))

    def _mail_client_garantie(self):
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le dossier avant de notifier le client.")
            return
        g = db.get_garantie(garantie_id=self.garantie_id)
        client = db.get_client(g["client_id"]) if g.get("client_id") else None
        moteur = db.get_moteur(g["moteur_id"]) if g.get("moteur_id") else None
        email = (row_get(client, "email") if client else "")
        if not email:
            messagebox.showwarning("Email manquant",
                "Aucun email renseigné pour ce client.")
            return
        num_ems = row_get(g, "num_ems")
        fiche_path = ""
        _d = Path(_dossiers_root_garanties()) / num_ems
        for ext in (".pdf", ".html"):
            p = _d / f"fiche_garantie{ext}"
            if p.exists():
                fiche_path = str(p)
                break
        _, pj_auto = mailer.email_garantie_client(g, client or {}, moteur or {},
                                                   fiche_path)
        if not pj_auto and fiche_path:
            messagebox.showinfo("Pièce jointe",
                f"Joignez la fiche manuellement depuis :\n{_d}")
        db.mark_notifie_garantie(self.garantie_id, "client")

    def _mail_tech_garantie(self):
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le dossier avant de notifier le responsable.")
            return
        g = db.get_garantie(garantie_id=self.garantie_id)
        responsable = row_get(g, "responsable")
        if not responsable:
            messagebox.showwarning("Responsable manquant",
                "Aucun responsable désigné pour ce dossier.")
            return
        t = db.get_technicien_by_nom(responsable)
        tech_email = row_get(t, "email") if t else ""
        if not tech_email:
            messagebox.showwarning("Email manquant",
                f"Aucun email renseigné pour {responsable}.")
            return
        client = db.get_client(g["client_id"]) if g.get("client_id") else None
        moteur = db.get_moteur(g["moteur_id"]) if g.get("moteur_id") else None
        num_ems = row_get(g, "num_ems")
        fiche_path = ""
        _d = Path(_dossiers_root_garanties()) / num_ems
        for ext in (".pdf", ".html"):
            p = _d / f"fiche_garantie{ext}"
            if p.exists():
                fiche_path = str(p)
                break
        _, pj_auto = mailer.email_garantie_technicien(g, client or {}, moteur or {},
                                                       tech_email, fiche_path)
        if not pj_auto and fiche_path:
            messagebox.showinfo("Pièce jointe",
                f"Joignez la fiche manuellement depuis :\n{_d}")
        db.mark_notifie_garantie(self.garantie_id, "tech")

    def _get_intervention_id_lie(self):
        """Retourne l'id de l'intervention liee selectionnee, ou '' si aucune."""
        lab = self.inv_lie_var.get().strip()
        if not lab:
            return ""
        inv = self._inv_labels.get(lab)
        return inv["id"] if inv else ""

    def _save(self, generer=False):
        ns = self.moteur_var.get().strip()
        desc = self.txt_desc.get("1.0", "end").strip()
        if not ns or not desc:
            messagebox.showwarning(
                "Champs manquants",
                "Le moteur (N° série) et la description sont obligatoires.")
            return
        m = self._moteur_by_ns.get(ns)
        if not m:
            messagebox.showwarning(
                "Moteur inconnu",
                f"'{ns}' ne correspond à aucun moteur enregistré.")
            return
        for lbl, v in [("ouverture", self.d_ouv_var.get().strip()),
                       ("clôture", self.d_clo_var.get().strip())]:
            if v and not DateEntry.is_valid(v):
                messagebox.showwarning(
                    "Date invalide",
                    f"Date de {lbl} invalide : '{v}'\nFormat : JJ/MM/AAAA")
                return
        data = {
            "num_constructeur": self.num_constr_var.get().strip(),
            "moteur_id": m["id"],
            "client_id": m["client_id"],
            "attribution": self.attr_var.get(),
            "statut": self.statut_var.get(),
            "responsable": self.responsable_var.get().strip(),
            "nom_demandeur": self.demandeur_var.get().strip(),
            "email_demandeur": self.email_demand_var.get().strip(),
            "telephone_demandeur": self.tel_demand_var.get().strip(),
            "date_ouverture": self.d_ouv_var.get().strip(),
            "date_cloture": self.d_clo_var.get().strip(),
            "montant": self.montant_var.get().strip(),
            "description": desc,
            "commentaires": self.txt_comm.get("1.0", "end").strip(),
            "intervention_id": self._get_intervention_id_lie(),
        }
        
        if self.is_edit:
            db.update_garantie(self.garantie_id, data)
            num = db.get_garantie(garantie_id=self.garantie_id)["num_ems"]
        else:
            gid, num = db.create_garantie(data)
            self.garantie_id = gid
            self.is_edit = True

        if self.on_save:
            self.on_save()

        if generer:
            from garantie_generator import sauvegarder_fiche
            g = db.get_garantie(garantie_id=self.garantie_id)
            path = sauvegarder_fiche(g)
            messagebox.showinfo("Fiche générée",
                                f"✅ {num}\nFiche enregistrée :\n{path}")
            ouvrir_fichier(path)
            self.destroy()
        else:
            messagebox.showinfo("Enregistré",
                                f"✅ Garantie {num} enregistrée.")
            self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# AMÉLIORATION CONTINUE
# ══════════════════════════════════════════════════════════════════════════════
class AmeliorationsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        mk_header(self, "Amélioration continue",
                  "   Tickets de demande d'amélioration clients")

        sf = tk.Frame(self, bg=C["bg"])
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="🔍 Recherche :", bg=C["bg"], font=("Arial", 10)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(sf, textvariable=self.search_var, width=26).pack(side="left", padx=4)

        tk.Label(sf, text="  Statut :", bg=C["bg"], font=("Arial", 10)).pack(side="left", padx=(12, 4))
        self.statut_var = tk.StringVar(value="Tous")
        ttk.Combobox(sf, textvariable=self.statut_var,
                     values=["Tous"] + db.AMELIO_STATUTS, width=12,
                     state="readonly").pack(side="left")
        self.statut_var.trace_add("write", lambda *_: self.refresh())

        tk.Label(sf, text="  Priorité :", bg=C["bg"], font=("Arial", 10)).pack(side="left", padx=(12, 4))
        self.prio_var = tk.StringVar(value="Toutes")
        ttk.Combobox(sf, textvariable=self.prio_var,
                     values=["Toutes"] + db.AMELIO_PRIORITES, width=12,
                     state="readonly").pack(side="left")
        self.prio_var.trace_add("write", lambda *_: self.refresh())

        mk_btn(sf, "➕ Nouveau sujet",
               lambda: AmeliorationDialog(self, self.app, on_save=self.refresh)
               ).pack(side="right")

        cols = ("num", "titre", "client", "prio", "statut", "maj")
        col_defs = [
            ("num",    "N° Ticket",   130),
            ("titre",  "Sujet",       320),
            ("client", "Client",      180),
            ("prio",   "Priorité",    100, "center"),
            ("statut", "Statut",      110, "center"),
            ("maj",    "Modifié",     120),
        ]
        tf, self.tree = mk_tree(self, cols, col_defs, height=20)
        tf.pack(fill="both", expand=True, padx=20, pady=4)
        # Tags couleur priorité
        self.tree.tag_configure("prio_crit", background="#fef2f2",
                                 foreground=C["urg_critique"])
        self.tree.tag_configure("prio_haute", background="#fff7ed",
                                 foreground=C["urg_urgente"])

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=3)
        mk_btn(af, "📄 Générer fiche", self._fiche).pack(side="left", padx=3)
        mk_btn(af, "📁 Dossier", self._dossier).pack(side="left", padx=3)
        mk_btn(af, "🗑️ Supprimer", self._supprimer,
               color=C["danger"]).pack(side="left", padx=3)
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self._cache = []

    def refresh(self):
        items = db.get_ameliorations(statut=self.statut_var.get(),
                                      search=self.search_var.get(),
                                      priorite=self.prio_var.get())
        self._cache = list(items)
        self.tree.delete(*self.tree.get_children())
        for i, a in enumerate(self._cache):
            tags = ["even" if i % 2 == 0 else "odd"]
            if a["priorite"] == "Critique":
                tags.append("prio_crit")
            elif a["priorite"] == "Haute":
                tags.append("prio_haute")
            self.tree.insert("", "end", iid=str(i), values=(
                a["num_ticket"], a["titre"], a["client_nom"] or "",
                a["priorite"], a["statut"],
                db.fmt_paris_short(a["updated_at"]),
            ), tags=tuple(tags))

    def _sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un sujet.")
            return None
        return self._cache[int(sel[0])]

    def _modifier(self):
        r = self._sel()
        if r:
            AmeliorationDialog(self, self.app, amelio_id=r["id"],
                               on_save=self.refresh)

    def _fiche(self):
        r = self._sel()
        if not r:
            return
        from amelioration_generator import sauvegarder_fiche
        a = db.get_amelioration(amelio_id=r["id"])
        path = sauvegarder_fiche(a)
        ouvrir_fichier(path)

    def _dossier(self):
        r = self._sel()
        if not r:
            return
        d = Path(__file__).parent / "ameliorations" / r["num_ticket"]
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    def _supprimer(self):
        r = self._sel()
        if r and messagebox.askyesno(
                "Supprimer",
                f"Supprimer le ticket {r['num_ticket']} ?\n\n"
                "(Le dossier sur disque n'est pas supprimé.)"):
            db.delete_amelioration(r["id"])
            self.refresh()


class AmeliorationDialog(tk.Toplevel):
    def __init__(self, parent, app, amelio_id=None, on_save=None):
        super().__init__(parent)
        self.app = app
        self.amelio_id = amelio_id
        self.on_save = on_save
        self.is_edit = amelio_id is not None
        self.title("Modifier le sujet" if self.is_edit
                   else "Nouveau sujet d'amélioration")
        self.geometry("680x720")
        self.configure(bg=C["bg"])
        self.grab_set()

        self._clients = list(db.get_clients())
        self._techs = list(db.get_techniciens())

        mk_header(self, "Sujet d'amélioration",
                  "   Demande client / piste d'amélioration")

        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=24, pady=14)

        # Titre
        tk.Label(body, text="Titre / Sujet *", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.titre_var = tk.StringVar()
        ttk.Entry(body, textvariable=self.titre_var, width=64).pack(
            fill="x", pady=(2, 10))

        # Client + Priorité (ligne)
        row = tk.Frame(body, bg=C["bg"])
        row.pack(fill="x", pady=(0, 10))
        cl = tk.Frame(row, bg=C["bg"])
        cl.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Label(cl, text="Client demandeur", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.client_var = tk.StringVar()
        self.client_combo = SearchableCombobox(
            cl, textvariable=self.client_var,
            values=[c["nom"] for c in self._clients], width=30)
        self.client_combo.pack(fill="x", pady=(2, 0))
        pr = tk.Frame(row, bg=C["bg"])
        pr.pack(side="left")
        tk.Label(pr, text="Priorité", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.prio_var = tk.StringVar(value=db.AMELIO_PRIORITE_DEFAULT)
        ttk.Combobox(pr, textvariable=self.prio_var,
                     values=db.AMELIO_PRIORITES, width=14,
                     state="readonly").pack(pady=(2, 0))

        # Statut
        row2 = tk.Frame(body, bg=C["bg"])
        row2.pack(fill="x", pady=(0, 10))
        st = tk.Frame(row2, bg=C["bg"])
        st.pack(side="left")
        tk.Label(st, text="Statut", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.statut_var = tk.StringVar(value=db.AMELIO_STATUT_DEFAULT)
        ttk.Combobox(st, textvariable=self.statut_var,
                     values=db.AMELIO_STATUTS, width=16,
                     state="readonly").pack(pady=(2, 0))

        # Description
        tk.Label(body, text="Description de la demande *", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.txt_desc = tk.Text(body, height=6, font=("Segoe UI", 10),
                                 wrap="word", relief="solid", bd=1,
                                 padx=6, pady=4)
        self.txt_desc.pack(fill="x", pady=(2, 10))

        # Commentaires
        tk.Label(body, text="Commentaires / Suivi", bg=C["bg"],
                 font=F["body_bold"], anchor="w").pack(anchor="w")
        self.txt_comm = tk.Text(body, height=5, font=("Segoe UI", 10),
                                 wrap="word", relief="solid", bd=1,
                                 padx=6, pady=4)
        self.txt_comm.pack(fill="x", pady=(2, 4))

        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(side="bottom", pady=12)
        mk_btn(bf, "💾 Enregistrer", lambda: self._save(False)).pack(
            side="left", padx=6)
        mk_btn(bf, "📄 Enregistrer + Fiche", lambda: self._save(True)).pack(
            side="left", padx=6)
        mk_btn(bf, "Annuler", self.destroy, color=C["btn3"]).pack(
            side="left", padx=6)

        if self.is_edit:
            self.after(50, self._load)

    def _load(self):
        a = db.get_amelioration(amelio_id=self.amelio_id)
        if not a:
            return
        self.titre_var.set(row_get(a, "titre"))
        c = next((x for x in self._clients
                  if x["id"] == row_get(a, "client_id")), None)
        if c:
            self.client_combo.set(c["nom"])
        self.prio_var.set(row_get(a, "priorite", db.AMELIO_PRIORITE_DEFAULT))
        self.statut_var.set(row_get(a, "statut", db.AMELIO_STATUT_DEFAULT))
        self.txt_desc.delete("1.0", "end")
        self.txt_desc.insert("1.0", row_get(a, "description"))
        self.txt_comm.delete("1.0", "end")
        self.txt_comm.insert("1.0", row_get(a, "commentaires"))

    def _save(self, generer=False):
        titre = self.titre_var.get().strip()
        desc = self.txt_desc.get("1.0", "end").strip()
        if not titre or not desc:
            messagebox.showwarning(
                "Champs manquants",
                "Le titre et la description sont obligatoires.")
            return
        client = next((c for c in self._clients
                       if c["nom"] == self.client_var.get()), None)
        data = {
            "titre": titre,
            "client_id": client["id"] if client else "",
            "description": desc,
            "priorite": self.prio_var.get(),
            "statut": self.statut_var.get(),
            "commentaires": self.txt_comm.get("1.0", "end").strip(),
        }
        if self.is_edit:
            db.update_amelioration(self.amelio_id, data)
            num = db.get_amelioration(amelio_id=self.amelio_id)["num_ticket"]
        else:
            aid, num = db.create_amelioration(data)
            self.amelio_id = aid
            self.is_edit = True

        if self.on_save:
            self.on_save()

        if generer:
            from amelioration_generator import sauvegarder_fiche
            a = db.get_amelioration(amelio_id=self.amelio_id)
            path = sauvegarder_fiche(a)
            messagebox.showinfo(
                "Fiche générée",
                f"✅ {num}\nFiche enregistrée dans :\n{path}")
            ouvrir_fichier(path)
            self.destroy()
        else:
            messagebox.showinfo("Enregistré",
                                f"✅ Sujet {num} enregistré.")
            self.destroy()


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
                                 on_save=lambda: self.app.goto("dashboard"))
               ).pack(pady=20, ipadx=20, ipady=8)

    def refresh(self): pass


# ══════════════════════════════════════════════════════════════════════════════
# IMPORT / EXPORT — transfert de bons hors connexion
# ══════════════════════════════════════════════════════════════════════════════
class ImportExportFrame(tk.Frame):
    """Onglet Import/Export : exporter un bon en .ems, l'importer ou le modifier hors-ligne."""

    def __init__(self, parent, app):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        self._bundle      = None   # bundle chargé depuis un fichier .ems
        self._bundle_path = None
        self._exp_cache   = []     # liste des bons affichés dans le panneau export

        mk_header(self, "Import / Export", "   Transfert de bons hors connexion")

        # ── Bandeau statut connexion ─────────────────────────────────────────
        self._st_frame = tk.Frame(self, bg="#1a3a1a")
        self._st_frame.pack(fill="x", padx=20, pady=(0, 6))
        self._st_lbl = tk.Label(self._st_frame, text="Vérification…",
                                 font=("Arial", 9), bg="#1a3a1a", fg="#a0f0a0", anchor="w")
        self._st_lbl.pack(side="left", padx=10, pady=4)
        tk.Button(self._st_frame, text="↻ Actualiser", command=self._check_connexion,
                  font=("Arial", 8), relief="flat",
                  bg="#1a3a1a", fg="#88cc88", bd=0).pack(side="right", padx=8)

        # ── Deux panneaux côte à côte ────────────────────────────────────────
        panels = tk.Frame(self, bg=C["bg"])
        panels.pack(fill="both", expand=True, padx=20, pady=4)
        panels.columnconfigure(0, weight=1)
        panels.columnconfigure(1, weight=1)
        panels.rowconfigure(0, weight=1)

        self._build_import_panel(panels)
        self._build_export_panel(panels)

        self.after(200, self._check_connexion)

    # ── Connexion ─────────────────────────────────────────────────────────────
    def _check_connexion(self):
        """Vérifie la connexion en thread de fond pour ne pas bloquer l'UI."""
        import threading
        def _bg():
            en_ligne = db.ping(timeout=2.0)
            if en_ligne:
                bg, fg, txt = "#1a3a1a", "#a0f0a0", "● Serveur en ligne — connexion active"
            else:
                bg, fg, txt = "#3a1a1a", "#f0a0a0", "○ Serveur hors ligne — mode import/export uniquement"
            try:
                self.after(0, lambda: (
                    self._st_frame.config(bg=bg) or
                    self._st_lbl.config(bg=bg, fg=fg, text=txt)
                ))
            except Exception:
                pass
        threading.Thread(target=_bg, daemon=True).start()
        self.after(15000, self._check_connexion)

    # ── Panneau EXPORTER ──────────────────────────────────────────────────────
    def _build_export_panel(self, parent):
        lf = tk.LabelFrame(parent, text="  📤  Exporter un bon  ",
                            font=("Arial", 10, "bold"),
                            bg=C["bg"], fg=C["header"], bd=1, relief="groove",
                            padx=8, pady=8)
        lf.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=4)
        lf.rowconfigure(1, weight=1)
        lf.columnconfigure(0, weight=1)

        tk.Label(lf, text="Sélectionnez un bon à exporter vers un fichier .ems :",
                 bg=C["bg"], font=("Arial", 9)).grid(row=0, column=0, sticky="w", pady=(0, 4))

        # Barre de recherche
        sf = tk.Frame(lf, bg=C["bg"])
        sf.grid(row=0, column=0, sticky="ew")
        tk.Label(sf, text="🔍", bg=C["bg"]).pack(side="left")
        self._exp_search = tk.StringVar()
        self._exp_search.trace_add("write", lambda *_: self._refresh_export_list())
        ttk.Entry(sf, textvariable=self._exp_search, width=22).pack(side="left", padx=4)

        # Treeview liste des bons
        col_defs = [("num_bon", "N° Bon", 105), ("date", "Date", 80),
                    ("client", "Client", 145), ("statut", "Statut", 78)]
        tf, self._exp_tree = mk_tree(lf, ("num_bon","date","client","statut"), col_defs, height=14)
        tf.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

        mk_btn(lf, "📤 Exporter le bon sélectionné (.ems)", self._exporter
               ).grid(row=2, column=0, sticky="ew", pady=(8, 0))

    # ── Panneau IMPORTER ──────────────────────────────────────────────────────
    def _build_import_panel(self, parent):
        lf = tk.LabelFrame(parent, text="  📥  Importer / Modifier hors-ligne  ",
                            font=("Arial", 10, "bold"),
                            bg=C["bg"], fg=C["header"], bd=1, relief="groove",
                            padx=8, pady=8)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=4)
        lf.columnconfigure(0, weight=1)

        # Sélection du fichier
        ff = tk.Frame(lf, bg=C["bg"])
        ff.pack(fill="x", pady=(0, 8))
        mk_btn(ff, "📁 Ouvrir un fichier .ems…", self._choisir_fichier,
               color=C["btn2"]).pack(side="right", padx=(8, 0))
        self._file_var = tk.StringVar(value="Aucun fichier sélectionné")
        tk.Label(ff, textvariable=self._file_var, bg=C["bg"], font=("Arial", 9),
                 fg=C["header"], wraplength=260, justify="left").pack(side="left", fill="x", expand=True)

        # Aperçu
        pf = tk.LabelFrame(lf, text=" Aperçu ", font=("Arial", 9, "bold"),
                            bg=C["bg"], fg=C["header"], bd=1, relief="groove",
                            padx=6, pady=6)
        pf.pack(fill="x", pady=(0, 10))
        self._prev_vars = {}
        for lbl, key in [("N° Bon", "num_bon"), ("Date", "date_creation"),
                          ("Client", "client_nom"), ("Statut", "statut"),
                          ("Technicien(s)", "technicien"),
                          ("Modif. hors-ligne", "_offline")]:
            r = tk.Frame(pf, bg=C["bg"])
            r.pack(fill="x", pady=1)
            tk.Label(r, text=lbl + " :", bg=C["bg"], font=("Arial", 9),
                     width=17, anchor="e").pack(side="left")
            v = tk.StringVar(value="—")
            self._prev_vars[key] = v
            tk.Label(r, textvariable=v, bg=C["bg"], font=("Arial", 9),
                     anchor="w").pack(side="left", padx=4)

        # Boutons d'action
        self._btn_import = mk_btn(lf, "🔁 Importer en ligne (synchroniser)", self._importer)
        self._btn_import.pack(fill="x", pady=(0, 4))
        self._btn_import.config(state="disabled", fg="white",
                                activeforeground="white", disabledforeground="white")

        self._btn_offline = mk_btn(lf, "✏️ Ouvrir et modifier hors-ligne",
                                    self._modifier_hors_ligne, color=C["btn3"])
        self._btn_offline.pack(fill="x", pady=(0, 4))
        self._btn_offline.config(state="disabled", fg="white",
                                 activeforeground="white", disabledforeground="white")

        tk.Label(lf,
                 text="ℹ️  Flux de travail :\n"
                      "1. Exportez un bon depuis le panneau droit\n"
                      "2. Transférez le fichier .ems sur un PC hors connexion\n"
                      "3. Ouvrez-le ici avec « Modifier hors-ligne »\n"
                      "4. Une fois en ligne, importez pour synchroniser",
                 bg=C["bg"], font=("Arial", 8), fg="#333",
                 justify="left").pack(anchor="w", pady=(8, 0))

    # ── Dossier de travail .ems ───────────────────────────────────────────────
    @staticmethod
    def _ems_folder() -> Path:
        """Retourne (et crée si besoin) le dossier bons_ems/ à la racine du projet."""
        import sys
        base = (Path(sys.executable).parent
                if getattr(sys, "frozen", False)
                else Path(__file__).resolve().parent.parent)
        folder = base / "bons_ems"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    # ── Actualisation de la liste export ──────────────────────────────────────
    def _refresh_export_list(self):
        q = self._exp_search.get()
        try:
            invs = db.get_interventions(search=q)
        except Exception:
            invs = []
        self._exp_cache = list(invs)
        rows = [(r["num_bon"], r.get("date_creation", ""),
                 r.get("client_nom") or "", r.get("statut", ""))
                for r in self._exp_cache]
        fill_tree(self._exp_tree, rows)

    def refresh(self):
        self._refresh_export_list()

    # ── Export ────────────────────────────────────────────────────────────────
    def _sel_export(self):
        sel = self._exp_tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un bon à exporter.")
            return None
        return self._exp_cache[int(sel[0])]

    def _exporter(self):
        r = self._sel_export()
        if not r:
            return
        try:
            inv        = db.get_intervention(inv_id=r["id"])
            clients    = list(db.get_clients())
            moteurs    = list(db.get_moteurs())
            techniciens = list(db.get_techniciens())
            types      = list(db.get_types_intervention())
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger les données :\n{e}")
            return

        from shared import import_export as ie
        bundle   = ie.build_bundle(inv, clients, moteurs, techniciens, types)
        num      = r["num_bon"].replace("/", "-").replace("\\", "-")
        path = filedialog.asksaveasfilename(
            title="Enregistrer le bon exporté",
            initialdir=str(self._ems_folder()),
            initialfile=f"{num}.ems",
            defaultextension=".ems",
            filetypes=[("Bon EMS", "*.ems"), ("JSON", "*.json"), ("Tous", "*.*")],
        )
        if not path:
            return
        try:
            ie.save_bundle(bundle, path)
            messagebox.showinfo("Export réussi",
                f"✅ Bon {r['num_bon']} exporté.\n\n{path}\n\n"
                "Transférez ce fichier sur le PC hors connexion pour compléter\n"
                "l'intervention, puis importez-le ici une fois de retour.")
        except Exception as e:
            messagebox.showerror("Erreur d'export", str(e))

    # ── Choisir fichier à importer ────────────────────────────────────────────
    def _choisir_fichier(self):
        path = filedialog.askopenfilename(
            title="Ouvrir un bon exporté (.ems)",
            initialdir=str(self._ems_folder()),
            filetypes=[("Bon EMS", "*.ems"), ("JSON", "*.json"), ("Tous", "*.*")],
        )
        if not path:
            return
        from shared import import_export as ie
        try:
            bundle = ie.load_bundle(path)
        except Exception as e:
            messagebox.showerror("Fichier invalide", str(e))
            return

        self._bundle      = bundle
        self._bundle_path = path
        self._file_var.set(path)

        inv = bundle["intervention"]
        self._prev_vars["num_bon"].set(inv.get("num_bon", "?"))
        self._prev_vars["date_creation"].set(inv.get("date_creation", "?"))
        self._prev_vars["client_nom"].set(
            inv.get("client_nom") or inv.get("client_id") or "?")
        self._prev_vars["statut"].set(inv.get("statut", "?"))
        self._prev_vars["technicien"].set(inv.get("technicien", "?"))
        self._prev_vars["_offline"].set(
            "✅ Oui" if bundle.get("offline_edits") else "Non")

        self._btn_import.config(state="normal", fg="white", activeforeground="white")
        self._btn_offline.config(state="normal", fg="white", activeforeground="white")

    # ── Importer en ligne ─────────────────────────────────────────────────────
    def _importer(self):
        if not self._bundle:
            return
        orig  = self._bundle.get("intervention", {})
        edits = self._bundle.get("offline_edits") or {}
        data  = {**orig, **edits}
        base_ver = int(orig.get("version", 0))

        def _do_push(force=False):
            return db.push_bons(
                [{"data": data, "base_version": base_ver, "force": force}],
                device="bureau-import")

        try:
            result = _do_push()
        except Exception as e:
            messagebox.showerror("Erreur de synchronisation", str(e))
            return

        appliques = result.get("appliques", 0)
        conflits  = result.get("conflits", [])
        erreurs   = result.get("erreurs", [])

        if erreurs:
            messagebox.showerror("Erreurs d'import", "\n".join(erreurs))
            return

        if conflits:
            c = conflits[0]
            if messagebox.askyesno("Conflit de version",
                    f"Le bon {c['num_bon']} a été modifié sur le serveur\n"
                    f"(version serveur : {c['serveur_version']}, "
                    f"version exportée : {c['base_version']}).\n\n"
                    "Écraser la version serveur avec la version importée ?"):
                try:
                    result2 = _do_push(force=True)
                    messagebox.showinfo("Import réussi (forcé)",
                        f"✅ {result2.get('appliques', 0)} bon(s) synchronisé(s).")
                except Exception as e2:
                    messagebox.showerror("Erreur", str(e2))
            return

        messagebox.showinfo("Import réussi",
            f"✅ {appliques} bon(s) synchronisé(s) avec le serveur.")

        for key in ("interventions", "dashboard"):
            if self.app.frames.get(key):
                self.app.frames[key].refresh()

    # ── Modifier hors-ligne ───────────────────────────────────────────────────
    def _modifier_hors_ligne(self):
        if not self._bundle:
            return
        bundle      = self._bundle
        bundle_path = self._bundle_path

        def _on_offline_save(edits):
            from shared import import_export as ie
            ie.apply_offline_edits(bundle, edits, bundle_path)
            self._prev_vars["_offline"].set("✅ Oui")
            messagebox.showinfo("Enregistré localement",
                f"✅ Modifications enregistrées dans :\n{bundle_path}\n\n"
                "Pour les synchroniser, reconnectez-vous et utilisez\n"
                "« Importer en ligne ».")

        BonDialog(self, self.app,
                  offline_bundle=bundle,
                  offline_path=bundle_path,
                  on_offline_save=_on_offline_save)


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET TABLEAU MATÉRIELS (avec autocomplete pièces détachées, popup flottant)
# ══════════════════════════════════════════════════════════════════════════════
class MaterielsTable(tk.Frame):
    """Tableau editable pour les materiels utilises.
    Le champ 'Reference' a un menu deroulant flottant alimente par le catalogue."""
    COLS = [("qte", "Quantité", 8),
            ("ref", "Référence", 24),
            ("designation", "Désignation", 40)]

    def __init__(self, master, **kw):
        super().__init__(master, bg=C["bg"], **kw)
        self._rows = []
        self._cache = {}              # query -> [pieces]
        self._ref_to_libelle = {}     # ref -> libelle (auto-remplissage)

        # En-tete
        self.head = tk.Frame(self, bg=C["header"])
        self.head.pack(fill="x", pady=(0, 1))
        for key, lbl, w in self.COLS:
            tk.Label(self.head, text=lbl, bg=C["header"], fg="white",
                     font=("Arial", 9, "bold"),
                     width=w, anchor="w", padx=4).pack(side="left", padx=1)
        tk.Label(self.head, text="", bg=C["header"],
                  width=4).pack(side="left", padx=1)

        self.body = tk.Frame(self, bg=C["bg"])
        self.body.pack(fill="x")
        self.bf = tk.Frame(self, bg=C["bg"])
        self.bf.pack(fill="x", pady=(4, 0))
        mk_btn(self.bf, "➕ Ajouter une ligne", self.add_row,
                color=C["btn2"]).pack(side="left")
        tk.Label(self.bf, text="  💡 Cliquez sur ▾ ou tapez 2+ caractères "
                                "pour parcourir le catalogue",
                  bg=C["bg"], fg=C["text_muted"],
                  font=("Arial", 8, "italic")).pack(side="left", padx=(10, 0))
        for _ in range(3):
            self.add_row()

    # ── Construction d'une ligne ────────────────────────────────────────────
    def add_row(self, qte="", ref="", designation=""):
        rf = tk.Frame(self.body, bg=C["bg"])
        rf.pack(fill="x", pady=1)
        v_qte = tk.StringVar(value=qte)
        v_ref = tk.StringVar(value=ref)
        v_des = tk.StringVar(value=designation)

        ttk.Entry(rf, textvariable=v_qte, width=8).pack(side="left", padx=1)

        # Wrapper pour Entry + bouton ▾
        ref_wrap = tk.Frame(rf, bg=C["bg"])
        ref_wrap.pack(side="left", padx=1)
        ref_entry = ttk.Entry(ref_wrap, textvariable=v_ref, width=24)
        ref_entry.pack(side="left")
        ref_btn = tk.Button(ref_wrap, text="▾", font=("Arial", 8),
                             bg="#e8eef5", fg=C["text"],
                             relief="solid", bd=1, padx=4, pady=0,
                             cursor="hand2",
                             activebackground="#d0dceb")
        ref_btn.pack(side="left")

        ttk.Entry(rf, textvariable=v_des, width=40).pack(
            side="left", padx=1, fill="x", expand=True)

        row = {"qte": v_qte, "ref": v_ref, "designation": v_des,
                "frame": rf, "ref_entry": ref_entry, "ref_btn": ref_btn,
                "popup": None, "listbox": None, "after_id": None}

        # Recherche temps reel a la frappe
        v_ref.trace_add("write",
                          lambda *_, _r=row: self._on_ref_type(_r))
        # Bouton fleche : ouvre la liste manuellement
        ref_btn.config(command=lambda _r=row: self._show_all_or_search(_r))
        # Quitter le champ -> fermer la popup (avec petit delai pour permettre click)
        ref_entry.bind("<FocusOut>",
                        lambda e, _r=row: self.after(150,
                                                       lambda: self._close_popup(_r)))
        # Touches de navigation
        ref_entry.bind("<Down>",
                        lambda e, _r=row: self._popup_nav(_r, +1) or "break")
        ref_entry.bind("<Up>",
                        lambda e, _r=row: self._popup_nav(_r, -1) or "break")
        ref_entry.bind("<Return>",
                        lambda e, _r=row: self._popup_select(_r) or "break")
        ref_entry.bind("<Escape>",
                        lambda e, _r=row: self._close_popup(_r) or "break")

        tk.Button(rf, text="✕", bg="#ddd", fg=C["danger"],
                  font=("Arial", 9, "bold"), relief="flat", cursor="hand2",
                  padx=4, pady=0,
                  command=lambda r=row: self._remove(r)).pack(side="left",
                                                                 padx=2)
        self._rows.append(row)

    # ── AUTOCOMPLETE ─────────────────────────────────────────────────────────
    def _on_ref_type(self, row):
        """A la frappe : debounce 200ms puis recherche."""
        # Si on est en train de finaliser une selection, ignorer
        if row.get("selecting"):
            return
        if row.get("after_id"):
            try: self.after_cancel(row["after_id"])
            except tk.TclError: pass
            row["after_id"] = None
        query = row["ref"].get().strip()
        if len(query) < 2:
            self._close_popup(row)
            return
        row["after_id"] = self.after(
            200, lambda: self._do_search(row, query))

    def _show_all_or_search(self, row):
        """Click sur le bouton ▾ : montrer la liste meme si le champ est vide.
        Si vide, montrer les 20 premieres pieces du catalogue."""
        query = row["ref"].get().strip()
        if len(query) >= 2:
            self._do_search(row, query)
        else:
            # Pas de filtre : afficher les 20 premieres pieces
            try:
                pieces = db.get_pieces(search="", limit=20)
            except Exception:
                pieces = []
            self._show_popup(row, pieces)
        # Mettre le focus sur le champ pour pouvoir naviguer
        try: row["ref_entry"].focus_set()
        except tk.TclError: pass

    def _do_search(self, row, query):
        row["after_id"] = None
        if row["ref"].get().strip() != query:
            return
        if query in self._cache:
            pieces = self._cache[query]
        else:
            try:
                pieces = db.get_pieces(search=query, limit=20)
            except Exception:
                pieces = []
            self._cache[query] = pieces
            if len(self._cache) > 200:
                for k in list(self._cache)[:50]:
                    del self._cache[k]
        for p in pieces:
            r = p.get("reference", "")
            if r:
                self._ref_to_libelle[r] = p.get("libelle", "")
        self._show_popup(row, pieces)

    def _show_popup(self, row, pieces):
        """Affiche la liste flottante (Toplevel) sous le champ reference.
        Le popup passe PAR-DESSUS et ne decale pas les autres widgets."""
        self._close_popup(row)
        if not pieces:
            return
        entry = row["ref_entry"]
        try:
            x = entry.winfo_rootx()
            y = entry.winfo_rooty() + entry.winfo_height()
            w = max(480, entry.winfo_width() + 320)
        except tk.TclError:
            return
        # Hauteur = nb lignes * 22px + header 22 + marges, max 280
        h = min(280, 22 * len(pieces) + 26)

        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        try:
            popup.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        popup.wm_geometry(f"{w}x{h}+{x}+{y}")
        popup.configure(bg=C["border"])

        # En-tete
        tk.Label(popup,
                 text=f"  {len(pieces)} pièce(s) trouvée(s)"
                       + ("  (20 max — affinez la recherche)"
                          if len(pieces) >= 20 else ""),
                 bg=C["header"], fg="white", font=("Arial", 8, "bold"),
                 anchor="w", padx=6).pack(fill="x")

        # Listbox + scrollbar
        frame_list = tk.Frame(popup, bg="white")
        frame_list.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        vsb = ttk.Scrollbar(frame_list, orient="vertical")
        listbox = tk.Listbox(frame_list, font=("Consolas", 9),
                              bg="white", fg=C["text"],
                              selectbackground=C["btn"],
                              selectforeground="white",
                              activestyle="none", relief="flat",
                              highlightthickness=0,
                              yscrollcommand=vsb.set)
        vsb.config(command=listbox.yview)
        vsb.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)
        for p in pieces:
            ref = p.get("reference", "")
            lib = p.get("libelle", "")
            listbox.insert("end", f" {ref:<22}  {lib}")
        listbox.selection_set(0)
        listbox.activate(0)
        listbox._pieces = pieces
        row["popup"] = popup
        row["listbox"] = listbox

        listbox.bind("<ButtonRelease-1>",
                      lambda e, _r=row: (self._popup_select(_r), "break")[1])
        listbox.bind("<Return>",
                      lambda e, _r=row: self._popup_select(_r))
        listbox.bind("<Escape>",
                      lambda e, _r=row: self._close_popup(_r))
        # Molette = scroll de la listbox uniquement (pas de propagation)
        def _wheel(ev):
            listbox.yview_scroll(int(-1 * (ev.delta / 120)), "units")
            return "break"
        listbox.bind("<MouseWheel>", _wheel)
        popup.bind("<MouseWheel>", _wheel)

    def _close_popup(self, row):
        if row.get("popup"):
            try: row["popup"].destroy()
            except tk.TclError: pass
        row["popup"] = None
        row["listbox"] = None

    def _popup_nav(self, row, delta):
        """Naviguer (Up/Down) dans la liste meme si le focus est sur l'Entry."""
        lb = row.get("listbox")
        if not lb:
            return None
        try:
            cur = lb.curselection()
            cur_idx = cur[0] if cur else 0
            new_idx = max(0, min(lb.size() - 1, cur_idx + delta))
            lb.selection_clear(0, "end")
            lb.selection_set(new_idx)
            lb.activate(new_idx)
            lb.see(new_idx)
        except tk.TclError:
            pass
        return True   # consomme la touche

    def _popup_select(self, row):
        """Valider la pièce sélectionnée : remplir ref + designation et fermer."""
        lb = row.get("listbox")
        if not lb:
            return None
        sel = lb.curselection()
        if not sel:
            try:
                idx = lb.index("active")
                if idx >= 0:
                    sel = (idx,)
            except tk.TclError:
                pass
        if not sel:
            self._close_popup(row)
            return None
        try:
            piece = lb._pieces[sel[0]]
        except (AttributeError, IndexError):
            self._close_popup(row)
            return None
        ref = piece.get("reference", "")
        lib = piece.get("libelle", "")

        # Pose un flag pour que la trace ignore les modifs qui suivent
        row["selecting"] = True
        # Annuler toute recherche programmee
        if row.get("after_id"):
            try: self.after_cancel(row["after_id"])
            except tk.TclError: pass
            row["after_id"] = None
        row["ref"].set(ref)
        if lib:
            row["designation"].set(lib)
        self._close_popup(row)
        # Liberer le flag apres que toutes les traces se soient ecoulees
        self.after(50, lambda: row.update(selecting=False))

        # Sortir le focus du champ ref (curseur sur designation par exemple)
        try:
            row["ref_entry"].tk_focusNext().focus_set()
        except tk.TclError:
            pass
        return "break"
    
    # ── Methodes existantes ─────────────────────────────────────────────────
    def _remove(self, row):
        if len(self._rows) <= 1:
            row["qte"].set("")
            row["ref"].set("")
            row["designation"].set("")
            return
        self._close_popup(row)
        row["frame"].destroy()
        self._rows.remove(row)

    def load(self, items):
        for r in self._rows:
            self._close_popup(r)
            r["frame"].destroy()
        self._rows.clear()
        if not items:
            for _ in range(3):
                self.add_row()
        else:
            for item in items:
                self.add_row(qte=str(item.get("qte", "")),
                             ref=str(item.get("ref", "")),
                             designation=str(item.get("designation", "")))

    def to_list(self):
        out = []
        for r in self._rows:
            qte = r["qte"].get().strip()
            ref = r["ref"].get().strip()
            desg = r["designation"].get().strip()
            if qte or ref or desg:
                out.append({"qte": qte, "ref": ref, "designation": desg})
        return out


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET TEMPS & FRAIS (anciennement DEPLACEMENTS)
# ══════════════════════════════════════════════════════════════════════════════
import re as _re_depl

def _parse_h(s):
    """Parse '2h30', '1:30', '2.5', '30min' en heures (float). None si inconnu."""
    s = str(s).strip().lower().replace(',', '.')
    if not s:
        return None
    m = _re_depl.match(r'^(\d+):(\d{2})$', s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60
    m = _re_depl.match(r'^(\d+(?:\.\d+)?)h(\d{0,2})$', s)
    if m:
        return float(m.group(1)) + (int(m.group(2)) if m.group(2) else 0) / 60
    m = _re_depl.match(r'^(\d+(?:\.\d+)?)min$', s)
    if m:
        return float(m.group(1)) / 60
    m = _re_depl.match(r'^(\d+(?:\.\d+)?)$', s)
    if m:
        return float(m.group(1))
    return None

def _fmt_h(h):
    hi = int(h)
    mi = round((h - hi) * 60)
    if mi == 60:
        hi += 1; mi = 0
    return f"{hi}h{mi:02d}" if mi else f"{hi}h"

def _parse_km(s):
    s = str(s).strip().lower().replace(',', '.').replace(' ', '')
    if not s:
        return None
    m = _re_depl.match(r'^(\d+(?:\.\d+)?)(?:km)?$', s)
    return float(m.group(1)) if m else None

def _fmt_km(v):
    return f"{int(v)} km" if v == int(v) else f"{v:.1f} km"


class TechnicienFrame(tk.Frame):
    """Champs temps & frais pour un technicien au sein d'une journee."""
    TEXTE_FIELDS = [
        ("trajet_aller_retour", "Km parcourus"),
        ("duree_intervention",  "Duree de l'intervention"),
        ("temps_preparation",   "Temps de preparation"),
        ("temps_rangement",     "Temps de rangement"),
    ]
    CHECK_FIELDS = [
        ("frais_repas",  "Frais de repas"),
        ("frais_hotel",  "Frais d'hotel"),
        ("frais_peages", "Frais de peages"),
    ]

    def __init__(self, master, numero, on_change=None, **kw):
        super().__init__(master, bg=C["bg"], relief="solid", bd=1, **kw)
        self.numero = numero
        self.vars = {}
        self.check_vars = {}

        head = tk.Frame(self, bg="#f5f7fa")
        head.pack(fill="x")
        self._num_lbl = tk.Label(head, text=f"Technicien {numero}",
                                  bg="#f5f7fa", font=("Arial", 8, "bold"),
                                  fg="#4a5560")
        self._num_lbl.pack(side="left", padx=6, pady=2)
        self._del_btn = tk.Button(head, text="✕", bg="#f5f7fa",
                                   fg="#c62828", bd=0,
                                   font=("Arial", 8, "bold"), cursor="hand2",
                                   activebackground="#f5f7fa")
        self._del_btn.pack(side="right", padx=4, pady=1)

        inner = tk.Frame(self, bg=C["bg"])
        inner.pack(fill="x", padx=6, pady=3)

        # Nom
        row_nom = tk.Frame(inner, bg=C["bg"])
        row_nom.pack(fill="x", pady=1)
        tk.Label(row_nom, text="Technicien", bg=C["bg"], font=("Arial", 9),
                 width=30, anchor="w").pack(side="left", padx=(0, 4))
        self.vars["nom"] = tk.StringVar()
        nom_e = ttk.Entry(row_nom, textvariable=self.vars["nom"], width=18)
        nom_e.pack(side="left")
        if on_change:
            self.vars["nom"].trace_add("write", lambda *_: on_change())

        for key, lbl in self.TEXTE_FIELDS:
            row = tk.Frame(inner, bg=C["bg"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text=lbl, bg=C["bg"], font=("Arial", 9),
                     width=30, anchor="w").pack(side="left", padx=(0, 4))
            self.vars[key] = tk.StringVar()
            ttk.Entry(row, textvariable=self.vars[key], width=18).pack(side="left")
            if on_change:
                self.vars[key].trace_add("write", lambda *_: on_change())

        sep = tk.Frame(inner, bg="#d0d4d9", height=1)
        sep.pack(fill="x", pady=(3, 1))
        chk_row = tk.Frame(inner, bg=C["bg"])
        chk_row.pack(fill="x", pady=2)
        for key, lbl in self.CHECK_FIELDS:
            self.check_vars[key] = tk.IntVar(value=0)
            tk.Checkbutton(chk_row, text=lbl,
                           variable=self.check_vars[key],
                           bg=C["bg"], font=("Arial", 9),
                           activebackground=C["bg"],
                           selectcolor="white").pack(side="left", padx=(0, 14))
            if on_change:
                self.check_vars[key].trace_add("write", lambda *_: on_change())

    def set_numero(self, n):
        self.numero = n
        self._num_lbl.config(text=f"Technicien {n}")

    def set_delete_visible(self, visible):
        self._del_btn.config(state="normal" if visible else "disabled",
                             fg="#c62828" if visible else "#b8c0c9")

    def load(self, data):
        if not data:
            data = {}
        for key in ["nom"] + [k for k, _ in self.TEXTE_FIELDS]:
            self.vars[key].set(str(data.get(key, "")))
        for key, _ in self.CHECK_FIELDS:
            v = data.get(key, 0)
            try:
                self.check_vars[key].set(1 if int(v) else 0)
            except (ValueError, TypeError):
                self.check_vars[key].set(1 if str(v).strip() else 0)

    def to_dict(self):
        out = {}
        for key in ["nom"] + [k for k, _ in self.TEXTE_FIELDS]:
            v = self.vars[key].get().strip()
            if v:
                out[key] = v
        for key, _ in self.CHECK_FIELDS:
            out[key] = int(self.check_vars[key].get())
        return out


class JourFrame(tk.Frame):
    """Une journee d'intervention : date + un ou plusieurs techniciens."""

    def __init__(self, master, numero, on_change=None, **kw):
        super().__init__(master, bg=C["bg"], relief="groove", bd=1, **kw)
        self.numero = numero
        self._tech_frames = []
        self._on_change = on_change
        self.date_var = tk.StringVar()
        if on_change:
            self.date_var.trace_add("write", lambda *_: on_change())

        head = tk.Frame(self, bg="#eef2f7")
        head.pack(fill="x")
        self._num_lbl = tk.Label(head, text=f"Jour {numero}",
                                  bg="#eef2f7", font=("Arial", 9, "bold"),
                                  fg="#002b5c")
        self._num_lbl.pack(side="left", padx=8, pady=3)
        self._del_btn = tk.Button(head, text="✕", bg="#eef2f7",
                                   fg="#c62828", bd=0,
                                   font=("Arial", 9, "bold"), cursor="hand2",
                                   activebackground="#eef2f7")
        self._del_btn.pack(side="right", padx=6, pady=1)

        date_row = tk.Frame(self, bg=C["bg"])
        date_row.pack(fill="x", padx=8, pady=(4, 2))
        tk.Label(date_row, text="Date", bg=C["bg"], font=("Arial", 9),
                 width=30, anchor="w").pack(side="left", padx=(0, 4))
        ttk.Entry(date_row, textvariable=self.date_var, width=18).pack(side="left")

        self._tech_container = tk.Frame(self, bg=C["bg"])
        self._tech_container.pack(fill="x", padx=8, pady=(2, 2))

        add_row = tk.Frame(self, bg=C["bg"])
        add_row.pack(fill="x", padx=8, pady=(2, 6))
        mk_btn(add_row, "+ Technicien", self._ajouter_tech,
               color=C["btn2"]).pack(side="left")

        self._ajouter_tech()

    def _ajouter_tech(self, data=None):
        tf = TechnicienFrame(self._tech_container,
                              len(self._tech_frames) + 1,
                              on_change=self._on_change)
        tf._del_btn.config(command=lambda f=tf: self._supprimer_tech(f))
        self._tech_frames.append(tf)
        tf.pack(fill="x", pady=(0, 4))
        if data:
            tf.load(data)
        self._update_tech_delete()

    def _supprimer_tech(self, tf):
        if len(self._tech_frames) <= 1:
            return
        self._tech_frames.remove(tf)
        tf.destroy()
        for i, f in enumerate(self._tech_frames, 1):
            f.set_numero(i)
        self._update_tech_delete()
        if self._on_change:
            self._on_change()

    def _update_tech_delete(self):
        show = len(self._tech_frames) > 1
        for tf in self._tech_frames:
            tf.set_delete_visible(show)

    def set_numero(self, n):
        self.numero = n
        self._num_lbl.config(text=f"Jour {n}")

    def set_delete_visible(self, visible):
        self._del_btn.config(state="normal" if visible else "disabled",
                             fg="#c62828" if visible else "#b8c0c9")

    def load(self, data):
        if not data:
            data = {}
        self.date_var.set(str(data.get("date", "")))
        techs = data.get("techniciens")
        if techs and isinstance(techs, list):
            while len(self._tech_frames) > len(techs):
                tf = self._tech_frames.pop(); tf.destroy()
            while len(self._tech_frames) < len(techs):
                self._ajouter_tech()
            if not self._tech_frames:
                self._ajouter_tech()
            for tf, td in zip(self._tech_frames, techs):
                tf.load(td)
            for i, tf in enumerate(self._tech_frames, 1):
                tf.set_numero(i)
            self._update_tech_delete()
        else:
            # Retrocompat : charger les champs plats comme technicien 1
            while len(self._tech_frames) > 1:
                tf = self._tech_frames.pop(); tf.destroy()
            self._update_tech_delete()
            self._tech_frames[0].load(data)

    def sync_techniciens(self, names):
        """Ajoute un TechnicienFrame pour chaque nom manquant dans ce jour.
        Réutilise les frames vides (sans nom) plutôt qu'en créer de nouveaux."""
        named_frames = [tf for tf in self._tech_frames if tf.vars["nom"].get().strip()]
        empty_frames = [tf for tf in self._tech_frames if not tf.vars["nom"].get().strip()]
        existing_names = {tf.vars["nom"].get().strip() for tf in named_frames}
        for name in names:
            if not name or name in existing_names:
                continue
            if empty_frames:
                tf = empty_frames.pop(0)
                tf.vars["nom"].set(name)
            else:
                self._ajouter_tech({"nom": name})
            existing_names.add(name)
        self._update_tech_delete()

    def to_dict(self):
        return {"date": self.date_var.get().strip(),
                "techniciens": [tf.to_dict() for tf in self._tech_frames]}


class DeplacementsTable(tk.Frame):
    """Saisie des temps et frais : plusieurs jours, plusieurs techniciens par jour."""

    _T_KEYS = [("trajet_aller_retour", "Distance totale"),
               ("duree_intervention",  "Duree intervention"),
               ("temps_preparation",   "Preparation"),
               ("temps_rangement",     "Rangement")]
    _F_KEYS = [("frais_repas", "Repas"),
               ("frais_hotel", "Hotel"),
               ("frais_peages", "Peages")]

    def __init__(self, master, **kw):
        super().__init__(master, bg=C["bg"], **kw)
        self._jour_frames = []
        self._total_job  = None
        self._total_vars = {}
        self._current_techniciens = []

        self._container = tk.Frame(self, bg=C["bg"])
        self._container.pack(fill="x")

        add_row = tk.Frame(self, bg=C["bg"])
        add_row.pack(fill="x", pady=(6, 0))
        mk_btn(add_row, "+ Ajouter un jour", self._ajouter_jour,
               color=C["btn2"]).pack(side="left")

        # Section TOTAUX
        tk.Frame(self, bg="#002b5c", height=2).pack(fill="x", pady=(10, 4))
        tk.Label(self, text="TOTAUX", bg=C["bg"],
                 font=("Arial", 9, "bold"), fg="#002b5c").pack(anchor="w")
        tot_f = tk.Frame(self, bg=C["bg"])
        tot_f.pack(fill="x")
        for key, lbl in self._T_KEYS:
            row = tk.Frame(tot_f, bg=C["bg"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text=lbl, bg=C["bg"], font=("Arial", 9),
                     width=30, anchor="w").pack(side="left", padx=(0, 4))
            self._total_vars[key] = tk.StringVar(value="—")
            tk.Label(row, textvariable=self._total_vars[key],
                     bg="#f0f4f8", font=("Arial", 9, "bold"), fg="#002b5c",
                     width=14, anchor="w", relief="groove",
                     padx=6).pack(side="left")

        frais_row = tk.Frame(tot_f, bg=C["bg"])
        frais_row.pack(fill="x", pady=(4, 0))
        for key, lbl in self._F_KEYS:
            self._total_vars[key] = tk.StringVar(value="0")
            tk.Label(frais_row, text=f"{lbl} :", bg=C["bg"],
                     font=("Arial", 9)).pack(side="left", padx=(0, 2))
            tk.Label(frais_row, textvariable=self._total_vars[key],
                     bg="#f0f4f8", font=("Arial", 9, "bold"), fg="#002b5c",
                     width=3, anchor="center",
                     relief="groove").pack(side="left", padx=(0, 10))

        self._ajouter_jour()

    # ── gestion jours ────────────────────────────────────────────────────────

    def _ajouter_jour(self, data=None):
        jf = JourFrame(self._container, len(self._jour_frames) + 1,
                       on_change=self._schedule_totals)
        jf._del_btn.config(command=lambda f=jf: self._supprimer_jour(f))
        self._jour_frames.append(jf)
        jf.pack(fill="x", pady=(0, 6))
        if data:
            jf.load(data)
        else:
            techs = self._current_techniciens
            if not techs and len(self._jour_frames) > 1:
                techs = [tf.vars["nom"].get().strip()
                         for tf in self._jour_frames[0]._tech_frames
                         if tf.vars["nom"].get().strip()]
            if techs:
                jf.sync_techniciens(techs)
        self._update_jour_delete()
        self._schedule_totals()

    def _supprimer_jour(self, jf):
        if len(self._jour_frames) <= 1:
            return
        self._jour_frames.remove(jf)
        jf.destroy()
        for i, f in enumerate(self._jour_frames, 1):
            f.set_numero(i)
        self._update_jour_delete()
        self._schedule_totals()

    def _update_jour_delete(self):
        show = len(self._jour_frames) > 1
        for jf in self._jour_frames:
            jf.set_delete_visible(show)

    # ── totaux ───────────────────────────────────────────────────────────────

    def _schedule_totals(self):
        if self._total_job is not None:
            self.after_cancel(self._total_job)
        self._total_job = self.after(250, self._recalculate_totals)

    def _recalculate_totals(self):
        self._total_job = None
        sums  = {k: None for k, _ in self._T_KEYS}
        frais = {k: 0    for k, _ in self._F_KEYS}
        for jf in self._jour_frames:
            for tf in jf._tech_frames:
                for key, _ in self._T_KEYS:
                    parse = _parse_km if key == "trajet_aller_retour" else _parse_h
                    v = parse(tf.vars[key].get())
                    if v is not None:
                        sums[key] = (sums[key] or 0) + v
                for key, _ in self._F_KEYS:
                    frais[key] += int(tf.check_vars[key].get())
        for key, _ in self._T_KEYS:
            if sums[key] is None:
                self._total_vars[key].set("—")
            elif key == "trajet_aller_retour":
                self._total_vars[key].set(_fmt_km(sums[key]))
            else:
                self._total_vars[key].set(_fmt_h(sums[key]))
        for key, _ in self._F_KEYS:
            self._total_vars[key].set(str(frais[key]))

    # ── load / save ──────────────────────────────────────────────────────────

    def load(self, data):
        if not data:
            data = {}
        jours = data.get("jours")
        if jours and isinstance(jours, list):
            while len(self._jour_frames) > len(jours):
                jf = self._jour_frames.pop(); jf.destroy()
            while len(self._jour_frames) < len(jours):
                self._ajouter_jour()
            if not self._jour_frames:
                self._ajouter_jour()
            for jf, jdata in zip(self._jour_frames, jours):
                jf.load(jdata)
            for i, jf in enumerate(self._jour_frames, 1):
                jf.set_numero(i)
            self._update_jour_delete()
        else:
            # Retrocompat : ancien format plat
            while len(self._jour_frames) > 1:
                jf = self._jour_frames.pop(); jf.destroy()
            self._update_jour_delete()
            converted = {}
            for key in ["trajet_aller_retour", "duree_intervention",
                        "temps_preparation", "temps_rangement"]:
                val = data.get(key, "")
                if not val:
                    if key == "trajet_aller_retour":
                        a = str(data.get("trajet_aller", "")).strip()
                        r = str(data.get("trajet_retour", "")).strip()
                        val = f"Aller : {a} / Retour : {r}" if a and r else a or r
                    elif key == "duree_intervention":
                        deb = str(data.get("heure_debut_matin", "")).strip()
                        fin = str(data.get("heure_fin_apres", "")).strip()
                        val = f"{deb} -> {fin}" if deb and fin else deb or fin
                converted[key] = val
            for key in ["frais_repas", "frais_hotel", "frais_peages"]:
                converted[key] = data.get(key, 0)
            self._jour_frames[0].load(
                {"date": "", "techniciens": [converted]})
        self._recalculate_totals()

    def to_dict(self):
        return {"jours": [jf.to_dict() for jf in self._jour_frames]}

    def sync_techniciens(self, names):
        """Met à jour tous les jours avec les techniciens du bon."""
        self._current_techniciens = list(names)
        for jf in self._jour_frames:
            jf.sync_techniciens(names)
        self._schedule_totals()

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
# DIALOG : MOTEURS D'UN CLIENT
# ══════════════════════════════════════════════════════════════════════════════
class ClientMoteursDialog(tk.Toplevel):
    """Affiche la liste des moteurs d'un client donne avec acces a leur fiche."""

    def __init__(self, parent, app, client):
        super().__init__(parent)
        self.app = app
        self.client = client
        self.title(f"Moteurs du client - {client['nom']}")
        self.geometry("900x500")
        self.configure(bg=C["bg"])
        self.grab_set()

        # En-tete
        mk_header(self, f"🚢 Moteurs de {client['nom']}",
                  f"   {client.get('contact', '') or ''}")

        # Tableau
        cols = ("ns", "navire", "marque", "type", "annee")
        col_defs = [
            ("ns",     "N° Série",        140),
            ("navire", "Navire / Site",   200),
            ("marque", "Marque",          120),
            ("type",   "Type moteur",     180),
            ("annee",  "Mise en service", 120),
        ]
        tf, self.tree = mk_tree(self, cols, col_defs, height=18)
        tf.pack(fill="both", expand=True, padx=20, pady=8)
        self.tree.bind("<Double-1>", lambda e: self._ouvrir_fiche())

        # Boutons
        bf = tk.Frame(self, bg=C["bg"])
        bf.pack(fill="x", padx=20, pady=8)
        mk_btn(bf, "🔧 Voir la fiche", self._ouvrir_fiche).pack(side="left", padx=4)
        mk_btn(bf, "Fermer", self.destroy, color="#888").pack(side="right", padx=4)

        # Charger la liste
        self._charger()

    def _charger(self):
        # Recupere tous les moteurs et filtre sur le client_id
        all_moteurs = list(db.get_moteurs())
        self._moteurs = [m for m in all_moteurs
                         if m["client_id"] == self.client["id"]]
        if not self._moteurs:
            # Pas de moteurs : message visible
            fill_tree(self.tree, [])
            self.title(f"Moteurs du client - {self.client['nom']} (aucun)")
            return
        fill_tree(self.tree, [
            (m["num_serie"],
             row_get(m, "navire"),
             row_get(m, "marque"),
             row_get(m, "type_moteur") or row_get(m, "ref_constructeur"),
             row_get(m, "date_mise_service"))
            for m in self._moteurs
        ])

    def _ouvrir_fiche(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection",
                "Sélectionnez un moteur pour voir sa fiche.")
            return
        m = self._moteurs[int(sel[0])]
        MoteurFicheDialog(self, self.app, dict(m))

# ══════════════════════════════════════════════════════════════════════════════
# FICHE MOTEUR — Historique interventions + garanties
# ══════════════════════════════════════════════════════════════════════════════
class MoteurFicheDialog(tk.Toplevel):
    """Fiche detaillee d'un moteur : infos + historique interventions + garanties."""
    def __init__(self, parent, app, moteur):
        super().__init__(parent)
        self.app = app
        self.moteur = moteur
        self.title(f"Fiche moteur — {moteur.get('num_serie', '')}")
        self.configure(bg=C["bg"])
        self.geometry("900x680")
        self.grab_set()

        mk_header(self, f"Moteur {moteur.get('num_serie', '')}",
                   f"   {moteur.get('marque', '')} — {moteur.get('client_nom', '')}")

        # ── Bandeau d'infos ────────────────────────────────────────────────────
        info = tk.Frame(self, bg=C["surface"], relief="solid", bd=1)
        info.pack(fill="x", padx=16, pady=(10, 6))
        grille = tk.Frame(info, bg=C["surface"])
        grille.pack(fill="x", padx=12, pady=8)

        statut, jours = db.garantie_status(moteur.get("date_mise_service", ""),
                                            moteur.get("duree_garantie", ""))
        if statut == "Active":
            gtxt, gcol = f"Garantie active ({jours} j restants)", "#1e7e3e"
        elif statut == "Expiree":
            gtxt, gcol = f"Garantie expirée (depuis {jours} j)", "#c62828"
        else:
            gtxt, gcol = "Garantie : —", "#6b7785"

        cu_parts = [
            moteur.get("client_utilisateur_nom", ""),
            moteur.get("client_utilisateur_tel", ""),
            moteur.get("client_utilisateur_email", ""),
            moteur.get("client_utilisateur_adresse", ""),
        ]
        client_utilisateur_str = "  |  ".join(p for p in cu_parts if p)
        champs = [
            ("N° série", moteur.get("num_serie", "")),
            ("Marque", moteur.get("marque", "")),
            ("Type moteur", moteur.get("type_moteur", "")),
            ("Client", moteur.get("client_nom", "")),
            ("Navire / Site", moteur.get("navire", "")),
            ("Machine", moteur.get("machine", "")),
            ("Réf. constructeur", moteur.get("ref_constructeur", "")),
            ("Code affaire", moteur.get("code_affaire", "")),
            ("Mise en service", moteur.get("date_mise_service", "")),
            ("Durée garantie", f"{moteur.get('duree_garantie', '')} mois"
                if moteur.get("duree_garantie") else ""),
            ("Client utilisateur", client_utilisateur_str),
        ]
        for i, (lbl, val) in enumerate(champs):
            r, c = divmod(i, 2)
            cell = tk.Frame(grille, bg=C["surface"])
            cell.grid(row=r, column=c, sticky="w", padx=10, pady=2)
            tk.Label(cell, text=f"{lbl} : ", bg=C["surface"],
                     font=("Arial", 9, "bold"), fg="#555").pack(side="left")
            if lbl == "Code affaire" and val:
                def _open_affaire_folder(code=val):
                    try:
                        affaires = db.get_affaires(search=code)
                        match = next((a for a in affaires
                                      if a.get("num_affaire") == code
                                      or a.get("ref_interne") == code), None)
                        if match and match.get("dossier_path"):
                            import os as _os
                            p = Path(match["dossier_path"])
                            if p.exists():
                                _os.startfile(str(p))
                                return
                        messagebox.showinfo("Dossier affaire",
                            f"Aucun dossier trouvé pour le code '{code}'.\n"
                            "Vérifiez que l'affaire existe dans l'app Affaires.")
                    except Exception as e:
                        messagebox.showerror("Erreur", str(e))
                lnk = tk.Label(cell, text=str(val), bg=C["surface"],
                               font=("Arial", 9, "underline"), fg="#d97706",
                               cursor="hand2")
                lnk.pack(side="left")
                lnk.bind("<Button-1>", lambda _, fn=_open_affaire_folder: fn())
            else:
                tk.Label(cell, text=str(val) or "—", bg=C["surface"],
                         font=("Arial", 9), fg="#1a2332").pack(side="left")
        # Badge garantie
        tk.Label(grille, text=gtxt, bg=C["surface"], fg=gcol,
                 font=("Arial", 10, "bold")).grid(
            row=len(champs)//2 + 1, column=0, columnspan=2,
            sticky="w", padx=10, pady=(6, 2))

        # ── Onglets Interventions / Garanties ───────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=16, pady=8)

        # --- Onglet Interventions ---
        tab_inv = tk.Frame(nb, bg=C["bg"])
        nb.add(tab_inv, text="  Interventions  ")
        cols = ("num_bon", "date", "type", "statut", "tech")
        col_defs = [("num_bon", "N° Bon", 120), ("date", "Date", 95),
                    ("type", "Type", 160), ("statut", "Statut", 100),
                    ("tech", "Technicien", 180)]
        tf, self.tree_inv = mk_tree(tab_inv, cols, col_defs, height=9)
        tf.pack(fill="both", expand=True, padx=8, pady=(6, 2))
        self._inv_cache = []
        af_inv = tk.Frame(tab_inv, bg=C["bg"]); af_inv.pack(fill="x", padx=8, pady=(0, 6))
        mk_btn(af_inv, "📁 Ouvrir dossier", self._ouvrir_dossier_inv).pack(side="left", padx=4)
        self.tree_inv.bind("<Double-1>", lambda _: self._ouvrir_dossier_inv())

        # --- Onglet Garanties ---
        tab_gar = tk.Frame(nb, bg=C["bg"])
        nb.add(tab_gar, text="  Garanties  ")
        cols2 = ("num_ems", "attribution", "statut", "ouv", "clo")
        col_defs2 = [("num_ems", "N° EMS", 130),
                     ("attribution", "Attribution", 130),
                     ("statut", "Statut", 180),
                     ("ouv", "Ouverture", 95), ("clo", "Clôture", 95)]
        tf2, self.tree_gar = mk_tree(tab_gar, cols2, col_defs2, height=9)
        tf2.pack(fill="both", expand=True, padx=8, pady=(6, 2))
        self._gar_cache = []
        af_gar = tk.Frame(tab_gar, bg=C["bg"]); af_gar.pack(fill="x", padx=8, pady=(0, 6))
        mk_btn(af_gar, "📁 Ouvrir dossier", self._ouvrir_dossier_gar).pack(side="left", padx=4)
        self.tree_gar.bind("<Double-1>", lambda _: self._ouvrir_dossier_gar())

        # --- Onglet Sous-ensembles ---
        # Un sous-ensemble est un Moteur comme un autre (mêmes champs),
        # rattaché à ce moteur via parent_moteur_id.
        tab_se = tk.Frame(nb, bg=C["bg"])
        nb.add(tab_se, text="  Sous-ensembles  ")
        cols3 = ("num_serie", "marque", "machine", "type_moteur", "date")
        col_defs3 = [
            ("num_serie",   "N° Série",         120),
            ("marque",      "Marque",            90),
            ("machine",     "Machine",          110),
            ("type_moteur", "Type moteur",      130),
            ("date",        "Mise en service",  110),
        ]
        tf3, self.tree_se = mk_tree(tab_se, cols3, col_defs3, height=9)
        tf3.pack(fill="both", expand=True, padx=8, pady=(6, 2))
        self._se_cache = []
        af_se = tk.Frame(tab_se, bg=C["bg"]); af_se.pack(fill="x", padx=8, pady=(0, 6))
        mk_btn(af_se, "➕ Ajouter", self._se_add).pack(side="left", padx=4)
        mk_btn(af_se, "✏️ Modifier", self._se_edit, color=C["btn2"]).pack(side="left", padx=4)
        mk_btn(af_se, "🗑️ Supprimer", self._se_del, color=C["danger"]).pack(side="left", padx=4)
        self.tree_se.bind("<Double-1>", lambda _: self._se_edit())

        # Boutons
        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=10)
        mk_btn(bf, "Fermer", self.destroy, color="#888").pack(side="left", padx=6)

        self._charger()

    def _charger(self):
        mid = self.moteur.get("id")
        # Interventions
        try:
            self._inv_cache = list(db.get_interventions_for_moteur(mid))
        except Exception as e:
            self._inv_cache = []
            print(f"[FicheMoteur] interventions : {e}")
        self.tree_inv.delete(*self.tree_inv.get_children())
        if not self._inv_cache:
            self.tree_inv.insert("", "end",
                                 values=("—", "(aucune intervention)", "", "", ""))
        else:
            for i, inv in enumerate(self._inv_cache):
                self.tree_inv.insert("", "end", iid=str(i),
                    values=(inv.get("num_bon", ""),
                            inv.get("date_creation", ""),
                            inv.get("type_intervention", ""),
                            inv.get("statut", ""),
                            inv.get("technicien", "")))
        # Garanties
        try:
            self._gar_cache = list(db.get_garanties_moteur(mid))
        except Exception as e:
            self._gar_cache = []
            print(f"[FicheMoteur] garanties : {e}")
        self.tree_gar.delete(*self.tree_gar.get_children())
        if not self._gar_cache:
            self.tree_gar.insert("", "end",
                                 values=("—", "(aucune garantie)", "", "", ""))
        else:
            for g in self._gar_cache:
                self.tree_gar.insert("", "end",
                    values=(g.get("num_ems", ""),
                            g.get("attribution", ""),
                            g.get("statut", ""),
                            g.get("date_ouverture", ""),
                            g.get("date_cloture", "")))
        # Sous-ensembles
        try:
            self._se_cache = list(db.get_sous_ensembles(mid))
        except Exception as e:
            self._se_cache = []
            print(f"[FicheMoteur] sous-ensembles : {e}")
        self.tree_se.delete(*self.tree_se.get_children())
        if not self._se_cache:
            self.tree_se.insert("", "end",
                                values=("(aucun sous-ensemble)", "", "", "", ""))
        else:
            for i, se in enumerate(self._se_cache):
                tags = ("even" if i % 2 == 0 else "odd",)
                self.tree_se.insert("", "end", iid=f"se_{i}",
                    values=(se.get("num_serie", ""),
                            se.get("marque", ""),
                            se.get("machine", ""),
                            se.get("type_moteur", ""),
                            se.get("date_mise_service", "")),
                    tags=tags)

    def _se_add(self):
        MoteurDialog(self, self.app, parent_moteur=dict(self.moteur),
                     on_save=self._charger)

    def _se_edit(self):
        sel = self.tree_se.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un sous-ensemble.", parent=self)
            return
        idx = int(sel[0].replace("se_", ""))
        if idx >= len(self._se_cache):
            return
        MoteurDialog(self, self.app, moteur=dict(self._se_cache[idx]),
                     on_save=self._charger)

    def _se_del(self):
        sel = self.tree_se.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un sous-ensemble.", parent=self)
            return
        idx = int(sel[0].replace("se_", ""))
        if idx >= len(self._se_cache):
            return
        se = self._se_cache[idx]
        if not messagebox.askyesno("Supprimer",
                f"Supprimer le sous-ensemble « {se.get('num_serie', '')} » ?\n\n"
                "Cette action est irréversible.", parent=self):
            return
        try:
            db.delete_moteur(se["id"])
            self._charger()
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)

    def _ouvrir_dossier_inv(self):
        sel = self.tree_inv.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez une intervention.", parent=self)
            return
        idx = int(sel[0])
        if idx >= len(self._inv_cache):
            return
        num_bon = self._inv_cache[idx].get("num_bon", "")
        if not num_bon:
            return
        d = _dossiers_root() / num_bon
        if not d.exists():
            if messagebox.askyesno("Dossier absent",
                    f"Le dossier {num_bon} n'existe pas encore.\nVoulez-vous le créer ?",
                    parent=self):
                d.mkdir(parents=True, exist_ok=True)
            else:
                return
        ouvrir_fichier(d)

    def _ouvrir_dossier_gar(self):
        sel = self.tree_gar.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez une garantie.", parent=self)
            return
        try:
            idx = list(self.tree_gar.get_children()).index(sel[0])
        except ValueError:
            return
        if idx >= len(self._gar_cache):
            return
        num_ems = self._gar_cache[idx].get("num_ems", "")
        if not num_ems:
            return
        d = _dossiers_root() / num_ems
        if not d.exists():
            if messagebox.askyesno("Dossier absent",
                    f"Le dossier {num_ems} n'existe pas encore.\nVoulez-vous le créer ?",
                    parent=self):
                d.mkdir(parents=True, exist_ok=True)
            else:
                return
        ouvrir_fichier(d)


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG MOTEUR
# ══════════════════════════════════════════════════════════════════════════════
class MoteurDialog(tk.Toplevel):
    FIELDS = [
        ("num_serie",                  "N° Série *"),
        ("navire",                     "Navire / Site"),
        ("machine",                    "Machine"),
        ("type_moteur",                "Type moteur / inverseur"),
        ("marque",                     "Marque"),
        ("ref_constructeur",           "Réf. Constructeur"),
        ("cylindree",                  "Cylindrée"),
        ("famille",                    "Famille"),
        ("application",                "Application"),
        ("typologie",                  "Typologie"),
        ("collection",                 "Collection"),
        ("code_affaire",               "Code Affaire"),
        ("date_mise_service",          "Mise en service"),
        ("duree_garantie",             "Garantie (mois)"),
        ("client_utilisateur_nom",     "Client utilisateur"),
        ("client_utilisateur_email",   "Email utilisateur"),
        ("client_utilisateur_tel",     "Tél. utilisateur"),
        ("client_utilisateur_adresse", "Adresse utilisateur"),
    ]

    # Champs hérités du moteur principal à la création d'un sous-ensemble :
    # uniquement le contexte d'installation / client final. Un sous-ensemble
    # est un équipement à part entière (ex : l'inverseur d'un moteur), donc
    # sa marque/type/référence/cylindrée etc. ne sont PAS préremplis.
    INHERITED_FIELDS = {
        "navire", "machine", "code_affaire", "type_client",
        "date_mise_service", "duree_garantie",
        "client_utilisateur_nom", "client_utilisateur_email",
        "client_utilisateur_tel", "client_utilisateur_adresse",
    }

    def __init__(self, parent, app, moteur=None, on_save=None, parent_moteur=None):
        """parent_moteur : moteur principal, fourni uniquement à la création
        d'un sous-ensemble (mêmes champs qu'un moteur, rattaché via
        parent_moteur_id). Sert aussi à préremplir les champs partagés
        (client, navire, machine, marque…) pour gagner du temps de saisie."""
        super().__init__(parent)
        self.app = app; self.moteur = moteur; self.on_save = on_save
        self.parent_moteur = parent_moteur
        if moteur:
            titre = f"Modifier – {moteur['num_serie']}"
        elif parent_moteur:
            titre = f"Nouveau sous-ensemble — Moteur {parent_moteur.get('num_serie', '')}"
        else:
            titre = "Nouveau moteur"
        self.title(titre)
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.grab_set()
        self._clients = list(db.get_clients())
        self._types_moteur = db.get_types_moteur()
        self._type_to_marque = {t["libelle"]: t["marque"] for t in self._types_moteur}
        self.client_var = tk.StringVar()
        client_src = moteur or parent_moteur
        if client_src and client_src.get("client_id"):
            c = next((x for x in self._clients if x["id"]==client_src["client_id"]),None)
            if c: self.client_var.set(c["nom"])
        if parent_moteur and not moteur:
            ctx = tk.Frame(self, bg=C["surface"], relief="solid", bd=1)
            ctx.pack(fill="x", padx=24, pady=(14, 0))
            tk.Label(ctx,
                     text=f"Sous-ensemble du moteur {parent_moteur.get('num_serie', '—')} "
                          f"({parent_moteur.get('marque', '') or '—'} — "
                          f"{parent_moteur.get('navire', '') or '—'})",
                     bg=C["surface"], font=("Arial", 9, "bold"),
                     fg=C["header"]).pack(anchor="w", padx=10, pady=6)
        f = tk.Frame(self, bg=C["bg"])
        f.pack(padx=24, pady=18)
        tk.Label(f, text="Client *", bg=C["bg"], font=("Arial",10),
                 anchor="e", width=22).grid(row=0, column=0, sticky="ne", padx=(0,6), pady=4)
        self._client_combo = SearchableCombobox(
            f, textvariable=self.client_var,
            values=[c["nom"] for c in self._clients],
            width=33)
        self._client_combo.grid(row=0, column=1, pady=4, sticky="w")

        def _default(key):
            if moteur:
                return moteur.get(key, "")
            if parent_moteur and key in self.INHERITED_FIELDS:
                return parent_moteur.get(key, "")
            return ""
        self.v = {k: tk.StringVar(value=_default(k)) for k, _ in self.FIELDS}
        self._marque_combo = None
        self._type_moteur_combo = None
        for i,(key,lbl) in enumerate(self.FIELDS, start=1):
            tk.Label(f, text=lbl, bg=C["bg"], font=("Arial",10),
                     anchor="e", width=22).grid(row=i, column=0, sticky="e", padx=(0,6), pady=4)
            if key == "date_mise_service":
                DateEntry(f, textvariable=self.v[key], width=32).grid(row=i, column=1, pady=4)
            elif key == "client_utilisateur_email":
                EmailEntry(f, textvariable=self.v[key], width=34).grid(row=i, column=1, pady=4)
            elif key == "client_utilisateur_adresse":
                ttk.Entry(f, textvariable=self.v[key], width=34).grid(row=i, column=1, pady=4)
            elif key == "type_moteur":
                row_t = tk.Frame(f, bg=C["bg"])
                row_t.grid(row=i, column=1, pady=4, sticky="w")
                self._type_moteur_combo = SearchableCombobox(
                    row_t, textvariable=self.v["type_moteur"],
                    values=[t["libelle"] for t in self._types_moteur],
                    width=30, min_chars_to_open=1, allow_free_text=True)
                self._type_moteur_combo.pack(side="left")
                self._type_moteur_combo.bind("<<ComboboxSelected>>",
                                              self._on_type_moteur_selected)
                def _open_types_moteur():
                    TypesMoteurDialog(self, self.app,
                        on_close=self._reload_types_moteur)
                mk_btn(row_t, "⚙", _open_types_moteur, color=C["btn2"]).pack(side="left", padx=(8,0))
            elif key == "marque":
                row_m = tk.Frame(f, bg=C["bg"])
                row_m.grid(row=i, column=1, pady=4, sticky="w")
                self._marque_combo = SearchableCombobox(
                    row_m, textvariable=self.v["marque"],
                    values=db.get_marques(), width=30)
                self._marque_combo.pack(side="left")
                self._marque_combo.bind("<<ComboboxSelected>>",
                                         self._on_marque_selected)
                def _open_marques():
                    MarquesDialog(self, self.app,
                        on_close=lambda: self._marque_combo.set_values(db.get_marques()))
                mk_btn(row_m, "⚙", _open_marques, color=C["btn2"]).pack(side="left", padx=(8,0))
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
        data = {"client_id":client["id"] if client else "",
                **{k:self.v[k].get().strip() for k in self.v}}
        if self.parent_moteur and not self.moteur:
            data["parent_moteur_id"] = self.parent_moteur["id"]
        db.upsert_moteur(data, moteur_id=self.moteur["id"] if self.moteur else None)
        if self.on_save: self.on_save()
        self.destroy()

    def _on_type_moteur_selected(self, *_):
        libelle = self.v["type_moteur"].get()
        marque = self._type_to_marque.get(libelle, "")
        if marque and not self.v["marque"].get().strip():
            self.v["marque"].set(marque)

    def _on_marque_selected(self, *_):
        marque = self.v["marque"].get().strip()
        if marque:
            filtered = [t["libelle"] for t in self._types_moteur
                        if t.get("marque") == marque]
            types = filtered if filtered else [t["libelle"] for t in self._types_moteur]
        else:
            types = [t["libelle"] for t in self._types_moteur]
        if self._type_moteur_combo:
            self._type_moteur_combo.set_values(types)

    def _reload_types_moteur(self):
        self._types_moteur = db.get_types_moteur()
        self._type_to_marque = {t["libelle"]: t["marque"] for t in self._types_moteur}
        if self._type_moteur_combo:
            self._type_moteur_combo.set_values(
                [t["libelle"] for t in self._types_moteur])


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
        # Callback : on tente d'envoyer le nom (compat ancien/nouveau callback)
        if self.on_save:
            try:
                self.on_save(nom)
            except TypeError:
                # Callback sans argument (compatibilité)
                self.on_save()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# DIALOG TYPES D'INTERVENTION
# ══════════════════════════════════════════════════════════════════════════════
class MarquesDialog(tk.Toplevel):
    def __init__(self, parent, app, on_close=None):
        super().__init__(parent)
        self.app = app
        self.on_close = on_close
        self.title("Marques moteur")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.geometry("360x420")
        self.grab_set()

        tk.Label(self, text="Marques moteur", bg=C["bg"],
                 font=("Arial",11,"bold"), fg=C["header"]).pack(pady=(14,4))

        lf = tk.Frame(self, bg=C["bg"]); lf.pack(fill="both", expand=True, padx=16, pady=10)
        self.lst = tk.Listbox(lf, font=("Arial",10), height=12,
                              selectbackground=C["btn"], selectforeground="white")
        self.lst.pack(fill="both", expand=True)

        af = tk.Frame(self, bg=C["bg"]); af.pack(fill="x", padx=16, pady=(0,6))
        self.input_var = tk.StringVar()
        ttk.Entry(af, textvariable=self.input_var, width=26).pack(side="left", padx=(0,4))
        mk_btn(af, "➕ Ajouter", self._add).pack(side="left")

        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=10)
        mk_btn(bf, "✏️ Renommer", self._rename).pack(side="left", padx=4)
        mk_btn(bf, "🗑️ Supprimer", self._del, color=C["danger"]).pack(side="left", padx=4)
        mk_btn(bf, "Fermer", self._close, color="#888").pack(side="left", padx=4)
        self._refresh()

    def _refresh(self):
        self.lst.delete(0, "end")
        for m in db.get_marques():
            self.lst.insert("end", m)

    def _close(self):
        if self.on_close:
            self.on_close()
        self.destroy()

    def _add(self):
        v = self.input_var.get().strip()
        if not v: return
        if not db.add_marque(v):
            messagebox.showwarning("Doublon", f"La marque '{v}' existe déjà.")
            return
        self.input_var.set("")
        self._refresh()

    def _selected(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez une marque.")
            return None
        return self.lst.get(sel[0])

    def _rename(self):
        cur = self._selected()
        if not cur: return
        new = simpledialog.askstring("Renommer", f"Nouveau nom pour '{cur}' :",
                                      parent=self, initialvalue=cur)
        if new and new.strip() and new.strip() != cur:
            if db.update_marque(cur, new.strip()):
                self._refresh()
            else:
                messagebox.showwarning("Erreur", "Ce nom existe déjà.")

    def _del(self):
        cur = self._selected()
        if not cur: return
        if messagebox.askyesno("Supprimer",
                f"Supprimer la marque '{cur}' ?\n\n"
                "(Les bons existants conservent leur libellé.)"):
            db.delete_marque(cur)
            self._refresh()


class TypesMoteurDialog(tk.Toplevel):
    def __init__(self, parent, app, on_close=None):
        super().__init__(parent)
        self.app = app
        self.on_close = on_close
        self._items = []
        self.title("Types de moteur")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.geometry("420x500")
        self.grab_set()

        tk.Label(self, text="Types de moteur", bg=C["bg"],
                 font=("Arial",11,"bold"), fg=C["header"]).pack(pady=(14,4))

        lf = tk.Frame(self, bg=C["bg"]); lf.pack(fill="both", expand=True, padx=16, pady=4)
        self.lst = tk.Listbox(lf, font=("Arial",10), height=9,
                              selectbackground=C["btn"], selectforeground="white")
        self.lst.pack(fill="both", expand=True)
        self.lst.bind("<<ListboxSelect>>", self._on_select)

        ff = tk.Frame(self, bg=C["bg"]); ff.pack(fill="x", padx=16, pady=(6,2))
        tk.Label(ff, text="Type :", bg=C["bg"], width=17, anchor="e",
                 font=("Arial",10)).grid(row=0, column=0, padx=(0,6), pady=4)
        self.input_libelle = tk.StringVar()
        ttk.Entry(ff, textvariable=self.input_libelle, width=26).grid(row=0, column=1, sticky="w")
        tk.Label(ff, text="Marque associée :", bg=C["bg"], width=17, anchor="e",
                 font=("Arial",10)).grid(row=1, column=0, padx=(0,6), pady=4)
        self.input_marque = tk.StringVar()
        self._marque_cb = SearchableCombobox(ff, textvariable=self.input_marque,
                                              values=db.get_marques(), width=24,
                                              allow_free_text=True)
        self._marque_cb.grid(row=1, column=1, sticky="w")

        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=10)
        mk_btn(bf, "➕ Ajouter", self._add).pack(side="left", padx=4)
        mk_btn(bf, "✏️ Modifier", self._edit).pack(side="left", padx=4)
        mk_btn(bf, "🗑️ Supprimer", self._del, color=C["danger"]).pack(side="left", padx=4)
        mk_btn(bf, "Fermer", self._close, color="#888").pack(side="left", padx=4)
        self._refresh()

    def _refresh(self):
        self._items = db.get_types_moteur()
        self.lst.delete(0, "end")
        for t in self._items:
            label = t["libelle"]
            if t.get("marque"):
                label += f"  ← {t['marque']}"
            self.lst.insert("end", label)

    def _on_select(self, *_):
        sel = self.lst.curselection()
        if not sel: return
        item = self._items[sel[0]]
        self.input_libelle.set(item["libelle"])
        self.input_marque.set(item.get("marque") or "")

    def _close(self):
        if self.on_close:
            self.on_close()
        self.destroy()

    def _add(self):
        lib = self.input_libelle.get().strip()
        if not lib: return
        marque = self.input_marque.get().strip()
        if not db.add_type_moteur(lib, marque):
            messagebox.showwarning("Doublon", f"Le type '{lib}' existe déjà.")
            return
        self.input_libelle.set("")
        self.input_marque.set("")
        self._refresh()

    def _selected_item(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showwarning("Sélection", "Sélectionnez un type dans la liste.")
            return None
        return self._items[sel[0]]

    def _edit(self):
        item = self._selected_item()
        if not item: return
        new_lib = self.input_libelle.get().strip()
        new_marque = self.input_marque.get().strip()
        if not new_lib: return
        if not db.update_type_moteur(item["libelle"], new_lib, new_marque):
            messagebox.showwarning("Erreur", "Ce nom existe déjà.")
            return
        self._refresh()

    def _del(self):
        item = self._selected_item()
        if not item: return
        if messagebox.askyesno("Supprimer",
                f"Supprimer le type '{item['libelle']}' ?\n\n"
                "(Les moteurs existants conservent leur libellé.)"):
            db.delete_type_moteur(item["libelle"])
            self.input_libelle.set("")
            self.input_marque.set("")
            self._refresh()


class TypesDialog(tk.Toplevel):
    def __init__(self, parent, app, on_close=None):
        super().__init__(parent)
        self.app = app
        self._on_close = on_close
        self.title("Types d'intervention")
        self.resizable(False, False)
        self.configure(bg=C["bg"])
        self.geometry("400x460")
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)

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
        mk_btn(bf, "Fermer", self._close, color="#888").pack(side="left", padx=4)
        self._refresh()

    def _close(self):
        if self._on_close:
            self._on_close()
        self.destroy()

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
# DIALOG SELECTION PHOTOS POUR ANNEXE
# ══════════════════════════════════════════════════════════════════════════════
class PhotosAnnexeDialog(tk.Toplevel):
    """Popup : choisir et ordonner les photos du dossier a inclure en annexe.
    Retourne la liste ordonnee des chemins coches via le callback on_valider."""
    EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

    def __init__(self, parent, dossier, on_valider):
        super().__init__(parent)
        self.on_valider = on_valider
        self.dossier = Path(dossier)
        self._vars = {}
        self._thumbs = []
        self._thumb_cache = {}
        self._row_frames = {}   # str(img_path) -> Frame
        self._order = []        # Path objects dans l'ordre d'affichage
        self._drag_src = None   # cle de la ligne en cours de glissement
        self.title("Photos à inclure en annexe")
        self.configure(bg=C["bg"])
        self.geometry("920x740")
        self.grab_set()

        # Lister les images du dossier
        self._order = sorted(
            [f for f in self.dossier.iterdir()
             if f.is_file() and f.suffix.lower() in self.EXTS],
            key=lambda p: p.name.lower()) if self.dossier.is_dir() else []
        self._load_saved_order()

        tk.Label(self, text="📷 Photos disponibles dans le dossier",
                 bg=C["bg"], font=("Arial", 12, "bold"),
                 fg=C["header"]).pack(pady=(14, 2))
        tk.Label(self, text=f"Dossier : {self.dossier.name}",
                 bg=C["bg"], font=("Arial", 9), fg="#666").pack()

        if not self._order:
            tk.Label(self, text="\n(Aucune photo dans ce dossier)\n\n"
                                "Ajoutez des photos via « Ajouter fichier(s) » "
                                "puis régénérez le bon.",
                     bg=C["bg"], font=("Arial", 10), fg="#888",
                     justify="center").pack(pady=40)
            bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=10)
            mk_btn(bf, "Générer sans photos",
                   lambda: self._valider(), color=C["btn2"]).pack(side="left", padx=6)
            mk_btn(bf, "Annuler", self._annuler,
                   color="#888").pack(side="left", padx=6)
            return

        # Barre d'actions (tout cocher / décocher)
        ab = tk.Frame(self, bg=C["bg"]); ab.pack(fill="x", padx=20, pady=(8, 2))
        mk_btn(ab, "☑ Tout cocher", lambda: self._toggle_all(True),
               color=C["btn3"]).pack(side="left", padx=3)
        mk_btn(ab, "☐ Tout décocher", lambda: self._toggle_all(False),
               color=C["btn3"]).pack(side="left", padx=3)
        tk.Label(ab, text="  ↕ Glisser ⠿ pour réordonner",
                 bg=C["bg"], font=("Arial", 8), fg="#888").pack(side="left", padx=8)

        # Zone scrollable avec miniatures
        cont = tk.Frame(self, bg=C["bg"])
        cont.pack(fill="both", expand=True, padx=20, pady=6)
        self._canvas = tk.Canvas(cont, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(cont, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._inner = tk.Frame(self._canvas, bg=C["bg"])
        _win = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                   lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                    lambda e: self._canvas.itemconfig(_win, width=e.width))

        def _wheel(ev):
            self._canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        self._canvas.bind("<Enter>",
                    lambda e: self._canvas.bind_all("<MouseWheel>", _wheel))
        self._canvas.bind("<Leave>",
                    lambda e: self._canvas.unbind_all("<MouseWheel>"))

        for img_path in self._order:
            self._add_photo_row(img_path)

        # Boutons
        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=12)
        mk_btn(bf, "✓ Générer avec ces photos",
               self._valider, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "Annuler", self._annuler, color="#888").pack(side="left", padx=6)

    _PREFS_FILE = ".photos_order.json"

    def _load_saved_order(self):
        """Restaure l'ordre et les cases cochées depuis le fichier de préférences du bon."""
        import json as _json
        pref_file = self.dossier / self._PREFS_FILE
        if not pref_file.is_file():
            return
        try:
            with open(pref_file, encoding="utf-8") as f:
                data = _json.load(f)
        except Exception:
            return
        saved_order = data.get("order", [])
        saved_checked = data.get("checked", {})
        current_by_name = {p.name: p for p in self._order}
        reordered = []
        for name in saved_order:
            if name in current_by_name:
                reordered.append(current_by_name.pop(name))
        # Nouvelles images absentes de la sauvegarde ajoutées à la fin
        for p in self._order:
            if p.name in current_by_name:
                reordered.append(p)
        self._order = reordered
        for p in self._order:
            checked = saved_checked.get(p.name, True)
            self._vars[str(p)] = tk.IntVar(value=1 if checked else 0)

    def _save_order(self):
        """Sauvegarde l'ordre et les cases cochées dans un fichier caché du dossier bon."""
        import json as _json, platform as _platform
        pref_file = self.dossier / self._PREFS_FILE
        data = {
            "order": [p.name for p in self._order],
            "checked": {
                p.name: bool(self._vars.get(str(p), tk.IntVar(value=1)).get())
                for p in self._order
            },
        }
        try:
            with open(pref_file, "w", encoding="utf-8") as f:
                _json.dump(data, f, ensure_ascii=False, indent=2)
            # Marquer le fichier comme caché sur Windows
            if _platform.system() == "Windows":
                import ctypes
                FILE_ATTRIBUTE_HIDDEN = 0x02
                ctypes.windll.kernel32.SetFileAttributesW(str(pref_file),
                                                          FILE_ATTRIBUTE_HIDDEN)
        except Exception:
            pass

    def _add_photo_row(self, img_path):
        key = str(img_path)
        rf = tk.Frame(self._inner, bg=C["surface"], relief="solid", bd=1)
        rf.pack(fill="x", pady=3, padx=2)
        self._row_frames[key] = rf

        # Poignée de glissement
        handle = tk.Label(rf, text="⠿", font=("Arial", 14), bg=C["surface"],
                          fg="#aaa", cursor="fleur", padx=4)
        handle.pack(side="left", padx=2)
        handle.bind("<Button-1>", lambda e, k=key: self._drag_start(e, k))
        handle.bind("<B1-Motion>", self._drag_motion)
        handle.bind("<ButtonRelease-1>", self._drag_end)

        # Case à cocher (conserver l'état si la ligne est recréée)
        if key not in self._vars:
            self._vars[key] = tk.IntVar(value=1)
        tk.Checkbutton(rf, variable=self._vars[key],
                       bg=C["surface"]).pack(side="left", padx=6)

        # Miniature (avec cache pour éviter de rouvrir le fichier à chaque tri)
        if key not in self._thumb_cache:
            self._thumb_cache[key] = self._make_thumb(img_path)
        thumb = self._thumb_cache[key]
        if thumb:
            tk.Label(rf, image=thumb, bg=C["surface"]).pack(side="left", padx=8, pady=6)
            if thumb not in self._thumbs:
                self._thumbs.append(thumb)
        else:
            tk.Label(rf, text="🖼", font=("Arial", 48),
                     bg=C["surface"]).pack(side="left", padx=14)

        # Nom du fichier
        tk.Label(rf, text=img_path.name, bg=C["surface"],
                 font=("Arial", 10), anchor="w").pack(side="left", padx=8)

    def _drag_start(self, event, key):
        self._drag_src = key
        rf = self._row_frames.get(key)
        if rf:
            rf.configure(bg="#cce5ff")
            for w in rf.winfo_children():
                try: w.configure(bg="#cce5ff")
                except Exception: pass

    def _drag_motion(self, event):
        if not self._drag_src:
            return
        abs_y = event.widget.winfo_rooty() + event.y
        target_key = None
        for k, rf in self._row_frames.items():
            ry = rf.winfo_rooty()
            if ry <= abs_y <= ry + rf.winfo_height():
                target_key = k
                break
        if target_key and target_key != self._drag_src:
            src_idx = next((i for i, p in enumerate(self._order)
                            if str(p) == self._drag_src), None)
            tgt_idx = next((i for i, p in enumerate(self._order)
                            if str(p) == target_key), None)
            if src_idx is not None and tgt_idx is not None:
                item = self._order.pop(src_idx)
                self._order.insert(tgt_idx, item)
                self._rebuild_pack_order()

    def _drag_end(self, event):
        src = self._drag_src
        self._drag_src = None
        if src and src in self._row_frames:
            self._row_frames[src].configure(bg=C["surface"])
            for w in self._row_frames[src].winfo_children():
                try: w.configure(bg=C["surface"])
                except Exception: pass

    def _rebuild_pack_order(self):
        """Réordonner les frames via pack_forget/pack sans recréer les widgets."""
        for img_path in self._order:
            rf = self._row_frames.get(str(img_path))
            if rf:
                rf.pack_forget()
        for img_path in self._order:
            rf = self._row_frames.get(str(img_path))
            if rf:
                rf.pack(fill="x", pady=3, padx=2)

    def _make_thumb(self, img_path, size=160):
        """Tente de creer une miniature. Necessite Pillow ; sinon None."""
        try:
            from PIL import Image, ImageTk
            im = Image.open(img_path)
            im.thumbnail((size, size))
            return ImageTk.PhotoImage(im)
        except Exception:
            # Fallback : tk.PhotoImage (PNG/GIF only, pas de resize)
            try:
                if img_path.suffix.lower() in (".png", ".gif"):
                    ph = tk.PhotoImage(file=str(img_path))
                    # Sous-echantillonnage grossier si trop grand
                    f = max(1, ph.width() // size)
                    return ph.subsample(f, f)
            except Exception:
                pass
        return None

    def _toggle_all(self, value):
        for v in self._vars.values():
            v.set(1 if value else 0)

    def _valider(self):
        self._save_order()
        # Retourner dans l'ordre d'affichage, seulement les photos cochées
        photos = [str(p) for p in self._order
                  if self._vars.get(str(p), tk.IntVar()).get()]
        self.destroy()
        if self.on_valider:
            self.on_valider(photos)

    def _annuler(self):
        self.destroy()
        # Ne genere rien (callback avec None pour signaler annulation)
        if self.on_valider:
            self.on_valider(None)

# ══════════════════════════════════════════════════════════════════════════════
# DIALOG BON D'INTERVENTION (formulaire complet)
# ══════════════════════════════════════════════════════════════════════════════
class BonDialog(tk.Toplevel):
    """Formulaire création / modification bon, conforme au modèle EMS."""

    MOTEUR_INFO_FIELDS = [
        ("machine",           "Machine"),
        ("type_moteur",       "Type moteur / inverseur"),
        ("date_mise_service", "Mise en service"),
        ("duree_garantie",    "Garantie (mois)"),
    ]
    
    def __init__(self, parent, app, inv_id=None, on_save=None,
                 offline_bundle=None, offline_path=None, on_offline_save=None):
        super().__init__(parent)
        self.app              = app
        self.inv_id           = inv_id
        self.on_save          = on_save
        self._offline_bundle  = offline_bundle
        self._offline_path    = offline_path
        self._on_offline_save = on_offline_save
        self.is_offline       = offline_bundle is not None
        self.is_edit          = (inv_id is not None) or self.is_offline

        if self.is_offline:
            self.title("Modifier bon — Mode hors-ligne")
        elif self.is_edit:
            self.title("Modifier bon")
        else:
            self.title("Nouveau bon d'intervention")

        self.geometry("980x830")
        try:
            self.state("zoomed")     # plein écran Windows
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)   # Linux
            except tk.TclError:
                pass
        self.configure(bg=C["bg"])
        self.grab_set()

        if self.is_offline:
            refs = offline_bundle.get("refs", {})
            self._clients     = refs.get("clients", [])
            self._all_moteurs = refs.get("moteurs", [])
            self._techniciens = refs.get("techniciens", [])
            self._contacts    = []
        else:
            self._clients     = list(db.get_clients())
            self._all_moteurs = list(db.get_moteurs())
            self._techniciens = list(db.get_techniciens())
            try:
                self._contacts = list(db.get_contacts())
            except Exception:
                self._contacts = []

        self._client_by_id  = {c["id"]:  c for c in self._clients}
        self._client_by_nom = {c["nom"]: c for c in self._clients}
        self._moteur_by_id  = {m["id"]:  m for m in self._all_moteurs}
        self._moteur_by_ns  = {m["num_serie"]: m for m in self._all_moteurs}
        self._contact_by_nom = {c["nom"]: c for c in self._contacts}

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
        # Bannière hors-ligne
        if self.is_offline:
            banner = tk.Frame(self, bg="#4a2800")
            banner.pack(fill="x")
            tk.Label(banner,
                     text="⚠  MODE HORS-LIGNE — Les modifications seront enregistrées "
                          "localement dans le fichier .ems. Importez en ligne pour synchroniser.",
                     font=("Arial", 9, "bold"), bg="#4a2800", fg="#ffcc88",
                     anchor="w").pack(padx=12, pady=5)

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
        # Bind permanent de la molette sur toute la dialog : insensible au
        # focus interne (combobox, chips, etc.) qui ne perturbent plus le scroll
        self.bind("<MouseWheel>", _on_wheel)
        # bind_class garantit que tout widget enfant relaie la molette
        # Binder la molette sur les widgets simples (pas les comboboxes ni Text :
        # ces derniers gerent leur propre scroll quand un dropdown est ouvert)
        for cls in ("TFrame", "Frame", "TLabel", "Label",
                    "TCheckbutton", "Checkbutton",
                    "TButton", "Button"):
            self.bind_class(cls, "<MouseWheel>", _on_wheel, add="+")

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
        if not self.is_offline:
            mk_btn(row, "+", lambda: ClientDialog(self, self.app, on_save=self._reload_refs),
                   color=C["btn2"]).pack(side="left", padx=(8,0))

        self.lieu_var       = tk.StringVar()
        self.signataire_var = tk.StringVar()
        self.email_sig_var  = tk.StringVar()
        self.tel_sig_var    = tk.StringVar()
        self.demandeur_var      = tk.StringVar()
        self.email_demand_var   = tk.StringVar()
        self.tel_demand_var     = tk.StringVar()

        contact_noms = [] if self.is_offline else [c["nom"] for c in self._contacts]

        # Lieu de l'intervention
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=2)
        tk.Label(r, text="Lieu de l'intervention", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        ttk.Entry(r, textvariable=self.lieu_var, width=46).pack(side="left")

        # Nom du signataire — SearchableCombobox alimenté par les contacts connus
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=2)
        tk.Label(r, text="Nom du signataire", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.signataire_combo = SearchableCombobox(r, textvariable=self.signataire_var,
                                                   values=contact_noms, width=44,
                                                   allow_free_text=True)
        self.signataire_combo.pack(side="left")
        self.signataire_combo.bind("<<ComboboxSelected>>", self._on_signataire_selected)

        for lbl, var, cls in [("Courriel du signataire", self.email_sig_var,  EmailEntry),
                               ("Telephone signataire",   self.tel_sig_var,    ttk.Entry)]:
            r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=2)
            tk.Label(r, text=lbl, bg=C["bg"], font=("Arial",10),
                     width=24, anchor="e").pack(side="left", padx=(0,8))
            cls(r, textvariable=var, width=46).pack(side="left")

        # Separateur visuel + demandeur (personne qui a appele pour declencher l'intervention)
        tk.Label(p, text="  Demandeur (personne ayant appele)",
                 bg=C["bg"], font=("Arial", 9, "italic"),
                 fg="#6b7785").pack(anchor="w", padx=20, pady=(8, 2))

        # Nom du demandeur — SearchableCombobox
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=2)
        tk.Label(r, text="Nom du demandeur", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.demandeur_combo = SearchableCombobox(r, textvariable=self.demandeur_var,
                                                  values=contact_noms, width=44,
                                                  allow_free_text=True)
        self.demandeur_combo.pack(side="left")
        self.demandeur_combo.bind("<<ComboboxSelected>>", self._on_demandeur_selected)

        for lbl, var, cls in [("Courriel du demandeur",   self.email_demand_var, EmailEntry),
                               ("Telephone du demandeur",  self.tel_demand_var,   ttk.Entry)]:
            r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=2)
            tk.Label(r, text=lbl, bg=C["bg"], font=("Arial",10),
                     width=24, anchor="e").pack(side="left", padx=(0,8))
            cls(r, textvariable=var, width=46).pack(side="left")

        # N° de commande client
        r = tk.Frame(p, bg=C["bg"])
        r.pack(fill="x", padx=20, pady=2)
        tk.Label(r, text="N° commande client", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.num_cmd_var = tk.StringVar()
        ttk.Entry(r, textvariable=self.num_cmd_var, width=46).pack(side="left")

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
        if not self.is_offline:
            mk_btn(row2, "+", lambda: MoteurDialog(self, self.app, on_save=self._reload_refs),
                   color=C["btn2"]).pack(side="left", padx=(8,0))

        # Navire/Site — combobox interactive (filtre les N° Série disponibles)
        r_nav = tk.Frame(p, bg=C["bg"])
        r_nav.pack(fill="x", padx=20, pady=2)
        tk.Label(r_nav, text="Navire / Site", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.navire_var = tk.StringVar()
        _all_navires = sorted({row_get(m, "navire", "") for m in self._all_moteurs
                                if row_get(m, "navire", "")})
        self.navire_combo = SearchableCombobox(r_nav, textvariable=self.navire_var,
            values=_all_navires, width=44)
        self.navire_combo.pack(side="left")
        self.navire_combo.bind("<<ComboboxSelected>>", self._on_navire_selected)

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
        ttk.Entry(rh, textvariable=self.nb_heures_var, width=20).pack(side="left")

        rg = tk.Frame(p, bg=C["bg"]); rg.pack(fill="x", padx=20, pady=2)
        self._garantie_info_row = rg
        tk.Label(rg, text="Garantie constructeur", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.garantie_lbl = tk.Label(rg, text="—", bg="#eef2f7", font=("Arial",9),
                                      anchor="w", relief="groove", padx=6,
                                      justify="left")
        self.garantie_lbl.pack(side="left", fill="x", expand=True, ipady=3)

        # Lier l'intervention à une garantie en cours (affiché dynamiquement)
        self.garantie_lien_var = tk.StringVar()
        self._garanties_ouvertes = []
        self._garanties_by_num = {}
        self._all_garanties_moteur = []
        self._garantie_lien_frame = tk.Frame(p, bg=C["bg"])
        tk.Label(self._garantie_lien_frame, text="Demande de garantie", bg=C["bg"],
                 font=("Arial", 10), width=24, anchor="e").pack(side="left", padx=(0, 8))
        self._garantie_lien_combo = ttk.Combobox(
            self._garantie_lien_frame, textvariable=self.garantie_lien_var,
            values=[], width=44, state="readonly")
        self._garantie_lien_combo.pack(side="left")
        self._garantie_lien_combo.bind("<<ComboboxSelected>>",
                                       self._on_garantie_lien_selected)

        # ── MOTEURS SUPPLÉMENTAIRES ───────────────────────────────────────────
        self._extra_moteurs_rows = []
        self._extra_moteurs_container = tk.Frame(p, bg=C["bg"])
        # Container démarré caché — affiché/masqué dynamiquement

        self._add_moteur_lnk = tk.Label(p, text="＋ ajouter un moteur", bg=C["bg"],
                                         fg="#6b7785", font=("Segoe UI", 9, "italic"),
                                         cursor="hand2")
        self._add_moteur_lnk.pack(anchor="w", padx=20, pady=(1, 6))
        self._add_moteur_lnk.bind("<Button-1>", lambda _: self._add_moteur_row())

        # ── DÉTAILS INTERVENTION ──────────────────────────────────────────────
        section_bar(p, "DÉTAILS INTERVENTION")
        self.type_var    = tk.StringVar()
        self.date_var    = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.statut_var  = tk.StringVar(value="En cours")
        self.urgence_var = tk.StringVar(value="Normale")

        # Type d'intervention
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=4)
        tk.Label(r, text="Type d'intervention *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        if self.is_offline:
            _types = self._offline_bundle.get("refs", {}).get("types", [])
        else:
            try:
                _types = db.get_types_intervention()
            except Exception:
                _types = []
        self._type_combo = ttk.Combobox(r, textvariable=self.type_var,
                                        values=_types,
                                        width=44, state="readonly")
        self._type_combo.pack(side="left")
        if not self.is_offline:
            mk_btn(r, "⚙", lambda: TypesDialog(self, self.app,
                   on_close=lambda: self._type_combo.config(
                       values=db.get_types_intervention())),
                   color=C["btn2"]).pack(side="left", padx=(8,0))

        self.marque_var = tk.StringVar()

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
        tk.Checkbutton(cls_frame, text="Facturable", variable=self.chk["facturable"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 12))
        tk.Checkbutton(cls_frame, text="Garantie", variable=self.chk["garantie_intervention"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left", padx=(0, 12))
        tk.Checkbutton(cls_frame, text="Interne", variable=self.chk["interne"],
                       bg=C["bg"], font=("Arial",10)).pack(side="left")

        # Technicien(s) — un ou plusieurs
        r = tk.Frame(p, bg=C["bg"]); r.pack(fill="x", padx=20, pady=(4, 8))
        tk.Label(r, text="Technicien(s) *", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="ne").pack(side="left", padx=(0,8), pady=(2, 0))
        # Le picker prend la place et gère ses propres boutons + chips
        def _open_new_tech(then_callback):
            """Ouvre TechnicienDialog ; à la sauvegarde, callback avec le nom créé."""
            def on_save_new(new_name=None):
                self._reload_refs()
                # Met à jour le picker et auto-ajoute le tech créé
                then_callback(new_name)
            TechnicienDialog(self, self.app, on_save=on_save_new)
        self.tech_picker = TechniciensPicker(r,
            available=[t["nom"] for t in self._techniciens],
            on_add_new=None if self.is_offline else _open_new_tech)
        self.tech_picker.pack(side="left", fill="x", expand=True)

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
        tk.Label(op2, text="Photos :", bg=C["bg"],
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

        # ── POUR INFORMATION (encart dédié) ───────────────────────────────
        info_f = tk.Frame(p, bg=C["bg"])
        info_f.pack(fill="x", padx=20, pady=(8, 2))
        tk.Checkbutton(info_f, text="Pour information",
                       variable=self.chk["pour_information"],
                       bg=C["bg"], font=("Arial", 10, "bold")
                       ).pack(side="left")
        self.txt_info = tk.Text(p, height=3, font=("Arial", 10), wrap="word",
                                 relief="solid", bd=1, padx=6, pady=4)
        self.txt_info.pack(fill="x", padx=20, pady=(2, 6))

        # ── PRÉCONISATION (encart dédié) ──────────────────────────────────
        preco_f = tk.Frame(p, bg=C["bg"])
        preco_f.pack(fill="x", padx=20, pady=(8, 2))
        tk.Checkbutton(preco_f, text="Préconisation",
                       variable=self.chk["preconisation"],
                       bg=C["bg"], font=("Arial", 10, "bold")
                       ).pack(side="left")
        self.txt_preco = tk.Text(p, height=3, font=("Arial", 10), wrap="word",
                                  relief="solid", bd=1, padx=6, pady=4)
        self.txt_preco.pack(fill="x", padx=20, pady=(2, 6))

        # ── DÉPLACEMENTS ──────────────────────────────────────────────────────
        section_bar(p, "DÉPLACEMENTS – Temps & frais")
        self.depl_tbl = DeplacementsTable(p)
        self.depl_tbl.pack(fill="x", padx=20, pady=(4, 8))
        # Sync auto : quand un tech est ajouté/retiré du picker, l'ajouter à tous les jours
        self.tech_picker._on_change_cb = lambda names: self.depl_tbl.sync_techniciens(names)

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

        # ── DATES & TRAÇABILITÉ ────────────────────────────────────────────────
        section_bar(p, "DATES & TRAÇABILITÉ")
        dates_f = tk.Frame(p, bg=C["bg"])
        dates_f.pack(fill="x", padx=20, pady=(4, 8))

        # 1ère ligne : Date d'ouverture (auto) + Date de clôture (éditable)
        rd1 = tk.Frame(dates_f, bg=C["bg"]); rd1.pack(fill="x", pady=2)
        tk.Label(rd1, text="Date d'ouverture", bg=C["bg"], font=("Arial",10),
                 width=22, anchor="e").pack(side="left", padx=(0,8))
        self.date_ouv_lbl = tk.Label(rd1, text="—", bg="#eef2f7",
                                       font=("Arial",10), fg="#333",
                                       relief="groove", padx=8, width=22,
                                       anchor="w")
        self.date_ouv_lbl.pack(side="left", ipady=2)

        tk.Label(rd1, text="   Date de clôture", bg=C["bg"], font=("Arial",10),
                 width=18, anchor="e").pack(side="left", padx=(12,8))
        self.date_clot_var = tk.StringVar()
        DateEntry(rd1, textvariable=self.date_clot_var, width=14).pack(side="left")
        tk.Label(rd1, text=" (JJ/MM/AAAA)", bg=C["bg"], fg="#888",
                 font=("Arial",8)).pack(side="left", padx=(4,0))

        # 2ème ligne : Signature client + Signature technicien
        rd2 = tk.Frame(dates_f, bg=C["bg"]); rd2.pack(fill="x", pady=2)
        tk.Label(rd2, text="Signature client", bg=C["bg"], font=("Arial",10),
                 width=22, anchor="e").pack(side="left", padx=(0,8))
        self.date_sig_cli_lbl = tk.Label(rd2, text="—", bg="#eef2f7",
                                          font=("Arial",10), fg="#333",
                                          relief="groove", padx=8, width=22,
                                          anchor="w")
        self.date_sig_cli_lbl.pack(side="left", ipady=2)

        tk.Label(rd2, text="   Signature tech.", bg=C["bg"], font=("Arial",10),
                 width=18, anchor="e").pack(side="left", padx=(12,8))
        self.date_sig_tech_lbl = tk.Label(rd2, text="—", bg="#eef2f7",
                                           font=("Arial",10), fg="#333",
                                           relief="groove", padx=8, width=22,
                                           anchor="w")
        self.date_sig_tech_lbl.pack(side="left", ipady=2)

        # ── COMMENTAIRE ───────────────────────────────────────────────────────
        section_bar(p, "COMMENTAIRE")
        self.txt_commentaire = tk.Text(p, height=4, font=("Arial",10), wrap="word",
                                        relief="solid", bd=1, padx=6, pady=4)
        self.txt_commentaire.pack(fill="x", padx=20, pady=(4, 8))

        # ── BOUTONS ───────────────────────────────────────────────────────────
        bf = tk.Frame(p, bg=C["bg"])
        bf.pack(pady=14)
        mk_btn(bf, "💾 Enregistrer",            lambda: self._save(generer=False)).pack(side="left", padx=6)
        mk_btn(bf, "📄 Enregistrer + Générer",  lambda: self._save(generer=True)).pack(side="left", padx=6)
           
        mk_btn(bf, "✍ Faire signer",           self._faire_signer, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "📧 Prévenir client",        self._mail_client_inline,  color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "📧 Prévenir technicien",    self._mail_tech_inline,    color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "✅ Mail clôture",            self._mail_cloture_inline, color=C["btn2"]).pack(side="left", padx=6)
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
        # Le picker conserve sa sélection, on recharge juste la liste dispo
        old_techs  = self.tech_picker.get_names()

        if self.is_offline:
            refs = self._offline_bundle.get("refs", {})
            self._clients     = refs.get("clients", [])
            self._all_moteurs = refs.get("moteurs", [])
            self._techniciens = refs.get("techniciens", [])
            self._contacts    = []
        else:
            self._clients     = list(db.get_clients())
            self._all_moteurs = list(db.get_moteurs())
            self._techniciens = list(db.get_techniciens())
            try:
                self._contacts = list(db.get_contacts())
            except Exception:
                self._contacts = []
        self._client_by_id  = {c["id"]:  c for c in self._clients}
        self._client_by_nom = {c["nom"]: c for c in self._clients}
        self._moteur_by_id  = {m["id"]:  m for m in self._all_moteurs}
        self._moteur_by_ns  = {m["num_serie"]: m for m in self._all_moteurs}

        self.client_combo.set_values([c["nom"] for c in self._clients])
        self.tech_picker.set_available([t["nom"] for t in self._techniciens])
        contact_noms = [c["nom"] for c in self._contacts]
        self._contact_by_nom = {c["nom"]: c for c in self._contacts}
        self.signataire_combo.set_values(contact_noms)
        self.demandeur_combo.set_values(contact_noms)
        cli = self._client_by_nom.get(old_client)
        filtrés = ([m for m in self._all_moteurs if m["client_id"] == cli["id"]]
                   if cli else self._all_moteurs)
        self.moteur_combo.set_values([m["num_serie"] for m in filtrés])
        self.client_var.set(old_client)
        self.moteur_var.set(old_moteur)
        self.tech_picker.set_names(old_techs)

    # ─── Événements ───────────────────────────────────────────────────────────
    def _moteurs_filtres(self):
        """Retourne la liste des moteurs filtrée selon le client sélectionné."""
        client = self._client_by_nom.get(self.client_var.get())
        if client:
            return [m for m in self._all_moteurs if m["client_id"] == client["id"]]
        return self._all_moteurs

    def _on_client_selected(self, *_):
        client = self._client_by_nom.get(self.client_var.get())
        if client:
            if not self.signataire_var.get():
                self.signataire_combo.set(row_get(client, "contact"))
            if not self.email_sig_var.get():
                self.email_sig_var.set(row_get(client, "email"))
        filtres = self._moteurs_filtres()
        series = [m["num_serie"] for m in filtres]
        navires = sorted({row_get(m, "navire", "") for m in filtres if row_get(m, "navire", "")})
        self.navire_combo.set_values(navires)
        self.navire_var.set("")
        self.moteur_combo.set_values(series)
        self.moteur_var.set("")
        for lw in self._info_lbls.values():
            lw.config(text="")
        self.garantie_lbl.config(text="—", fg="black")
        # Mettre à jour les combos des moteurs supplémentaires
        for e in self._extra_moteurs_rows:
            e["combo"].set_values(series)
            e["var"].set("")
            for lw in e["info_lbls"].values():
                lw.config(text="")
            if "gar_lbl" in e:
                e["gar_lbl"].config(text="—", fg="black")

    def _on_signataire_selected(self, *_):
        c = self._contact_by_nom.get(self.signataire_var.get())
        if c:
            if not self.email_sig_var.get():
                self.email_sig_var.set(c.get("email", ""))
            if not self.tel_sig_var.get():
                self.tel_sig_var.set(c.get("telephone", ""))

    def _on_demandeur_selected(self, *_):
        c = self._contact_by_nom.get(self.demandeur_var.get())
        if c:
            if not self.email_demand_var.get():
                self.email_demand_var.set(c.get("email", ""))
            if not self.tel_demand_var.get():
                self.tel_demand_var.set(c.get("telephone", ""))

    def _on_navire_selected(self, *_):
        navire = self.navire_var.get()
        if not navire:
            return
        filtres = [m for m in self._moteurs_filtres() if row_get(m, "navire", "") == navire]
        series = [m["num_serie"] for m in filtres]
        self.moteur_combo.set_values(series)
        if len(filtres) == 1:
            self.moteur_var.set(filtres[0]["num_serie"])
            self._on_moteur_selected()
        else:
            self.moteur_var.set("")
            for lw in self._info_lbls.values():
                lw.config(text="")
            self.garantie_lbl.config(text="—", fg="black")

    def _on_moteur_selected(self, *_):
        m = self._moteur_by_ns.get(self.moteur_var.get())
        if m:
            self._set_moteur_info(m)
            self.marque_var.set(row_get(m, "marque", ""))
            if not self.client_var.get():
                c = self._client_by_id.get(m["client_id"])
                if c:
                    self.client_combo.set(c["nom"])
                    # Mettre à jour le navire combo selon les moteurs du client
                    client_moteurs = [mot for mot in self._all_moteurs
                                      if mot["client_id"] == c["id"]]
                    navires = sorted({row_get(mot, "navire", "") for mot in client_moteurs
                                      if row_get(mot, "navire", "")})
                    self.navire_combo.set_values(navires)
                    if not self.signataire_var.get():
                        self.signataire_combo.set(row_get(c, "contact"))
                    if not self.email_sig_var.get():
                        self.email_sig_var.set(row_get(c, "email"))

    def _set_moteur_info(self, m):
        self.navire_combo.set(str(row_get(m, "navire", "")))
        for key, lw in self._info_lbls.items():
            lw.config(text=str(row_get(m, key, "")))

        # ── Statut garantie constructeur (date_mise_service + duree_garantie) ──
        gstat, gjours = db.garantie_status(
            row_get(m, "date_mise_service", ""),
            row_get(m, "duree_garantie", ""))
        if gstat == "Active":
            self.garantie_lbl.config(
                text=f"✅ Garantie constructeur ACTIVE — {gjours} jour(s) restant(s)",
                fg="#0f5132")
        elif gstat == "Expiree":
            self.garantie_lbl.config(
                text=f"⛔ Garantie constructeur expirée (il y a {gjours} jour(s))",
                fg="#b91c1c")
        else:
            self.garantie_lbl.config(
                text="—  (aucune garantie constructeur renseignée)",
                fg="#666")

        # ── Demandes de garantie (dossiers app garanties) ────────────────────
        if self.is_offline:
            gars = []
        else:
            try:
                gars = db.get_garanties_moteur(row_get(m, "id"))
            except Exception:
                gars = []

        self._all_garanties_moteur = gars
        ouvertes = [g for g in gars if g["statut"] != "Clôturée"]
        self._garanties_ouvertes = ouvertes
        self._garanties_by_num = {g["num_ems"]: g for g in ouvertes}

        if ouvertes:
            labels = ["— Aucune (ne pas lier)"] + [
                f"{g['num_ems']} · {g['attribution']} · {g['statut']}"
                for g in ouvertes
            ]
            self._garantie_lien_combo.config(values=labels)
            cur = self.garantie_lien_var.get()
            if not cur or cur not in labels:
                self.garantie_lien_var.set("— Aucune (ne pas lier)")
            self._garantie_lien_frame.pack(fill="x", padx=20, pady=2,
                                           after=self._garantie_info_row)
        else:
            self._garantie_lien_frame.pack_forget()
            self.garantie_lien_var.set("")

    def _on_garantie_lien_selected(self, *_):
        sel = self.garantie_lien_var.get()
        if sel and sel != "— Aucune (ne pas lier)":
            self.chk["garantie_intervention"].set(1)
            num_ems = sel.split(" · ")[0] if " · " in sel else sel
            g = self._garanties_by_num.get(num_ems)
            if g and not self.demandeur_var.get():
                nom = row_get(g, "nom_demandeur")
                if nom:
                    self.demandeur_combo.set(nom)
                    if not self.email_demand_var.get():
                        self.email_demand_var.set(row_get(g, "email_demandeur"))
                    if not self.tel_demand_var.get():
                        self.tel_demand_var.set(row_get(g, "telephone_demandeur"))

    # ─── Moteurs supplémentaires ─────────────────────────────────────────────
    def _add_moteur_row(self, initial_ns="", initial_nb_heures=""):
        # Afficher le container la première fois (avant le lien)
        if not self._extra_moteurs_rows:
            self._extra_moteurs_container.pack(fill="x", padx=20, pady=(2, 0),
                                               before=self._add_moteur_lnk)

        idx = len(self._extra_moteurs_rows) + 2

        # Conteneur global du bloc moteur supplémentaire
        bloc = tk.Frame(self._extra_moteurs_container, bg=C["bg"],
                        bd=1, relief="groove", padx=6, pady=4)
        bloc.pack(fill="x", pady=(4, 0))

        # ── ligne titre + N° série + × ──
        row_top = tk.Frame(bloc, bg=C["bg"])
        row_top.pack(fill="x")
        tk.Label(row_top, text=f"Moteur {idx}  –  N° Série",
                 bg=C["bg"], font=("Arial", 10, "bold"),
                 width=24, anchor="e").pack(side="left", padx=(0, 8))

        var = tk.StringVar(value=initial_ns)
        combo = SearchableCombobox(row_top, textvariable=var,
                                   values=[m["num_serie"] for m in self._moteurs_filtres()],
                                   width=38)
        combo.pack(side="left")

        entry = {"var": var, "bloc": bloc, "combo": combo, "info_lbls": {}}
        self._extra_moteurs_rows.append(entry)

        def _remove(e=entry):
            e["bloc"].pack_forget()
            e["bloc"].destroy()
            if e in self._extra_moteurs_rows:
                self._extra_moteurs_rows.remove(e)
            # Masquer le container quand il ne reste plus aucun moteur
            if not self._extra_moteurs_rows:
                self._extra_moteurs_container.pack_forget()

        btn_x = tk.Label(row_top, text=" ×", bg=C["bg"], fg="#c62828",
                         font=("Arial", 13, "bold"), cursor="hand2")
        btn_x.pack(side="left", padx=(6, 0))
        btn_x.bind("<Button-1>", lambda _e, r=_remove: r())

        # ── lignes d'info (identiques au moteur principal) ──
        INFO_FIELDS = [
            ("navire",            "Navire / Site"),
            ("machine",           "Machine"),
            ("type_moteur",       "Type moteur / inverseur"),
            ("date_mise_service", "Mise en service"),
            ("duree_garantie",    "Garantie (mois)"),
        ]
        for key, lbl_txt in INFO_FIELDS:
            r = tk.Frame(bloc, bg=C["bg"])
            r.pack(fill="x", pady=1)
            tk.Label(r, text=lbl_txt, bg=C["bg"], font=("Arial", 9),
                     width=24, anchor="e").pack(side="left", padx=(0, 8))
            lw = tk.Label(r, text="", bg="#eef2f7", font=("Arial", 9),
                          anchor="w", relief="groove", width=46, padx=4)
            lw.pack(side="left", ipady=2)
            entry["info_lbls"][key] = lw

        # Nb heures de fonctionnement (saisie manuelle)
        r_h = tk.Frame(bloc, bg=C["bg"])
        r_h.pack(fill="x", pady=1)
        tk.Label(r_h, text="Nb heures de fonctionnement", bg=C["bg"], font=("Arial", 9),
                 width=24, anchor="e").pack(side="left", padx=(0, 8))
        nb_h_var = tk.StringVar(value=initial_nb_heures)
        ttk.Entry(r_h, textvariable=nb_h_var, width=20).pack(side="left")
        entry["nb_heures_var"] = nb_h_var

        # Garanties du moteur
        r_g = tk.Frame(bloc, bg=C["bg"])
        r_g.pack(fill="x", pady=1)
        tk.Label(r_g, text="Garanties du moteur", bg=C["bg"], font=("Arial", 9),
                 width=24, anchor="e").pack(side="left", padx=(0, 8))
        gar_lbl = tk.Label(r_g, text="—", bg="#eef2f7", font=("Arial", 9),
                           anchor="w", relief="groove", padx=6, justify="left")
        gar_lbl.pack(side="left", fill="x", expand=True, ipady=2)
        entry["gar_lbl"] = gar_lbl

        # Peuplement auto des infos à la sélection
        def _on_sel(*_, e=entry):
            m = self._moteur_by_ns.get(e["var"].get().strip())
            for key, lw in e["info_lbls"].items():
                lw.config(text=str(row_get(m, key, "") if m else ""))
            # Garanties
            lbl = e["gar_lbl"]
            if not m:
                lbl.config(text="—", fg="black")
                return
            if self.is_offline:
                gars = []
            else:
                try:
                    gars = db.get_garanties_moteur(row_get(m, "id"))
                except Exception:
                    gars = []
            if not gars:
                lbl.config(text="—  (aucune garantie enregistrée pour ce moteur)", fg="#666")
            else:
                ouvertes = [g for g in gars if g["statut"] != "Clôturée"]
                if ouvertes:
                    parts = [f"{g['num_ems']} · {g['attribution']} · {g['statut']}"
                             for g in ouvertes[:3]]
                    txt = "🛡 " + "   |   ".join(parts)
                    if len(ouvertes) > 3:
                        txt += f"   (+{len(ouvertes) - 3} autre(s))"
                    lbl.config(text=txt, fg="#0f5132")
                else:
                    lbl.config(text=f"🛡 {len(gars)} garantie(s) — toutes clôturées", fg="#666")

        combo.bind("<<ComboboxSelected>>", _on_sel)

        # Si initial_ns fourni (chargement), peupler immédiatement
        if initial_ns:
            bloc.after(50, _on_sel)

    # ─── Fichiers ─────────────────────────────────────────────────────────────
    def _current_dossier(self):
        if self.is_offline or not self.is_edit or not self.inv_id:
            return None
        inv = db.get_intervention(inv_id=self.inv_id)
        if not inv: return None
        return _dossiers_root() / inv["num_bon"]

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
            title="Sélectionner fichier(s) ou photo(s) (Ctrl+clic pour multi-sélection)",
            filetypes=[("Tous", "*.*"),
                       ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.heic"),
                       ("PDF",    "*.pdf"),
                       ("Documents", "*.doc *.docx *.xls *.xlsx *.txt")])
        if not files: return
        d.mkdir(parents=True, exist_ok=True)

        # Tentative d'import Pillow pour la compression
        try:
            from PIL import Image
            pillow_ok = True
        except ImportError:
            pillow_ok = False

        IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".heic"}
        MAX_DIM = 1920          # px max sur la plus grande dimension
        JPEG_QUALITY = 85       # qualite compression JPEG
        SEUIL_COMPRESS = 500_000  # compresser uniquement > 500 Ko

        n_total = 0
        n_compressed = 0
        gain_total = 0
        for f in files:
            src = Path(f)
            try:
                dst = d / src.name
                # Compression si image, Pillow dispo, et > seuil
                if (pillow_ok and src.suffix.lower() in IMAGE_EXTS
                        and src.stat().st_size > SEUIL_COMPRESS):
                    try:
                        img = Image.open(src)
                        # Convertir RGBA -> RGB pour JPEG
                        if img.mode in ("RGBA", "LA", "P"):
                            bg = Image.new("RGB", img.size, (255, 255, 255))
                            bg.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                            img = bg
                        # Redimensionner si trop grand
                        if max(img.size) > MAX_DIM:
                            ratio = MAX_DIM / max(img.size)
                            new_size = (int(img.size[0] * ratio),
                                        int(img.size[1] * ratio))
                            img = img.resize(new_size, Image.LANCZOS)
                        # Toujours en JPEG (sauf si dest .png on garde png)
                        if src.suffix.lower() == ".png":
                            dst = d / (src.stem + ".jpg")  # convertir PNG > JPG
                        else:
                            dst = d / src.name
                        img.save(dst, "JPEG", quality=JPEG_QUALITY, optimize=True)
                        gain_total += src.stat().st_size - dst.stat().st_size
                        n_compressed += 1
                        n_total += 1
                        continue
                    except Exception:
                        # Si la compression echoue, on retombe sur copie brute
                        pass
                # Copie brute (non-image, ou Pillow absent, ou < seuil)
                shutil.copy2(src, dst)
                n_total += 1
            except (OSError, shutil.SameFileError) as e:
                messagebox.showwarning("Copie", f"Erreur sur {src.name} :\n{e}")

        if n_total:
            msg = f"{n_total} fichier(s) ajouté(s) dans {d.name}."
            if n_compressed:
                gain_mo = gain_total / (1024 * 1024)
                msg += f"\n\n{n_compressed} image(s) compressée(s) — gain : {gain_mo:.1f} Mo."
            messagebox.showinfo("Import", msg)
        self._refresh_files()

    def _ouvrir_dossier(self):
        d = self._current_dossier()
        if not d:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon pour créer le dossier.")
            return
        d.mkdir(parents=True, exist_ok=True)
        ouvrir_fichier(d)

    # ─── Signature client ─────────────────────────────────────────────────────
    def _faire_signer(self):
        if self.is_offline:
            SignatureDialog(self, self.app,
                            offline_bundle=self._offline_bundle,
                            offline_path=self._offline_path,
                            on_offline_save=self._reload_apres_signature)
            return
        if not self.is_edit or not self.inv_id:
            messagebox.showinfo(
                "Enregistrer d'abord",
                "Enregistrez le bon avant de le faire signer.")
            return
        SignatureDialog(self, self.app, inv_id=self.inv_id,
                        on_save=self._reload_apres_signature)

    def _reload_apres_signature(self):
        """Apres signature, recharge les valeurs depuis l'API (le statut a
        pu basculer automatiquement en 'A facturer')."""
        try:
            self._load()
        except Exception as e:
            print(f"[BonDialog] Rechargement apres signature : {e}")

    # ─── Notifications ────────────────────────────────────────────────────────
    def _mail_client_inline(self):
        if self.is_offline:
            messagebox.showwarning("Hors ligne",
                "L'envoi d'email n'est pas disponible en mode hors-ligne.\n"
                "Importez le bon en ligne pour utiliser cette fonction.")
            return
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon avant de notifier le client.")
            return
        inv = db.get_intervention(inv_id=self.inv_id)
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        email = (row_get(inv, "email_demandeur") or
                 row_get(inv, "email_signataire") or
                 (row_get(client, "email") if client else ""))
        if not email:
            messagebox.showwarning("Email manquant",
                "Aucun email renseigné (demandeur, signataire ou client).")
            return
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_client(inv, client, moteur, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")
        db.mark_notifie(self.inv_id, "client")

    def _mail_tech_inline(self):
        if self.is_offline:
            messagebox.showwarning("Hors ligne",
                "L'envoi d'email n'est pas disponible en mode hors-ligne.\n"
                "Importez le bon en ligne pour utiliser cette fonction.")
            return
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon avant de notifier le technicien.")
            return
        inv = db.get_intervention(inv_id=self.inv_id)
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        # Collecter les emails de TOUS les techniciens assignés
        tech_names = db.parse_techniciens(row_get(inv, "technicien"))
        emails = []
        for n in tech_names:
            t = db.get_technicien_by_nom(n)
            em = row_get(t, "email")
            if em:
                emails.append(em)
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_technicien(inv, client, moteur, emails, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")
        db.mark_notifie(self.inv_id, "tech")

    def _mail_cloture_inline(self):
        if self.is_offline:
            messagebox.showwarning("Hors ligne",
                "L'envoi d'email n'est pas disponible en mode hors-ligne.\n"
                "Importez le bon en ligne pour utiliser cette fonction.")
            return
        if not self.is_edit:
            messagebox.showinfo("Sauvegarde requise",
                "Enregistrez d'abord le bon avant d'envoyer le mail de clôture.")
            return
        inv    = db.get_intervention(inv_id=self.inv_id)
        client = db.get_client(inv["client_id"]) if inv["client_id"] else None
        moteur = db.get_moteur(inv["moteur_id"]) if inv["moteur_id"] else None
        if not row_get(inv, "email_demandeur") and not row_get(inv, "email_signataire"):
            messagebox.showwarning("Destinataires manquants",
                "Aucun email demandeur ni signataire renseigné sur ce bon.")
            return
        tech_emails = []
        for n in db.parse_techniciens(row_get(inv, "technicien")):
            t = db.get_technicien_by_nom(n)
            em = row_get(t, "email")
            if em:
                tech_emails.append(em)
        num_bon = row_get(inv, "num_bon")
        _d = _dossiers_root() / num_bon
        _pdf = _d / f"{num_bon}.pdf"
        _html = _d / f"{num_bon}.html"
        if _pdf.exists():
            path = _pdf
        elif _html.exists():
            path = _html
        else:
            try:
                path = sauvegarder_bon(inv, generer_pdf=True)
            except (PermissionError, RuntimeError) as exc:
                messagebox.showerror("PDF impossible", str(exc)); return
        _, pj_auto = mailer.email_cloture(inv, client, moteur, tech_emails, str(path))
        if not pj_auto:
            messagebox.showinfo("Pièce jointe",
                f"Joignez le bon manuellement depuis le dossier qui vient de s'ouvrir :\n{Path(path).parent}")

    # ─── Chargement ───────────────────────────────────────────────────────────
    def _load(self):
        if self.is_offline:
            from shared import import_export as ie
            inv = ie.get_data(self._offline_bundle)
        else:
            inv = db.get_intervention(inv_id=self.inv_id)
            if not inv:
                return

        c = self._client_by_id.get(row_get(inv, "client_id"))
        if c:
            self.client_combo.set(c["nom"])
        self.lieu_var.set(row_get(inv, "lieu_intervention"))
        self.signataire_combo.set(row_get(inv, "nom_signataire"))
        self.email_sig_var.set(row_get(inv, "email_signataire"))
        self.tel_sig_var.set(row_get(inv, "telephone_signataire"))
        self.demandeur_combo.set(row_get(inv, "nom_demandeur"))
        self.email_demand_var.set(row_get(inv, "email_demandeur"))
        self.tel_demand_var.set(row_get(inv, "telephone_demandeur"))

        if c:
            filtrés = [m for m in self._all_moteurs if m["client_id"] == c["id"]]
        else:
            filtrés = self._all_moteurs
        navires_load = sorted({row_get(m, "navire", "") for m in filtrés if row_get(m, "navire", "")})
        self.navire_combo.set_values(navires_load)
        self.moteur_combo.set_values([m["num_serie"] for m in filtrés])
        m = self._moteur_by_id.get(row_get(inv, "moteur_id"))
        if m:
            self.moteur_combo.set(m["num_serie"])
            self._set_moteur_info(m)
            # Pré-sélectionner la garantie déjà liée à cette intervention
            if not self.is_offline and self.inv_id:
                linked = next(
                    (g for g in self._all_garanties_moteur
                     if g.get("intervention_id") == self.inv_id),
                    None
                )
                if linked:
                    lbl = f"{linked['num_ems']} · {linked['attribution']} · {linked['statut']}"
                    vals = list(self._garantie_lien_combo.cget("values"))
                    if lbl not in vals:
                        # Garantie clôturée mais liée — on l'ajoute pour affichage
                        vals = (vals if vals else ["— Aucune (ne pas lier)"]) + [lbl]
                        self._garantie_lien_combo.config(values=vals)
                        self._garanties_by_num[linked["num_ems"]] = linked
                        self._garantie_lien_frame.pack(
                            fill="x", padx=20, pady=2,
                            after=self._garantie_info_row)
                    self.garantie_lien_var.set(lbl)
        self.nb_heures_var.set(row_get(inv, "nb_heures_fct"))
        self.num_cmd_var.set(row_get(inv, "num_commande_client"))
        self.type_var.set(row_get(inv, "type_intervention"))
        self.tech_picker.set(row_get(inv, "technicien"))
        self.date_var.set(row_get(inv, "date_creation"))
        # Dates & tracabilite
        self.date_clot_var.set(row_get(inv, "date_cloture"))
        # created_at vient en ISO, on l'affiche en JJ/MM/AAAA HH:MM
        ouv = db.fmt_paris_short(row_get(inv, "created_at"))
        self.date_ouv_lbl.config(text=ouv or "—")
        sig_cli = row_get(inv, "signature_date")
        self.date_sig_cli_lbl.config(text=sig_cli or "—")
        sig_tech = row_get(inv, "signature_tech_date")
        self.date_sig_tech_lbl.config(text=sig_tech or "—")
        self.statut_var.set(row_get(inv, "statut", "En cours"))
        self.urgence_var.set(row_get(inv, "urgence", "Normale") or "Normale")
        self.marque_var.set(row_get(inv, "marque", ""))

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
        self.txt_preco.delete("1.0", "end")
        self.txt_preco.insert("1.0", row_get(inv, "preconisation_text"))
        self.txt_commentaire.delete("1.0", "end")
        self.txt_commentaire.insert("1.0", row_get(inv, "commentaire"))

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

        # Moteurs supplémentaires
        try:
            extra_ms = json.loads(row_get(inv, "moteurs_supplementaires_json", "[]") or "[]")
        except (json.JSONDecodeError, TypeError):
            extra_ms = []
        for em in extra_ms:
            self._add_moteur_row(initial_ns=em.get("num_serie", ""),
                                 initial_nb_heures=em.get("nb_heures_fct", ""))

        self._refresh_files()

    def _exporter_central(self):
        """Exporte le bon selectionne vers la base centrale (serveur atelier)."""
        r = self._sel()
        if not r:
            return
        try:
            from ems_client import sync_client
        except ImportError:
            messagebox.showerror("Module manquant",
                "Le module de synchronisation n'est pas disponible.")
            return

        if not sync_client.serveur_central_joignable():
            messagebox.showwarning("Serveur central injoignable",
                "Impossible de joindre la base centrale.\n\n"
                "Vérifiez votre connexion au réseau de l'atelier, "
                "puis réessayez. Le bon reste enregistré localement.")
            return

        res = sync_client.exporter_bon(r["id"])
        status = res["status"]
        if status == sync_client.EXPORT_OK:
            messagebox.showinfo("Export réussi", res["message"])
        elif status == sync_client.EXPORT_CONFLIT:
            if messagebox.askyesno("Conflit détecté",
                    res["message"] + "\n\n"
                    "• OUI : votre version remplacera celle du bureau\n"
                    "• NON : annuler l'export"):
                res2 = sync_client.exporter_bon(r["id"], force=True)
                if res2["status"] == sync_client.EXPORT_OK:
                    messagebox.showinfo("Export réussi (forcé)", res2["message"])
                else:
                    messagebox.showerror("Échec", res2["message"])
        elif status == sync_client.EXPORT_HORS_LIGNE:
            messagebox.showwarning("Hors ligne", res["message"])
        else:
            messagebox.showerror("Erreur d'export", res["message"])

        try:
            from ems_client import sync_client
        except ImportError:
            messagebox.showerror(
                "Module manquant",
                "Le module de synchronisation n'est pas disponible.")
            return

        # Test rapide de joignabilité (évite une longue attente)
        if not sync_client.serveur_central_joignable():
            messagebox.showwarning(
                "Serveur central injoignable",
                "Impossible de joindre la base centrale.\n\n"
                "Vérifiez que vous êtes connecté au réseau de l'atelier, "
                "puis réessayez. Le bon reste enregistré localement.")
            return

        res = sync_client.exporter_bon(self.inv_id)
        status = res["status"]

        if status == sync_client.EXPORT_OK:
            messagebox.showinfo("Export réussi", res["message"])

        elif status == sync_client.EXPORT_CONFLIT:
            # Le bon a aussi été modifié au bureau : demander confirmation
            ecraser = messagebox.askyesno(
                "Conflit détecté",
                res["message"] + "\n\n"
                "• OUI : votre version remplacera celle du bureau\n"
                "• NON : annuler l'export (rien n'est modifié)")
            if ecraser:
                res2 = sync_client.exporter_bon(self.inv_id, force=True)
                if res2["status"] == sync_client.EXPORT_OK:
                    messagebox.showinfo("Export réussi (forcé)", res2["message"])
                else:
                    messagebox.showerror("Échec", res2["message"])

        elif status == sync_client.EXPORT_HORS_LIGNE:
            messagebox.showwarning("Hors ligne", res["message"])

        else:
            messagebox.showerror("Erreur d'export", res["message"])    


    # ─── Sauvegarde ───────────────────────────────────────────────────────────
    def _save(self, generer=False):
        cn   = self.client_var.get().strip()
        ns   = self.moteur_var.get().strip()
        ti   = self.type_var.get().strip()
        tech_names = self.tech_picker.get_names()
        tech = db.format_techniciens(tech_names)
        date_str = self.date_var.get().strip()
        demande  = self.txt_demande.get("1.0", "end").strip()

        if not all([cn, ns, ti, tech, demande]):
            messagebox.showwarning("Champs manquants",
                "Client, N° Série, Type, Technicien(s) et Demande du client sont obligatoires.")
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

        materiels    = self.materiels_tbl.to_list()
        deplacements = self.depl_tbl.to_dict()

        data = {
            "client_id":         client["id"] if client else "",
            "moteur_id":         moteur["id"] if moteur else "",
            "type_intervention": ti,
            "urgence":           self.urgence_var.get(),
            "technicien":        tech,
            "date_creation":     date_str,
            "date_cloture":      self.date_clot_var.get().strip(),
            "statut":            self.statut_var.get(),

            "lieu_intervention": self.lieu_var.get().strip(),
            "nom_signataire":      self.signataire_var.get().strip(),
            "email_signataire":    self.email_sig_var.get().strip(),
            "telephone_signataire": self.tel_sig_var.get().strip(),
            "nom_demandeur":       self.demandeur_var.get().strip(),
            "email_demandeur":     self.email_demand_var.get().strip(),
            "telephone_demandeur": self.tel_demand_var.get().strip(),
            "nb_heures_fct":         self.nb_heures_var.get().strip(),
            "num_commande_client":   self.num_cmd_var.get().strip(),
            "marque":                self.marque_var.get().strip(),

            **{k: int(v.get()) for k, v in self.chk.items()},

            "demande_client":    demande,
            "constat":           self.txt_constat.get("1.0", "end").strip(),
            "travaux":           self.txt_travaux.get("1.0", "end").strip(),
            "informations":      self.txt_info.get("1.0", "end").strip(),
            "preconisation_text": self.txt_preco.get("1.0", "end").strip(),
            "commentaire":        self.txt_commentaire.get("1.0", "end").strip(),

            "description":       demande,
            "pieces":            "",

            "materiels_json":    json.dumps(materiels, ensure_ascii=False),
            "deplacements_json": json.dumps(deplacements, ensure_ascii=False),
            "moteurs_supplementaires_json": json.dumps(
                [
                    {
                        **{k: row_get(m, k) for k in ("id","num_serie","navire","machine",
                                                       "type_moteur","marque","ref_constructeur",
                                                       "date_mise_service","duree_garantie")},
                        "nb_heures_fct": e.get("nb_heures_var", tk.StringVar()).get().strip(),
                    }
                    for e in self._extra_moteurs_rows
                    if (m := self._moteur_by_ns.get(e["var"].get().strip()))  # noqa: E231
                ],
                ensure_ascii=False),
        }

        if self.is_offline:
            if self._on_offline_save:
                self._on_offline_save(data)
            messagebox.showinfo("Enregistré hors-ligne",
                "✅ Les modifications ont été enregistrées dans le fichier .ems.\n\n"
                "Importez le fichier en ligne pour synchroniser avec le serveur.")
            return

        try:
            if self.is_edit:
                db.update_intervention(self.inv_id, data)
                num_bon = db.get_intervention(inv_id=self.inv_id)["num_bon"]
            else:
                iid, num_bon = db.create_intervention(data)
                self.inv_id = iid
                self.is_edit = True
        except Exception as exc:
            messagebox.showerror("Erreur d'enregistrement",
                f"Impossible d'enregistrer le bon :\n\n{exc}")
            return

        # Lier la garantie sélectionnée à cette intervention
        if not self.is_offline:
            sel_lbl = self.garantie_lien_var.get()
            sel_gar = None
            for num, g in self._garanties_by_num.items():
                if sel_lbl.startswith(num):
                    sel_gar = g
                    break
            if sel_gar:
                try:
                    db.update_garantie(sel_gar["id"], {"intervention_id": self.inv_id})
                except Exception as exc:
                    messagebox.showwarning("Liaison garantie",
                        f"Bon enregistré, mais impossible de lier la garantie :\n{exc}")

        if self.on_save:
            self.on_save()

        # Rafraîchir la liste des contacts après save (intègre les nouveaux noms)
        if not self.is_offline:
            try:
                self._contacts = list(db.get_contacts())
                contact_noms = [c["nom"] for c in self._contacts]
                self._contact_by_nom = {c["nom"]: c for c in self._contacts}
                self.signataire_combo.set_values(contact_noms)
                self.demandeur_combo.set_values(contact_noms)
            except Exception:
                pass

        if generer:
            dossier = self._current_dossier()

            def _faire_gen(photos):
                if photos is None:
                    return   # annule
                inv = db.get_intervention(inv_id=self.inv_id)
                try:
                    path = sauvegarder_bon(inv, photos_annexe=photos, generer_pdf=True)
                except PermissionError as exc:
                    messagebox.showerror("PDF verrouillé", str(exc))
                    return
                messagebox.showinfo("Bon généré",
                    f"✅ {num_bon}\nEnregistré dans :\n{path}\n\n"
                    "Ouverture dans le navigateur.")
                ouvrir_fichier(path)
                self.destroy()

            PhotosAnnexeDialog(self, dossier, on_valider=_faire_gen)
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

    # Libellés français pour les champs DB — mapping fixe affiché par défaut
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

    # Champs ajoutables manuellement via le bouton "+"
    EXTRA_FIELD_LABELS = [
        ("date_mise_service", "Date de mise en service"),
        ("duree_garantie",    "Durée de garantie (mois)"),
        ("machine",           "Machine (type)"),
    ]

    def __init__(self, parent, app, on_save=None):
        super().__init__(parent)
        self.app = app
        self.on_save = on_save
        self.title("Importer le parc depuis un CSV")
        self.geometry("960x720")
        self.configure(bg=C["bg"])
        self.grab_set()

        self.csv_path        = None
        self.headers         = []
        self.preview_rows    = []
        self.total_lines     = 0
        self.mapping         = {}
        self.delimiter       = ";"
        self.encoding        = "utf-8"
        self.column_vars     = {}   # field → StringVar (header sélectionné)
        self._extra_rows_data = []  # [{"field": str, "col_var": StringVar, "frame": Frame}]
        self._extra_cont     = None

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

        self.csv_path         = path
        self.headers          = hd
        self.preview_rows     = prev
        self.total_lines      = total
        self.mapping          = mapping
        self.delimiter        = delim
        self.encoding         = enc
        self._extra_rows_data = []

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

        # ── Liaisons supplémentaires ─────────────────────────────────────────
        tk.Frame(inner, bg="#c0c8d0", height=1).pack(fill="x", padx=4, pady=(12, 4))
        hdr_extra = tk.Frame(inner, bg=C["bg"])
        hdr_extra.pack(fill="x", padx=4, pady=(0, 4))
        tk.Label(hdr_extra, text="Liaisons supplémentaires",
                 bg=C["bg"], font=("Arial", 9, "bold"), fg="#002b5c").pack(side="left")
        mk_btn(hdr_extra, "+ Ajouter",
               lambda c=choices: self._open_extra_field_picker(c),
               color=C["btn2"]).pack(side="left", padx=8)

        self._extra_cont = tk.Frame(inner, bg=C["bg"])
        self._extra_cont.pack(fill="x")

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

    def _open_extra_field_picker(self, choices):
        """Popup pour choisir un champ supplémentaire à mapper."""
        used = {r["field"] for r in self._extra_rows_data}
        available = [(k, lbl) for k, lbl in self.EXTRA_FIELD_LABELS if k not in used]
        if not available:
            messagebox.showinfo("Aucun champ disponible",
                                "Tous les champs supplémentaires ont déjà été ajoutés.")
            return
        labels = [lbl for _, lbl in available]
        keys   = [k   for k, _ in available]
        dlg = tk.Toplevel(self)
        dlg.title("Ajouter un champ")
        dlg.geometry("380x130")
        dlg.configure(bg=C["bg"])
        dlg.resizable(False, False)
        dlg.grab_set()
        tk.Label(dlg, text="Champ moteur à ajouter :",
                 bg=C["bg"], font=("Arial", 10)).pack(pady=(14, 6), padx=20, anchor="w")
        var = tk.StringVar(value=labels[0])
        ttk.Combobox(dlg, textvariable=var, values=labels,
                     state="readonly", width=38).pack(padx=20)
        bf = tk.Frame(dlg, bg=C["bg"])
        bf.pack(pady=10)
        def _confirm():
            idx = labels.index(var.get())
            self._add_extra_row(keys[idx], labels[idx], choices)
            dlg.destroy()
        mk_btn(bf, "Ajouter", _confirm, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "Annuler", dlg.destroy, color="#888").pack(side="left", padx=6)

    def _add_extra_row(self, field_key, field_label, choices):
        """Ajoute une ligne de liaison supplémentaire dans _extra_cont."""
        col_var = tk.StringVar(value="(ignorer)")
        # Auto-détection si le champ figure dans le mapping détecté
        idx = self.mapping.get(field_key)
        if idx is not None and idx < len(self.headers):
            col_var.set(self.headers[idx])

        row_f = tk.Frame(self._extra_cont, bg=C["bg"])
        row_f.pack(fill="x", padx=4, pady=2)
        tk.Label(row_f, text=field_label, bg=C["bg"], font=("Arial", 10),
                 anchor="e", width=28).pack(side="left", padx=(0, 8))
        ttk.Combobox(row_f, textvariable=col_var, values=choices,
                     width=38, state="readonly").pack(side="left", fill="x", expand=True)

        row_data = {"field": field_key, "col_var": col_var, "frame": row_f}
        self._extra_rows_data.append(row_data)

        def _remove():
            self._extra_rows_data.remove(row_data)
            row_f.destroy()
        mk_btn(row_f, "✕", _remove, color=C["danger"]).pack(side="left", padx=(6, 0))

    def _build_mapping_from_ui(self):
        """Reconstruit le dict mapping à partir des combos utilisateur."""
        mapping = {}
        for field, var in self.column_vars.items():
            h = var.get()
            if h == "(ignorer)" or h not in self.headers:
                mapping[field] = None
            else:
                mapping[field] = self.headers.index(h)
        for row_data in self._extra_rows_data:
            h = row_data["col_var"].get()
            if h == "(ignorer)" or h not in self.headers:
                mapping[row_data["field"]] = None
            else:
                mapping[row_data["field"]] = self.headers.index(h)
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


# ══════════════════════════════════════════════════════════════════════════════
# SIGNATURES (client + technicien) — écran tactile tablette
# ══════════════════════════════════════════════════════════════════════════════
class _SignaturePad(tk.Frame):
    """
    Zone de signature tactile réutilisable : canvas blanc, traits noirs.
    Expose .has_signature() et .export_b64() (PNG base64).
    Dimensions par défaut adaptées à une tablette (large + confortable).
    """
    def __init__(self, parent, width=1300, height=360, titre="Signature",
                 sous_titre=""):
        super().__init__(parent, bg=C["bg"])
        self.w, self.h = width, height
        self._strokes = []
        self._points = []
        self._last = None

        tk.Label(self, text=titre, bg=C["bg"], fg=C["header"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 2))
        if sous_titre:
            tk.Label(self, text=sous_titre, bg=C["bg"], fg=C["text_muted"],
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 4))
        wrap = tk.Frame(self, bg=C["accent"])
        wrap.pack()
        self.canvas = tk.Canvas(wrap, width=width, height=height, bg="white",
                                 highlightthickness=0, cursor="pencil")
        self.canvas.pack(padx=2, pady=2)
        self._dessiner_ligne_base()
        self.canvas.bind("<ButtonPress-1>", self._pt_start)
        self.canvas.bind("<B1-Motion>", self._pt_move)
        self.canvas.bind("<ButtonRelease-1>", self._pt_end)

    def _dessiner_ligne_base(self):
        self.canvas.create_line(20, self.h - 50, self.w - 20, self.h - 50,
                                fill="#d0d0d0", dash=(4, 3))
        self.canvas.create_text(self.w // 2, self.h - 24,
                                 text="Signez ici", fill="#c8c8c8",
                                 font=("Segoe UI", 9))

    def _pt_start(self, ev):
        self._last = (ev.x, ev.y)
        self._points = [(ev.x, ev.y)]

    def _pt_move(self, ev):
        if self._last is not None:
            self.canvas.create_line(self._last[0], self._last[1], ev.x, ev.y,
                                    width=4, fill="#0a0a28",
                                    capstyle="round", smooth=True)
            self._last = (ev.x, ev.y)
            self._points.append((ev.x, ev.y))

    def _pt_end(self, _ev):
        if len(self._points) > 1:
            self._strokes.append(list(self._points))
        self._points = []
        self._last = None

    def effacer(self):
        self.canvas.delete("all")
        self._dessiner_ligne_base()
        self._strokes = []
        self._points = []
        self._last = None

    def has_signature(self):
        return bool(self._strokes)

    def export_b64(self):
        """Rend les traits en PNG base64 (PIL requis)."""
        import base64, io
        try:
            from PIL import Image, ImageDraw
            im = Image.new("RGB", (self.w, self.h), "white")
            dr = ImageDraw.Draw(im)
            for stroke in self._strokes:
                if len(stroke) >= 2:
                    dr.line(stroke, fill=(10, 10, 40), width=4,
                            joint="curve")
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            return base64.b64encode(buf.getvalue()).decode("ascii")
        except Exception:
            return None


class SignatureDialog(tk.Toplevel):
    """
    Écran de signature tactile en deux étapes successives :
      Étape 1 — Signature CLIENT (avec acceptation)
      Étape 2 — Signature TECHNICIEN EMS
    À la validation finale, les deux signatures sont enregistrées et
    la fenêtre se ferme automatiquement.
    """
    def __init__(self, parent, app, inv_id=None, on_save=None,
                 offline_bundle=None, offline_path=None, on_offline_save=None):
        super().__init__(parent)
        self.app = app
        self.inv_id = inv_id
        self.on_save = on_save
        self._offline_bundle = offline_bundle
        self._offline_path = offline_path
        self._on_offline_save = on_offline_save
        self.is_offline = offline_bundle is not None

        if self.is_offline:
            orig  = offline_bundle.get("intervention", {})
            edits = offline_bundle.get("offline_edits") or {}
            self.inv = {**orig, **edits}
        else:
            self.inv = db.get_intervention(inv_id=inv_id)
            if not self.inv:
                messagebox.showerror("Erreur", "Intervention introuvable.")
                self.destroy()
                return

        # Données accumulées entre les étapes
        self._client_b64 = None
        self._client_nom = ""
        self._tech_b64 = None
        self._tech_nom = ""
        self._etape = 1

        self.title("Signatures – Validation de l'intervention")
        self.configure(bg=C["bg"])
        self.attributes("-topmost", True)
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                self.geometry("1100x820")
        self.grab_set()
        self.bind("<Escape>", lambda _e: self._annuler())

        # Résolution écran — stockée pour dimensionner tout le dialog
        try:
            self._sh = self.winfo_screenheight()
            self._sw = self.winfo_screenwidth()
        except tk.TclError:
            self._sh, self._sw = 900, 1600

        # ── Header compact et responsive ─────────────────────────────────
        hdr_h    = 36 if self._sh < 900 else (48 if self._sh < 1080 else 58)
        hdr_fsz  = 11 if self._sh < 900 else (13 if self._sh < 1080 else 15)
        bar = tk.Frame(self, bg=C["header"], height=hdr_h)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Frame(bar, bg=C["accent"], width=4).pack(side="left", fill="y")
        tk.Label(bar, text="Validation de l'intervention",
                 font=("Segoe UI", hdr_fsz, "bold"),
                 bg=C["header"], fg="white", padx=20).pack(side="left")

        self._recap_widget()

        # ── Boutons (re-créés à chaque étape) — packés en premier pour rester visibles ──
        self.bf = tk.Frame(self, bg=C["bg"])
        self.bf.pack(side="bottom", pady=(6, 16))

        # ── Indicateur d'étape ──────────────────────────────────────────
        self.steps_bar = tk.Frame(self, bg=C["bg"])
        self.steps_bar.pack(side="bottom", fill="x", padx=24, pady=(0, 4))
        self._maj_indicateur()

        # ── Zone d'étape (remplacée à chaque étape) ──────────────────────
        self.zone = tk.Frame(self, bg=C["bg"])
        self.zone.pack(fill="both", expand=True, padx=24, pady=4)

        self._afficher_etape_client()

    # ── Récapitulatif ─────────────────────────────────────────────────────
    def _recap_widget(self):
        sh = self._sh
        if self.is_offline:
            client_nom = row_get(self.inv, "client_nom") or ""
            client = {"nom": client_nom} if client_nom else None
            moteur = self.inv
        else:
            client = db.get_client(self.inv["client_id"]) if self.inv["client_id"] else None
            moteur = db.get_moteur(self.inv["moteur_id"]) if self.inv["moteur_id"] else None

        font_sz = 8  if sh < 800  else (9  if sh < 1000 else 10)
        pad_y   = (2, 1) if sh < 800 else ((4, 2) if sh < 1000 else (6, 3))
        cell_px = 8  if sh < 800  else (12 if sh < 1000 else 18)

        recap = tk.Frame(self, bg=C["surface"], bd=1, relief="solid")
        recap.pack(fill="x", padx=24, pady=pad_y)
        lignes = [
            ("N° de bon", self.inv["num_bon"]),
            ("Client", (client["nom"] if client else "") or "—"),
            ("Navire / Site", row_get(moteur, "navire") or
                              row_get(self.inv, "lieu_intervention") or "—"),
        ]
        row_frame = tk.Frame(recap, bg=C["surface"])
        row_frame.pack(fill="x", padx=cell_px, pady=pad_y)
        for lbl, val in lignes:
            cell = tk.Frame(row_frame, bg=C["surface"])
            cell.pack(side="left", padx=cell_px)
            tk.Label(cell, text=lbl + " : ", bg=C["surface"],
                     fg=C["text_muted"],
                     font=("Segoe UI", font_sz, "bold")).pack(side="left")
            tk.Label(cell, text=str(val), bg=C["surface"], fg=C["text"],
                     font=("Segoe UI", font_sz)).pack(side="left")
        # Priorité au nom du signataire saisi sur le bon, sinon nom du client
        self._client_nom_defaut = (row_get(self.inv, "nom_signataire") or
                                   (client["nom"] if client else "") or "")

    def _maj_indicateur(self):
        for w in self.steps_bar.winfo_children():
            w.destroy()
        for n, libelle in [(1, "1. Signature CLIENT"),
                            (2, "2. Signature TECHNICIEN")]:
            actif = (n == self._etape)
            fait = (n < self._etape)
            if fait:
                bg, fg, prefix = C["btn2"], "white", "✓ "
            elif actif:
                bg, fg, prefix = C["header"], "white", "▶ "
            else:
                bg, fg, prefix = C["bg_alt"], C["text_muted"], "  "
            fsz = 8 if self._sh < 800 else (9 if self._sh < 1000 else 10)
            pxy = 3 if self._sh < 800 else (4 if self._sh < 1000 else 5)
            tk.Label(self.steps_bar, text=prefix + libelle,
                     bg=bg, fg=fg, font=("Segoe UI", fsz, "bold"),
                     padx=10, pady=pxy).pack(side="left", padx=2)

    # ── ÉTAPE 1 : CLIENT ─────────────────────────────────────────────────
    def _afficher_etape_client(self):
        self._etape = 1
        self._maj_indicateur()
        for w in self.zone.winfo_children():
            w.destroy()
        for w in self.bf.winfo_children():
            w.destroy()

        sh = self._sh
        t_fsz  = 11 if sh < 800 else (12 if sh < 1000 else 14)
        b_fsz  = 9  if sh < 800 else 10
        py_t   = (3, 1) if sh < 800 else ((5, 1) if sh < 1000 else (8, 2))
        py_sub = (0, 2) if sh < 800 else ((0, 3) if sh < 1000 else (0, 4))
        py_chk = (0, 3) if sh < 800 else ((0, 4) if sh < 1000 else (0, 6))
        py_nf  = (4, 2) if sh < 800 else ((6, 3) if sh < 1000 else (10, 4))
        py_acc = 2 if sh < 800 else 4

        tk.Label(self.zone, text="✍  SIGNATURE CLIENT",
                 bg=C["bg"], fg=C["header"],
                 font=("Segoe UI", t_fsz, "bold")).pack(pady=py_t)
        if sh >= 800:
            tk.Label(self.zone,
                     text="À faire signer au client (au doigt sur la tablette)",
                     bg=C["bg"], fg=C["text_muted"],
                     font=("Segoe UI", b_fsz)).pack(pady=py_sub)

        # ── Case "client absent" ──────────────────────────────────────────
        self.client_absent_var = tk.IntVar(value=0)
        self._absent_chk = tk.Checkbutton(
            self.zone,
            text="  Client absent — passer directement à la signature technicien",
            variable=self.client_absent_var,
            bg=C["bg"], fg=C["warn"], font=("Segoe UI", b_fsz, "bold"),
            activebackground=C["bg"], selectcolor="white",
            command=self._toggle_client_absent)
        self._absent_chk.pack(pady=py_chk)

        # Badge affiché quand client absent est coché
        self._absent_badge = tk.Label(self.zone, text="⚠  CLIENT ABSENT — signature ignorée",
                                       bg="#fff3cd", fg="#856404",
                                       font=("Segoe UI", b_fsz + 1, "bold"), padx=10, pady=4,
                                       relief="solid", bd=1)
        # (non packé par défaut — apparaît via _toggle_client_absent)

        pad_w = max(800, min(self._sw - 200, 1600))
        pad_h = max(180, min(sh - 560, 400))
        self.pad_client = _SignaturePad(self.zone, width=pad_w, height=pad_h,
                                         titre="", sous_titre="")
        self.pad_client.pack()

        nf = tk.Frame(self.zone, bg=C["bg"])
        nf.pack(pady=py_nf)
        tk.Label(nf, text="Nom du signataire * : ", bg=C["bg"],
                 font=("Segoe UI", b_fsz), fg=C["text"]).pack(side="left")
        self.nom_client_var = tk.StringVar(value=self._client_nom_defaut)
        self._nom_client_entry = ttk.Entry(nf, textvariable=self.nom_client_var,
                                            width=40, font=("Segoe UI", b_fsz + 1))
        self._nom_client_entry.pack(side="left")

        self.accept_var = tk.IntVar(value=0)
        self._accept_chk = tk.Checkbutton(
            self.zone,
            text=" Le client reconnaît avoir pris connaissance de "
                 "l'intervention et en accepte la réalisation.",
            variable=self.accept_var, bg=C["bg"], fg=C["text"],
            font=("Segoe UI", b_fsz), activebackground=C["bg"],
            selectcolor="white", anchor="w")
        self._accept_chk.pack(pady=py_acc)

        self._btn_effacer_client = mk_btn(self.bf, "🧹 Effacer",
                                           self.pad_client.effacer,
                                           color=C["btn3"])
        self._btn_effacer_client.pack(side="left", padx=8)
        mk_btn(self.bf, "Annuler", self._annuler,
               color=C["danger"]).pack(side="left", padx=8)
        mk_btn(self.bf, "Suivant : signature technicien  ▸",
               self._passer_a_technicien,
               color=C["btn2"]).pack(side="left", padx=8)

    def _toggle_client_absent(self):
        absent = bool(self.client_absent_var.get())
        if absent:
            self._absent_badge.pack(pady=(0, 6))
            self.pad_client.pack_forget()
            self._nom_client_entry.config(state="disabled")
            self.nom_client_var.set("Client absent")
            self._accept_chk.config(state="disabled")
            self._btn_effacer_client.config(state="disabled")
        else:
            self._absent_badge.pack_forget()
            self.pad_client.pack()
            self._nom_client_entry.config(state="normal")
            self.nom_client_var.set(self._client_nom_defaut)
            self._accept_chk.config(state="normal")
            self._btn_effacer_client.config(state="normal")

    def _passer_a_technicien(self):
        # ── Client absent : bypass total ──────────────────────────────────
        if self.client_absent_var.get():
            self._client_b64 = ""
            self._client_nom = "Client absent"
            self._afficher_etape_technicien()
            return

        # ── Validation normale ────────────────────────────────────────────
        if not self.pad_client.has_signature():
            self._msg("showwarning", "Signature requise",
                                    "Merci de signer dans le cadre.")
            return
        nom = self.nom_client_var.get().strip()
        if not nom:
            self._msg("showwarning", "Nom requis",
                                    "Indiquez le nom du signataire client.")
            return
        if not self.accept_var.get():
            self._msg("showwarning", "Acceptation requise",
                "Cochez la case d'acceptation.")
            return
        b64 = self.pad_client.export_b64()
        if not b64:
            self._msg("showerror", "Erreur",
                "Impossible de générer l'image de signature.")
            return
        self._client_b64 = b64
        self._client_nom = nom
        self._afficher_etape_technicien()

    # ── ÉTAPE 2 : TECHNICIEN ─────────────────────────────────────────────
    def _afficher_etape_technicien(self):
        self._etape = 2
        self._maj_indicateur()
        for w in self.zone.winfo_children():
            w.destroy()
        for w in self.bf.winfo_children():
            w.destroy()

        sh = self._sh
        t_fsz = 11 if sh < 800 else (12 if sh < 1000 else 14)
        b_fsz = 9  if sh < 800 else 10
        py_t  = (3, 1) if sh < 800 else ((5, 1) if sh < 1000 else (8, 2))
        py_sub = (0, 4) if sh < 800 else (0, 8)
        py_nf  = (4, 2) if sh < 800 else ((6, 3) if sh < 1000 else (10, 4))

        tk.Label(self.zone, text="✍  SIGNATURE TECHNICIEN EMS",
                 bg=C["bg"], fg=C["header"],
                 font=("Segoe UI", t_fsz, "bold")).pack(pady=py_t)
        if sh >= 800:
            tk.Label(self.zone,
                     text="Le technicien atteste la réalisation des travaux",
                     bg=C["bg"], fg=C["text_muted"],
                     font=("Segoe UI", b_fsz)).pack(pady=py_sub)

        pad_w = max(800, min(self._sw - 200, 1600))
        pad_h = max(180, min(sh - 560, 400))
        self.pad_tech = _SignaturePad(self.zone, width=pad_w, height=pad_h,
                                       titre="", sous_titre="")
        self.pad_tech.pack()

        nf = tk.Frame(self.zone, bg=C["bg"])
        nf.pack(pady=py_nf)
        tk.Label(nf, text="Nom du technicien * : ", bg=C["bg"],
                 font=("Segoe UI", b_fsz), fg=C["text"]).pack(side="left")
        self.nom_tech_var = tk.StringVar(
            value=row_get(self.inv, "technicien") or "")
        ttk.Entry(nf, textvariable=self.nom_tech_var, width=40,
                  font=("Segoe UI", b_fsz + 1)).pack(side="left")

        if sh >= 800:
            tk.Label(self.zone,
                     text=" J'atteste sur l'honneur la réalisation des travaux "
                          "décrits ci-dessus.",
                     bg=C["bg"], fg=C["text_muted"],
                     font=("Segoe UI", b_fsz, "italic")).pack(pady=2)

        mk_btn(self.bf, "◂ Retour", self._afficher_etape_client,
               color=C["btn3"]).pack(side="left", padx=8)
        mk_btn(self.bf, "🧹 Effacer", self.pad_tech.effacer,
               color=C["btn3"]).pack(side="left", padx=8)
        mk_btn(self.bf, "Annuler", self._annuler,
               color=C["danger"]).pack(side="left", padx=8)
        mk_btn(self.bf, "✓ Valider et enregistrer",
               self._valider_final,
               color=C["btn2"]).pack(side="left", padx=8)

    def _annuler(self):
        if self._msg("askyesno", "Annuler",
                               "Abandonner les signatures ?"):
            self.destroy()

    # ── Validation finale ────────────────────────────────────────────────
    def _valider_final(self):
        if not self.pad_tech.has_signature():
            self._msg("showwarning", "Signature requise",
                                    "Merci de signer dans le cadre.")
            return
        nom = self.nom_tech_var.get().strip()
        if not nom:
            self._msg("showwarning", "Nom requis",
                                    "Indiquez le nom du technicien.")
            return
        b64 = self.pad_tech.export_b64()
        if not b64:
            self._msg("showerror", "Erreur",
                "Impossible de générer l'image de signature.")
            return
        self._tech_b64 = b64
        self._tech_nom = nom

        # ── Mode hors-ligne : enregistrement dans le bundle .ems ─────────────
        if self.is_offline:
            from shared import import_export as ie
            from datetime import datetime as _dt
            edits = dict(ie.get_data(self._offline_bundle))
            now = _dt.now().strftime("%Y-%m-%d %H:%M")
            edits["signature_b64"] = self._client_b64 or ""
            edits["signature_nom"] = self._client_nom or ""
            edits["signature_date"] = now if self._client_b64 else ""
            edits["signature_tech_b64"] = self._tech_b64
            edits["signature_tech_nom"] = self._tech_nom
            edits["signature_tech_date"] = now
            try:
                ie.apply_offline_edits(self._offline_bundle, edits,
                                       self._offline_path)
            except Exception as e:
                self._msg("showerror", "Erreur d'enregistrement",
                          f"Impossible de sauvegarder les signatures :\n{e}")
                return
            if self._client_b64:
                self._sauver_png(self._client_b64, "signature_client.png")
            self._sauver_png(self._tech_b64, "signature_technicien.png")
            if self._on_offline_save:
                try:
                    self._on_offline_save()
                except Exception as e_cb:
                    print(f"[Signature] on_offline_save erreur : {e_cb}")
            try:
                self.attributes("-topmost", False)
            except tk.TclError:
                pass
            msg = ("✅ Signatures enregistrées hors-ligne.\n\n"
                   "Elles seront transmises au serveur lors de l'import.")
            parent = self.master
            try:
                self.destroy()
            except tk.TclError:
                pass
            try:
                messagebox.showinfo("Signatures enregistrées", msg,
                                    parent=parent)
            except tk.TclError:
                messagebox.showinfo("Signatures enregistrées", msg)
            return

        # ── Mode en ligne : enregistrement via API + régénération bon ────────
        try:
            h_c = db.enregistrer_signature(
                self.inv_id, self._client_b64, self._client_nom,
                role="client")
            self._sauver_png(self._client_b64, "signature_client.png")
            h_t = db.enregistrer_signature(
                self.inv_id, self._tech_b64, self._tech_nom,
                role="technicien")
            self._sauver_png(self._tech_b64, "signature_technicien.png")
            try:
                inv2 = db.get_intervention(inv_id=self.inv_id)
                sauvegarder_bon(inv2, generer_pdf=True)
            except Exception as e_html:
                print(f"[Signature] HTML non régénéré : {e_html}")
        except Exception as e:
            self._msg("showerror", "Erreur d'enregistrement",
                      f"Impossible d'enregistrer les signatures :\n{e}")
            return

        if self.on_save:
            try:
                self.on_save()
            except Exception as e_cb:
                print(f"[Signature] on_save erreur : {e_cb}")

        # Désactiver topmost AVANT, et FERMER la fenêtre AVANT le message
        # de confirmation (sinon le popup reste invisible derrière).
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass
        msg = (f"✅ Client : {self._client_nom} ({h_c})\n"
               f"✅ Technicien : {self._tech_nom} ({h_t})\n\n"
               "Les signatures ont été ajoutées au bon d'intervention.")
        parent = self.master
        try:
            self.destroy()
        except tk.TclError:
            pass
        try:
            messagebox.showinfo("Signatures enregistrées", msg,
                                parent=parent)
        except tk.TclError:
            messagebox.showinfo("Signatures enregistrées", msg)

    # ── Helpers messagebox (visibles au-dessus de la fenêtre topmost) ────
    def _msg(self, kind, title, message):
        """Affiche un messagebox de façon visible, en désactivant
        temporairement le topmost. parent=self pour rattacher la boîte."""
        try:
            self.attributes("-topmost", False)
            self.update_idletasks()
        except tk.TclError:
            pass
        try:
            fn = getattr(messagebox, kind)
            return fn(title, message, parent=self)
        finally:
            try:
                self.attributes("-topmost", True)
                self.lift()
            except tk.TclError:
                pass

    def _sauver_png(self, b64, nom_fichier):
        """Sauvegarde une signature PNG dans un sous-dossier cache .signatures
        pour ne pas polluer le dossier principal du bon."""
        try:
            import base64 as _b64
            dossier = _dossiers_root() / self.inv["num_bon"]
            sub = dossier / ".signatures"
            sub.mkdir(parents=True, exist_ok=True)
            # Cacher le sous-dossier (Windows : attribut H + S)
            try:
                import subprocess
                subprocess.run(["attrib", "+H", "+S", str(sub)],
                               check=False, capture_output=True)
            except Exception:
                pass
            (sub / nom_fichier).write_bytes(_b64.b64decode(b64))
        except OSError:
            pass

# ─── Lancement ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    # Mode passé en argument : --mode=bons | --mode=parc | (défaut : full)
    _mode = "full"
    for a in sys.argv[1:]:
        if a.startswith("--mode="):
            _mode = a.split("=", 1)[1].strip()
        elif a in ("bons", "parc", "full"):
            _mode = a
    app = AppEMS(mode=_mode)
    app.mainloop()
