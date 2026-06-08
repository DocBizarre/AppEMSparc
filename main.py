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
from shared.bon_generator import sauvegarder_bon, ouvrir_fichier


def _check_api_startup():
    """Affiche un avertissement si le serveur est injoignable au démarrage."""
    ok, msg = db.check_api()
    if not ok:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        _r = _tk.Tk(); _r.withdraw()
        if not _mb.askokcancel(
            "Serveur introuvable",
            f"{msg}\n\nVérifiez que le serveur EMS est démarré.\n\n"
            "Continuer quand même ?"):
            raise SystemExit(0)


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
    "afact":        "#6366f1",  # À facturer - indigo (intermédiaire)
    "afact_bg":     "#eef2ff",
    "fact":         "#3b82f6",  # Facturé - bleu vif
    "fact_bg":      "#eff6ff",
    "clos":         "#10b981",  # Clos - vert
    "clos_bg":      "#ecfdf5",
}
STATUTS  = ["En cours", "À facturer", "Facturé", "Clos"]
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
                 case_sensitive=False, max_visible=8, **kw):
        for k in ("state", "values", "textvariable", "validate", "validatecommand"):
            kw.pop(k, None)
        super().__init__(master, bg=master.cget("bg") if "bg" not in kw else kw.get("bg"))

        self._all_values     = list(values or [])
        self._filtered       = list(self._all_values)
        self._case_sensitive = case_sensitive
        self._max_visible    = max_visible
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
        if query:
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
        # Coloration informative si valeur hors liste
        v = self.var.get().strip()
        try:
            if v and v not in self._all_values:
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
            self._filtered = self._filter(self.var.get())
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
    def __init__(self, master, available=None, on_add_new=None, **kw):
        super().__init__(master, bg=master.cget("bg") if "bg" not in kw else kw.get("bg"))
        self._available = list(available or [])
        self._selected  = []
        self._on_add_new = on_add_new  # callback : ouvre TechnicienDialog

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

    def _remove(self, name):
        if name in self._selected:
            self._selected.remove(name)
            self._render_chips()

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
        "bons": ["dashboard", "interventions", "nouveau"],
        "parc": ["clients", "moteurs", "techniciens"],
        "full": ["dashboard", "interventions", "clients",
                 "moteurs", "techniciens", "nouveau"],
    }
    TITRES = {
        "bons": "EMS – Bons d'intervention",
        "parc": "EMS – Gestion de parc",
        "full": "EMS – Gestion des Interventions",
    }

    def __init__(self, mode="full"):
        super().__init__()
        self.mode = mode if mode in self.NAV_PARELEMENTS else "full"
        self._onglets_actifs = self.NAV_PARELEMENTS[self.mode]
        self.title(self.TITRES[self.mode])
        self.geometry("1340x830")
        self.minsize(960, 620)
        self.configure(bg=C["bg"])
        self._init_ttk_styles()
        db.init_db()
        self._build()
        self.show(self._onglets_actifs[0])

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
            "En cours": stats["En cours"], "À facturer": stats["À facturer"],
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
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
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
        cols = ("tech","ec","afact","fact","clos","total")
        col_defs = [("tech","Technicien",180),
                    ("ec",   "En cours",   75, "center"),
                    ("afact","À facturer", 80, "center"),
                    ("fact", "Facturé",    75, "center"),
                    ("clos", "Clos",       70, "center"),
                    ("total","Total",      65, "center")]
        tf, self.tree = mk_tree(self, cols, col_defs, height=5)
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
            BonDialog(self, self.app, inv_id=r["id"], on_save=lambda: (self.app.refresh_frame("dashboard"), self.app.refresh_frame("interventions")))
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
                "par_technicien", "non_notifies"])
            db.set_dashboard_cards([
                "En cours", "À facturer", "Facturé", "Clos", "Total",
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
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
        mailer.email_client(inv, client, moteur, str(path))
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
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
        mailer.email_technicien(inv, client, moteur, emails, str(path))
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
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
        mailer.email_cloture(inv, client, moteur, tech_emails, str(path))

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
        tf, self.tree = mk_tree(self, cols, col_defs, height=22)
        tf.pack(fill="both", expand=True, padx=20, pady=4)

        af = tk.Frame(self, bg=C["bg"])
        af.pack(fill="x", padx=20, pady=8)
        mk_btn(af, "✏️ Modifier", self._modifier).pack(side="left", padx=4)
        mk_btn(af, "📋 Fiche & historique", self._voir_inv).pack(side="left", padx=4)
        self.btn_suppr = mk_btn(af, "🗑️ Supprimer",
                                 self._supprimer, color=C["danger"])
        self.btn_suppr.pack(side="left", padx=4)
        self.sel_info = tk.Label(af, text="", bg=C["bg"],
                                  fg=C["text_muted"], font=F["small"])
        self.sel_info.pack(side="left", padx=(12, 0))
        self.tree.bind("<Double-1>", lambda e: self._modifier())
        self.tree.bind("<<TreeviewSelect>>", self._on_select_change)
        self._cache = []

    def _on_select_change(self, _ev=None):
        n = len(self.tree.selection())
        if n <= 1:
            self.sel_info.config(text="")
            self.btn_suppr.config(text="🗑️ Supprimer")
        else:
            self.sel_info.config(
                text=f"{n} moteurs sélectionnés "
                     "(Ctrl/Maj-clic pour ajuster)")
            self.btn_suppr.config(text=f"🗑️ Supprimer ({n})")

    def refresh(self):
        ms = db.get_moteurs(search=self.search_var.get(),
                             serie_only=bool(self.serie_only_var.get()))
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
        """Premier moteur sélectionné (compat avec actions unitaires)."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection","Sélectionnez un moteur."); return None
        return self._cache[int(sel[0])]

    def _sel_multi(self):
        """Liste des moteurs sélectionnés (Ctrl/Shift click pris en charge)."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning(
                "Sélection",
                "Sélectionnez un ou plusieurs moteurs.\n"
                "(Ctrl-clic ou Maj-clic pour en sélectionner plusieurs)")
            return []
        return [self._cache[int(s)] for s in sel]

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
        self.geometry("700x760")
        self.configure(bg=C["bg"])
        self.grab_set()

        self._clients = list(db.get_clients())
        self._moteurs = list(db.get_moteurs())
        self._moteur_by_ns = {m["num_serie"]: m for m in self._moteurs}
        self._client_by_id = {c["id"]: c for c in self._clients}

        mk_header(self, "Garantie moteur",
                  "   Dossier de garantie constructeur / interne")

        body = tk.Frame(self, bg=C["bg"])
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
class DeplacementsTable(tk.Frame):
    """Saisie des temps et frais lies a l'intervention.
    
    Champs texte : trajet aller-retour, duree intervention,
                   temps preparation, temps rangement.
    Cases a cocher : frais repas, frais hotel, frais peages.
    
    Le 'load' assure la retrocompat avec l'ancien format
    (trajet_aller/trajet_retour, heure_debut_matin, etc.).
    """
    # Champs texte : (cle, libelle)
    TEXTE_FIELDS = [
        ("trajet_aller_retour", "Trajet aller-retour"),
        ("duree_intervention",  "Duree de l'intervention"),
        ("temps_preparation",   "Temps de preparation"),
        ("temps_rangement",     "Temps de rangement"),
    ]
    # Cases a cocher : (cle, libelle)
    CHECK_FIELDS = [
        ("frais_repas",  "Frais de repas"),
        ("frais_hotel",  "Frais d'hotel"),
        ("frais_peages", "Frais de peages"),
    ]

    def __init__(self, master, **kw):
        super().__init__(master, bg=C["bg"], **kw)
        self.vars = {}       # cle -> StringVar (champs texte)
        self.check_vars = {} # cle -> IntVar (cases)

        # Bloc gauche : champs texte (4 lignes)
        for key, lbl in self.TEXTE_FIELDS:
            row = tk.Frame(self, bg=C["bg"])
            row.pack(fill="x", pady=2)
            tk.Label(row, text=lbl, bg=C["bg"], font=("Arial", 9),
                     width=30, anchor="w").pack(side="left", padx=(0, 4))
            self.vars[key] = tk.StringVar()
            ttk.Entry(row, textvariable=self.vars[key], width=18).pack(side="left")

        # Separateur visuel + cases a cocher (frais)
        sep = tk.Frame(self, bg="#d0d4d9", height=1)
        sep.pack(fill="x", pady=(8, 4))

        tk.Label(self, text="Frais :", bg=C["bg"],
                 font=("Arial", 9, "bold"), anchor="w").pack(
                     anchor="w", padx=(0, 4))

        chk_row = tk.Frame(self, bg=C["bg"])
        chk_row.pack(fill="x", pady=2)
        for key, lbl in self.CHECK_FIELDS:
            self.check_vars[key] = tk.IntVar(value=0)
            tk.Checkbutton(chk_row, text=lbl,
                           variable=self.check_vars[key],
                           bg=C["bg"], font=("Arial", 9),
                           activebackground=C["bg"],
                           selectcolor="white").pack(side="left", padx=(0, 18))

    def load(self, data):
        """Charge les valeurs depuis le JSON. Gere la retrocompat avec
        l'ancien format (trajet_aller + trajet_retour -> trajet_aller_retour, etc.)"""
        if not data:
            data = {}
        # Champs texte
        for key, _ in self.TEXTE_FIELDS:
            val = data.get(key, "")
            # Retrocompat : si pas de valeur nouvelle, essayer agreger ancienne
            if not val:
                if key == "trajet_aller_retour":
                    aller = str(data.get("trajet_aller", "")).strip()
                    retour = str(data.get("trajet_retour", "")).strip()
                    if aller and retour:
                        val = f"Aller : {aller} / Retour : {retour}"
                    else:
                        val = aller or retour
                elif key == "duree_intervention":
                    # Tenter de combiner heure_debut_matin / heure_fin_apres
                    debut = str(data.get("heure_debut_matin", "")).strip()
                    fin = str(data.get("heure_fin_apres", "")).strip()
                    if debut and fin:
                        val = f"{debut} -> {fin}"
                    else:
                        val = debut or fin
            self.vars[key].set(str(val))
        # Cases a cocher : 1 si valeur truthy
        for key, _ in self.CHECK_FIELDS:
            v = data.get(key, 0)
            # Retrocompat : ancien format stockait texte (montant) ou 0
            try:
                self.check_vars[key].set(1 if int(v) else 0)
            except (ValueError, TypeError):
                # Texte non vide = case cochee
                self.check_vars[key].set(1 if str(v).strip() else 0)

    def to_dict(self):
        out = {}
        for key, _ in self.TEXTE_FIELDS:
            v = self.vars[key].get().strip()
            if v:
                out[key] = v
        for key, _ in self.CHECK_FIELDS:
            out[key] = int(self.check_vars[key].get())
        return out

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
        ]
        for i, (lbl, val) in enumerate(champs):
            r, c = divmod(i, 2)
            cell = tk.Frame(grille, bg=C["surface"])
            cell.grid(row=r, column=c, sticky="w", padx=10, pady=2)
            tk.Label(cell, text=f"{lbl} : ", bg=C["surface"],
                     font=("Arial", 9, "bold"), fg="#555").pack(side="left")
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
                 anchor="e", width=22).grid(row=0, column=0, sticky="ne", padx=(0,6), pady=4)
        self._client_combo = SearchableCombobox(
            f, textvariable=self.client_var,
            values=[c["nom"] for c in self._clients],
            width=33)
        self._client_combo.grid(row=0, column=1, pady=4, sticky="w")
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
# DIALOG SELECTION PHOTOS POUR ANNEXE
# ══════════════════════════════════════════════════════════════════════════════
class PhotosAnnexeDialog(tk.Toplevel):
    """Popup : choisir les photos du dossier a inclure en annexe du bon.
    Retourne la liste des chemins coches via le callback on_valider."""
    EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

    def __init__(self, parent, dossier, on_valider):
        super().__init__(parent)
        self.on_valider = on_valider
        self.dossier = Path(dossier)
        self._vars = {}
        self.title("Photos à inclure en annexe")
        self.configure(bg=C["bg"])
        self.geometry("640x560")
        self.grab_set()

        # Lister les images du dossier
        self._images = sorted(
            [f for f in self.dossier.iterdir()
             if f.is_file() and f.suffix.lower() in self.EXTS],
            key=lambda p: p.name.lower()) if self.dossier.is_dir() else []

        tk.Label(self, text="📷 Photos disponibles dans le dossier",
                 bg=C["bg"], font=("Arial", 12, "bold"),
                 fg=C["header"]).pack(pady=(14, 2))
        tk.Label(self, text=f"Dossier : {self.dossier.name}",
                 bg=C["bg"], font=("Arial", 9), fg="#666").pack()

        if not self._images:
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

        # Zone scrollable avec miniatures
        cont = tk.Frame(self, bg=C["bg"])
        cont.pack(fill="both", expand=True, padx=20, pady=6)
        canvas = tk.Canvas(cont, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(cont, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=C["bg"])
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win, width=e.width))

        def _wheel(ev):
            canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        canvas.bind("<Enter>",
                    lambda e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>",
                    lambda e: canvas.unbind_all("<MouseWheel>"))

        self._vars = {}
        self._thumbs = []   # garder les references aux PhotoImage
        for img_path in self._images:
            self._add_photo_row(inner, img_path)

        # Boutons
        bf = tk.Frame(self, bg=C["bg"]); bf.pack(pady=12)
        mk_btn(bf, "✓ Générer avec ces photos",
               self._valider, color=C["btn2"]).pack(side="left", padx=6)
        mk_btn(bf, "Annuler", self._annuler, color="#888").pack(side="left", padx=6)

    def _add_photo_row(self, parent, img_path):
        rf = tk.Frame(parent, bg=C["surface"], relief="solid", bd=1)
        rf.pack(fill="x", pady=3, padx=2)
        var = tk.IntVar(value=1)   # coché par défaut
        self._vars[str(img_path)] = var
        tk.Checkbutton(rf, variable=var, bg=C["surface"]).pack(side="left", padx=6)
        # Miniature
        thumb = self._make_thumb(img_path)
        if thumb:
            tk.Label(rf, image=thumb, bg=C["surface"]).pack(side="left", padx=6, pady=4)
            self._thumbs.append(thumb)
        else:
            tk.Label(rf, text="🖼", font=("Arial", 24),
                     bg=C["surface"]).pack(side="left", padx=14)
        # Nom
        tk.Label(rf, text=img_path.name, bg=C["surface"],
                 font=("Arial", 10), anchor="w").pack(side="left", padx=8)

    def _make_thumb(self, img_path, size=64):
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
        photos = [p for p, v in self._vars.items() if v.get()]
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
        try:
            self.state("zoomed")     # plein écran Windows
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)   # Linux
            except tk.TclError:
                pass
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
        mk_btn(row, "+", lambda: ClientDialog(self, self.app, on_save=self._reload_refs),
               color=C["btn2"]).pack(side="left", padx=(8,0))

        self.lieu_var       = tk.StringVar()
        self.signataire_var = tk.StringVar()
        self.email_sig_var  = tk.StringVar()
        self.tel_sig_var    = tk.StringVar()
        self.demandeur_var      = tk.StringVar()
        self.email_demand_var   = tk.StringVar()
        self.tel_demand_var     = tk.StringVar()

        # Lieu + signataire (personne qui signe le bon)
        for lbl, var, cls in [("Lieu de l'intervention", self.lieu_var,       ttk.Entry),
                              ("Nom du signataire",      self.signataire_var, ttk.Entry),
                              ("Courriel du signataire", self.email_sig_var,  EmailEntry),
                              ("Telephone signataire",   self.tel_sig_var,    ttk.Entry)]:
            r = tk.Frame(p, bg=C["bg"])
            r.pack(fill="x", padx=20, pady=2)
            tk.Label(r, text=lbl, bg=C["bg"], font=("Arial",10),
                     width=24, anchor="e").pack(side="left", padx=(0,8))
            cls(r, textvariable=var, width=46).pack(side="left")

        # Separateur visuel + demandeur (personne qui a appele pour declencher l'intervention)
        tk.Label(p, text="  Demandeur (personne ayant appele)",
                 bg=C["bg"], font=("Arial", 9, "italic"),
                 fg="#6b7785").pack(anchor="w", padx=20, pady=(8, 2))
        for lbl, var, cls in [("Nom du demandeur",        self.demandeur_var,    ttk.Entry),
                              ("Courriel du demandeur",   self.email_demand_var, EmailEntry),
                              ("Telephone du demandeur",  self.tel_demand_var,   ttk.Entry)]:
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
        ttk.Entry(rh, textvariable=self.nb_heures_var, width=20).pack(side="left")

        # N° de commande client (champ optionnel, utile pour la facturation)
        tk.Label(rh, text="  N° commande client", bg=C["bg"], font=("Arial",10)
            ).pack(side="left", padx=(16, 8))
        self.num_cmd_var = tk.StringVar()
        ttk.Entry(rh, textvariable=self.num_cmd_var, width=22).pack(side="left")

        rg = tk.Frame(p, bg=C["bg"]); rg.pack(fill="x", padx=20, pady=2)
        tk.Label(rg, text="Garanties du moteur", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.garantie_lbl = tk.Label(rg, text="—", bg="#eef2f7", font=("Arial",9),
                                      anchor="w", relief="groove", padx=6,
                                      justify="left")
        self.garantie_lbl.pack(side="left", fill="x", expand=True, ipady=3)

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
        ttk.Combobox(r, textvariable=self.type_var, values=db.get_types_intervention(),
                     width=44, state="readonly").pack(side="left")
        mk_btn(r, "⚙", lambda: TypesDialog(self, self.app), color=C["btn2"]).pack(side="left", padx=(8,0))

        # Marque moteur
        self.marque_var = tk.StringVar()
        r_m = tk.Frame(p, bg=C["bg"]); r_m.pack(fill="x", padx=20, pady=4)
        tk.Label(r_m, text="Marque moteur", bg=C["bg"], font=("Arial",10),
                 width=24, anchor="e").pack(side="left", padx=(0,8))
        self.marque_combo = ttk.Combobox(r_m, textvariable=self.marque_var,
                                          values=db.get_marques(), width=44,
                                          state="readonly")
        self.marque_combo.pack(side="left")
        def _open_marques_dlg():
            MarquesDialog(self, self.app, on_close=lambda: self.marque_combo.config(
                values=db.get_marques()))
        mk_btn(r_m, "⚙", _open_marques_dlg, color=C["btn2"]).pack(side="left", padx=(8,0))

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
            on_add_new=_open_new_tech)
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

        self._clients     = list(db.get_clients())
        self._all_moteurs = list(db.get_moteurs())
        self._techniciens = list(db.get_techniciens())
        self._client_by_id  = {c["id"]:  c for c in self._clients}
        self._client_by_nom = {c["nom"]: c for c in self._clients}
        self._moteur_by_id  = {m["id"]:  m for m in self._all_moteurs}
        self._moteur_by_ns  = {m["num_serie"]: m for m in self._all_moteurs}

        self.client_combo.set_values([c["nom"] for c in self._clients])
        self.tech_picker.set_available([t["nom"] for t in self._techniciens])
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
                self.signataire_var.set(row_get(client, "contact"))
            if not self.email_sig_var.get():
                self.email_sig_var.set(row_get(client, "email"))
        filtres = self._moteurs_filtres()
        series = [m["num_serie"] for m in filtres]
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
        # Afficher les garanties RÉELLES saisies dans l'app garanties
        # pour ce moteur (statut + attribution + N° EMS), au lieu de l'ancien
        # calcul automatique date de mise en service + durée.
        gars = db.get_garanties_moteur(row_get(m, "id"))
        if not gars:
            self.garantie_lbl.config(
                text="—  (aucune garantie enregistrée pour ce moteur)",
                fg="#666")
        else:
            ouvertes = [g for g in gars if g["statut"] != "Clôturée"]
            if ouvertes:
                parts = [
                    f"{g['num_ems']} · {g['attribution']} · {g['statut']}"
                    for g in ouvertes[:3]
                ]
                txt = "🛡 " + "   |   ".join(parts)
                if len(ouvertes) > 3:
                    txt += f"   (+{len(ouvertes) - 3} autre(s))"
                self.garantie_lbl.config(text=txt, fg="#0f5132")
            else:
                self.garantie_lbl.config(
                    text=f"🛡 {len(gars)} garantie(s) — toutes clôturées",
                    fg="#666")

    # ─── Moteurs supplémentaires ─────────────────────────────────────────────
    def _add_moteur_row(self, initial_ns=""):
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

        # Peuplement auto des infos à la sélection
        def _on_sel(*_, e=entry):
            m = self._moteur_by_ns.get(e["var"].get().strip())
            for key, lw in e["info_lbls"].items():
                lw.config(text=str(row_get(m, key, "") if m else ""))

        combo.bind("<<ComboboxSelected>>", _on_sel)

        # Si initial_ns fourni (chargement), peupler immédiatement
        if initial_ns:
            bloc.after(50, _on_sel)

    # ─── Fichiers ─────────────────────────────────────────────────────────────
    def _current_dossier(self):
        if not self.is_edit or not self.inv_id:
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
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
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
        # Collecter les emails de TOUS les techniciens assignés
        tech_names = db.parse_techniciens(row_get(inv, "technicien"))
        emails = []
        for n in tech_names:
            t = db.get_technicien_by_nom(n)
            em = row_get(t, "email")
            if em:
                emails.append(em)
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
        mailer.email_technicien(inv, client, moteur, emails, str(path))
        db.mark_notifie(self.inv_id, "tech")

    def _mail_cloture_inline(self):
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
        try:
            path = sauvegarder_bon(inv, generer_pdf=True)
        except (PermissionError, RuntimeError) as exc:
            messagebox.showerror("PDF impossible", str(exc)); return
        mailer.email_cloture(inv, client, moteur, tech_emails, str(path))

    # ─── Chargement ───────────────────────────────────────────────────────────
    def _load(self):
        inv = db.get_intervention(inv_id=self.inv_id)
        if not inv: return

        c = self._client_by_id.get(row_get(inv, "client_id"))
        if c:
            self.client_combo.set(c["nom"])
        self.lieu_var.set(row_get(inv, "lieu_intervention"))
        self.signataire_var.set(row_get(inv, "nom_signataire"))
        self.email_sig_var.set(row_get(inv, "email_signataire"))
        self.tel_sig_var.set(row_get(inv, "telephone_signataire"))
        self.demandeur_var.set(row_get(inv, "nom_demandeur"))
        self.email_demand_var.set(row_get(inv, "email_demandeur"))
        self.tel_demand_var.set(row_get(inv, "telephone_demandeur"))

        if c:
            filtrés = [m for m in self._all_moteurs if m["client_id"] == c["id"]]
        else:
            filtrés = self._all_moteurs
        self.moteur_combo.set_values([m["num_serie"] for m in filtrés])
        m = self._moteur_by_id.get(row_get(inv, "moteur_id"))
        if m:
            self.moteur_combo.set(m["num_serie"])
            self._set_moteur_info(m)
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
            self._add_moteur_row(initial_ns=em.get("num_serie", ""))

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
                    {k: row_get(m, k) for k in ("id","num_serie","navire","machine",
                                                  "type_moteur","marque","ref_constructeur",
                                                  "date_mise_service","duree_garantie")}
                    for e in self._extra_moteurs_rows
                    if (m := self._moteur_by_ns.get(e["var"].get().strip()))  # noqa: E231
                ],
                ensure_ascii=False),
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
    def __init__(self, parent, app, inv_id, on_save=None):
        super().__init__(parent)
        self.app = app
        self.inv_id = inv_id
        self.on_save = on_save

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

        # ── Récapitulatif (commun aux 2 étapes) ──────────────────────────
        mk_header(self, "Validation de l'intervention",
                  "   Vérifiez le récapitulatif puis signez ci-dessous")

        self._recap_widget()

        # ── Zone d'étape (remplacée à chaque étape) ──────────────────────
        self.zone = tk.Frame(self, bg=C["bg"])
        self.zone.pack(fill="both", expand=True, padx=24, pady=4)

        # ── Indicateur d'étape ──────────────────────────────────────────
        self.steps_bar = tk.Frame(self, bg=C["bg"])
        self.steps_bar.pack(fill="x", padx=24, pady=(0, 4))
        self._maj_indicateur()

        # ── Boutons (re-créés à chaque étape) ────────────────────────────
        self.bf = tk.Frame(self, bg=C["bg"])
        self.bf.pack(side="bottom", pady=(6, 16))

        self._afficher_etape_client()

    # ── Récapitulatif ─────────────────────────────────────────────────────
    def _recap_widget(self):
        client = db.get_client(self.inv["client_id"]) if self.inv["client_id"] else None
        moteur = db.get_moteur(self.inv["moteur_id"]) if self.inv["moteur_id"] else None
        recap = tk.Frame(self, bg=C["surface"], bd=1, relief="solid")
        recap.pack(fill="x", padx=24, pady=(12, 4))
        lignes = [
            ("N° de bon", self.inv["num_bon"]),
            ("Date", self.inv["date_creation"]),
            ("Client", (client["nom"] if client else "") or "—"),
            ("Navire / Site", row_get(moteur, "navire") or
                              row_get(self.inv, "lieu_intervention") or "—"),
            ("Moteur", row_get(moteur, "num_serie") or "—"),
            ("Type", self.inv["type_intervention"] or "—"),
            ("Technicien", self.inv["technicien"] or "—"),
        ]
        grid = tk.Frame(recap, bg=C["surface"])
        grid.pack(fill="x", padx=16, pady=10)
        for i, (lbl, val) in enumerate(lignes):
            r, c = divmod(i, 4)
            cell = tk.Frame(grid, bg=C["surface"])
            cell.grid(row=r, column=c, sticky="w", padx=14, pady=3)
            tk.Label(cell, text=lbl + " : ", bg=C["surface"],
                     fg=C["text_muted"],
                     font=("Segoe UI", 9, "bold")).pack(side="left")
            tk.Label(cell, text=str(val), bg=C["surface"], fg=C["text"],
                     font=("Segoe UI", 9)).pack(side="left")
        travaux = (row_get(self.inv, "travaux") or
                   row_get(self.inv, "description") or "").strip()
        if travaux:
            tk.Label(recap, text="Travaux : " + travaux[:280] +
                     ("…" if len(travaux) > 280 else ""),
                     bg=C["surface"], fg=C["text"],
                     font=("Segoe UI", 9), justify="left",
                     wraplength=1000, anchor="w").pack(
                fill="x", padx=16, pady=(0, 10))
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
            tk.Label(self.steps_bar, text=prefix + libelle,
                     bg=bg, fg=fg, font=("Segoe UI", 10, "bold"),
                     padx=14, pady=5).pack(side="left", padx=4)

    # ── ÉTAPE 1 : CLIENT ─────────────────────────────────────────────────
    def _afficher_etape_client(self):
        self._etape = 1
        self._maj_indicateur()
        for w in self.zone.winfo_children():
            w.destroy()
        for w in self.bf.winfo_children():
            w.destroy()

        tk.Label(self.zone, text="✍  SIGNATURE CLIENT",
                 bg=C["bg"], fg=C["header"],
                 font=("Segoe UI", 14, "bold")).pack(pady=(8, 2))
        tk.Label(self.zone,
                 text="À faire signer au client (au doigt sur la tablette)",
                 bg=C["bg"], fg=C["text_muted"],
                 font=("Segoe UI", 10)).pack(pady=(0, 4))

        # ── Case "client absent" ──────────────────────────────────────────
        self.client_absent_var = tk.IntVar(value=0)
        self._absent_chk = tk.Checkbutton(
            self.zone,
            text="  Client absent — passer directement à la signature technicien",
            variable=self.client_absent_var,
            bg=C["bg"], fg=C["warn"], font=("Segoe UI", 11, "bold"),
            activebackground=C["bg"], selectcolor="white",
            command=self._toggle_client_absent)
        self._absent_chk.pack(pady=(0, 6))

        # Badge affiché quand client absent est coché
        self._absent_badge = tk.Label(self.zone, text="⚠  CLIENT ABSENT — signature ignorée",
                                       bg="#fff3cd", fg="#856404",
                                       font=("Segoe UI", 12, "bold"), padx=12, pady=6,
                                       relief="solid", bd=1)
        # (non packé par défaut — apparaît via _toggle_client_absent)

        # Taille adaptée à la fenêtre plein écran (tablette)
        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
        except tk.TclError:
            sw, sh = 1600, 900
        pad_w = max(800, min(sw - 200, 1600))
        pad_h = max(280, min(sh - 480, 480))
        self.pad_client = _SignaturePad(self.zone, width=pad_w, height=pad_h,
                                         titre="", sous_titre="")
        self.pad_client.pack()

        nf = tk.Frame(self.zone, bg=C["bg"])
        nf.pack(pady=(10, 4))
        tk.Label(nf, text="Nom du signataire * : ", bg=C["bg"],
                 font=("Segoe UI", 11), fg=C["text"]).pack(side="left")
        self.nom_client_var = tk.StringVar(value=self._client_nom_defaut)
        self._nom_client_entry = ttk.Entry(nf, textvariable=self.nom_client_var,
                                            width=40, font=("Segoe UI", 12))
        self._nom_client_entry.pack(side="left")

        self.accept_var = tk.IntVar(value=0)
        self._accept_chk = tk.Checkbutton(
            self.zone,
            text=" Le client reconnaît avoir pris connaissance de "
                 "l'intervention et en accepte la réalisation.",
            variable=self.accept_var, bg=C["bg"], fg=C["text"],
            font=("Segoe UI", 10), activebackground=C["bg"],
            selectcolor="white", anchor="w")
        self._accept_chk.pack(pady=4)

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

        tk.Label(self.zone, text="✍  SIGNATURE TECHNICIEN EMS",
                 bg=C["bg"], fg=C["header"],
                 font=("Segoe UI", 14, "bold")).pack(pady=(8, 2))
        tk.Label(self.zone,
                 text="Le technicien atteste la réalisation des travaux",
                 bg=C["bg"], fg=C["text_muted"],
                 font=("Segoe UI", 10)).pack(pady=(0, 8))

        try:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
        except tk.TclError:
            sw, sh = 1600, 900
        pad_w = max(800, min(sw - 200, 1600))
        pad_h = max(280, min(sh - 480, 480))
        self.pad_tech = _SignaturePad(self.zone, width=pad_w, height=pad_h,
                                       titre="", sous_titre="")
        self.pad_tech.pack()

        nf = tk.Frame(self.zone, bg=C["bg"])
        nf.pack(pady=(10, 4))
        tk.Label(nf, text="Nom du technicien * : ", bg=C["bg"],
                 font=("Segoe UI", 11), fg=C["text"]).pack(side="left")
        self.nom_tech_var = tk.StringVar(
            value=row_get(self.inv, "technicien") or "")
        ttk.Entry(nf, textvariable=self.nom_tech_var, width=40,
                  font=("Segoe UI", 12)).pack(side="left")

        tk.Label(self.zone,
                 text=" J'atteste sur l'honneur la réalisation des travaux "
                      "décrits ci-dessus.",
                 bg=C["bg"], fg=C["text_muted"],
                 font=("Segoe UI", 10, "italic")).pack(pady=4)

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

        # Enregistrement des 2 signatures + régénération bon
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
