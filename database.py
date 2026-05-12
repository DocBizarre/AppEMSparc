"""
EMS – Base de données SQLite (v1.8)
- Modèle conforme au bon EMS officiel
- Critère d'urgence : Normale / Urgente / Critique
- Classifications : Garantie / Facturable / Interne (3 cases indépendantes)
- Timestamps en heure de Paris (Europe/Paris)
- Configuration persistante du tableau de bord
"""

import json
import sqlite3
from pathlib import Path
from datetime import datetime, timezone, timedelta, date as _date

DB_PATH = Path(__file__).parent / "data" / "ems.db"
DOSSIERS_PATH = Path(__file__).parent / "dossiers"

# ─── Constantes ───────────────────────────────────────────────────────────────
URGENCES = ["Normale", "Urgente", "Critique"]
URGENCE_DEFAULT = "Normale"


# ─── Heure de Paris ───────────────────────────────────────────────────────────
def _paris_tz():
    """
    Retourne le fuseau Europe/Paris.
    Tente zoneinfo (Python 3.9+, recommandé), sinon fallback offset fixe.
    """
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo("Europe/Paris")
    except (ImportError, Exception):
        # Fallback : approximation grossière (UTC+1, ne gère pas DST)
        # Mieux que rien si le système n'a pas tzdata
        return timezone(timedelta(hours=1), name="CET")


PARIS = _paris_tz()


def now_paris_iso():
    """Retourne la date/heure de Paris au format ISO 'YYYY-MM-DD HH:MM:SS'."""
    return datetime.now(PARIS).strftime("%Y-%m-%d %H:%M:%S")


def utc_to_paris(s):
    """
    Convertit une chaîne datetime UTC SQLite ('YYYY-MM-DD HH:MM:SS')
    en chaîne heure de Paris au même format.
    Retourne s tel quel si le parsing échoue.
    """
    if not s:
        return ""
    s = str(s).strip()
    # Si déjà au format Paris (sortie de now_paris_iso), pas de conversion
    # Stratégie : on suppose que les timestamps SQLite via datetime('now') sont en UTC
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            dt = datetime.strptime(s, fmt)
            dt = dt.replace(tzinfo=timezone.utc).astimezone(PARIS)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return s


def fmt_paris_short(s):
    """Format court 'JJ/MM HH:MM' pour affichage Treeview."""
    s = utc_to_paris(s)
    try:
        dt = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d/%m %H:%M")
    except (ValueError, TypeError):
        return s[:16] if s else ""


# ─── Connexion ────────────────────────────────────────────────────────────────
def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ─── Init / Migration ─────────────────────────────────────────────────────────
def init_db():
    DOSSIERS_PATH.mkdir(parents=True, exist_ok=True)
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS clients (
            id          TEXT PRIMARY KEY,
            nom         TEXT NOT NULL,
            contact     TEXT DEFAULT '',
            email       TEXT DEFAULT '',
            telephone   TEXT DEFAULT '',
            adresse     TEXT DEFAULT '',
            created_at  TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS moteurs (
            id                  TEXT PRIMARY KEY,
            client_id           TEXT REFERENCES clients(id) ON DELETE SET NULL,
            num_serie           TEXT NOT NULL UNIQUE,
            navire              TEXT DEFAULT '',
            machine             TEXT DEFAULT '',
            type_moteur         TEXT DEFAULT '',
            date_mise_service   TEXT DEFAULT '',
            duree_garantie      TEXT DEFAULT '',
            -- Champs étendus (import CSV / parc)
            cylindree           TEXT DEFAULT '',
            famille             TEXT DEFAULT '',
            marque              TEXT DEFAULT '',
            application         TEXT DEFAULT '',
            typologie           TEXT DEFAULT '',
            collection          TEXT DEFAULT '',
            ref_constructeur    TEXT DEFAULT '',
            code_affaire        TEXT DEFAULT '',
            type_client         TEXT DEFAULT '',
            created_at          TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS interventions (
            id                  TEXT PRIMARY KEY,
            num_bon             TEXT NOT NULL UNIQUE,
            client_id           TEXT REFERENCES clients(id) ON DELETE SET NULL,
            moteur_id           TEXT REFERENCES moteurs(id) ON DELETE SET NULL,

            type_intervention   TEXT DEFAULT '',
            urgence             TEXT DEFAULT 'Normale',

            -- Classifications (3 cases indépendantes)
            garantie_intervention INTEGER DEFAULT 0,
            facturable           INTEGER DEFAULT 0,
            interne              INTEGER DEFAULT 0,

            technicien          TEXT DEFAULT '',
            date_creation       TEXT DEFAULT '',
            statut              TEXT DEFAULT 'En cours',

            lieu_intervention   TEXT DEFAULT '',
            nom_signataire      TEXT DEFAULT '',
            email_signataire    TEXT DEFAULT '',
            nb_heures_fct       TEXT DEFAULT '',

            outil_diagnostic    INTEGER DEFAULT 0,
            memoriser_avant     INTEGER DEFAULT 0,
            memoriser_apres     INTEGER DEFAULT 0,
            photos_avant        INTEGER DEFAULT 0,
            photos_apres        INTEGER DEFAULT 0,
            pour_information    INTEGER DEFAULT 0,
            preconisation       INTEGER DEFAULT 0,

            demande_client      TEXT DEFAULT '',
            constat             TEXT DEFAULT '',
            travaux             TEXT DEFAULT '',
            informations        TEXT DEFAULT '',

            description         TEXT DEFAULT '',
            pieces              TEXT DEFAULT '',

            materiels_json      TEXT DEFAULT '[]',
            deplacements_json   TEXT DEFAULT '{}',

            dossier_path        TEXT DEFAULT '',
            client_notifie      INTEGER DEFAULT 0,
            tech_notifie        INTEGER DEFAULT 0,
            created_at          TEXT DEFAULT (datetime('now')),
            updated_at          TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS techniciens (
            id          TEXT PRIMARY KEY,
            nom         TEXT NOT NULL UNIQUE,
            email       TEXT DEFAULT '',
            telephone   TEXT DEFAULT '',
            created_at  TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS types_intervention (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            libelle TEXT NOT NULL UNIQUE,
            ordre   INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS config (
            cle     TEXT PRIMARY KEY,
            valeur  TEXT
        );

        INSERT OR IGNORE INTO config VALUES ('compteur_bon', '1');
        """)

        # Migration douce — toutes les colonnes nouvelles
        new_cols = [
            ("interventions", "client_notifie",        "INTEGER DEFAULT 0"),
            ("interventions", "tech_notifie",          "INTEGER DEFAULT 0"),
            ("interventions", "lieu_intervention",     "TEXT DEFAULT ''"),
            ("interventions", "nom_signataire",        "TEXT DEFAULT ''"),
            ("interventions", "email_signataire",      "TEXT DEFAULT ''"),
            ("interventions", "nb_heures_fct",         "TEXT DEFAULT ''"),
            ("interventions", "outil_diagnostic",      "INTEGER DEFAULT 0"),
            ("interventions", "memoriser_avant",       "INTEGER DEFAULT 0"),
            ("interventions", "memoriser_apres",       "INTEGER DEFAULT 0"),
            ("interventions", "photos_avant",          "INTEGER DEFAULT 0"),
            ("interventions", "photos_apres",          "INTEGER DEFAULT 0"),
            ("interventions", "pour_information",      "INTEGER DEFAULT 0"),
            ("interventions", "preconisation",         "INTEGER DEFAULT 0"),
            ("interventions", "demande_client",        "TEXT DEFAULT ''"),
            ("interventions", "constat",               "TEXT DEFAULT ''"),
            ("interventions", "informations",          "TEXT DEFAULT ''"),
            ("interventions", "materiels_json",        "TEXT DEFAULT '[]'"),
            ("interventions", "deplacements_json",     "TEXT DEFAULT '{}'"),
            ("interventions", "urgence",               "TEXT DEFAULT 'Normale'"),
            ("interventions", "garantie_intervention", "INTEGER DEFAULT 0"),
            ("interventions", "facturable",            "INTEGER DEFAULT 0"),
            ("interventions", "interne",               "INTEGER DEFAULT 0"),
            ("moteurs",       "type_moteur",           "TEXT DEFAULT ''"),
            ("moteurs",       "cylindree",             "TEXT DEFAULT ''"),
            ("moteurs",       "famille",               "TEXT DEFAULT ''"),
            ("moteurs",       "marque",                "TEXT DEFAULT ''"),
            ("moteurs",       "application",           "TEXT DEFAULT ''"),
            ("moteurs",       "typologie",             "TEXT DEFAULT ''"),
            ("moteurs",       "collection",            "TEXT DEFAULT ''"),
            ("moteurs",       "ref_constructeur",      "TEXT DEFAULT ''"),
            ("moteurs",       "code_affaire",          "TEXT DEFAULT ''"),
            ("moteurs",       "type_client",           "TEXT DEFAULT ''"),
        ]
        for table, col, ddl in new_cols:
            try:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")
            except sqlite3.OperationalError:
                pass

        # Seed types d'intervention
        n = conn.execute("SELECT COUNT(*) FROM types_intervention").fetchone()[0]
        if n == 0:
            ancien = conn.execute(
                "SELECT valeur FROM config WHERE cle='types_intervention'"
            ).fetchone()
            if ancien and ancien["valeur"]:
                items = [s.strip() for s in ancien["valeur"].split("|") if s.strip()]
            else:
                items = ["Entretien", "Dépannage", "Diagnostic", "Garantie",
                         "Maintenance préventive", "Mise en service",
                         "Expertise", "Révision", "Autre"]
            for i, lib in enumerate(items):
                conn.execute(
                    "INSERT OR IGNORE INTO types_intervention (libelle, ordre) VALUES (?, ?)",
                    (lib, i))

        # Config par défaut du tableau de bord
        if conn.execute("SELECT 1 FROM config WHERE cle='dashboard_widgets'").fetchone() is None:
            default_widgets = json.dumps([
                "stats_cards", "urgentes", "activite_recente",
                "garantie_expirante", "par_technicien", "non_notifies"
            ])
            conn.execute("INSERT INTO config VALUES ('dashboard_widgets', ?)",
                         (default_widgets,))
        if conn.execute("SELECT 1 FROM config WHERE cle='dashboard_cards'").fetchone() is None:
            default_cards = json.dumps([
                "En cours", "Clos", "Facturé", "Total",
                "Urgentes", "Clients", "Moteurs", "Tech."
            ])
            conn.execute("INSERT INTO config VALUES ('dashboard_cards', ?)",
                         (default_cards,))
        conn.commit()


# ─── CONFIG ───────────────────────────────────────────────────────────────────
def get_config(cle):
    with get_conn() as conn:
        row = conn.execute("SELECT valeur FROM config WHERE cle=?", (cle,)).fetchone()
        return row["valeur"] if row else None


def set_config(cle, valeur):
    conn = get_conn()
    try:
        conn.execute("INSERT OR REPLACE INTO config VALUES (?,?)", (cle, str(valeur)))
        conn.commit()
    finally:
        conn.close()


def get_dashboard_widgets():
    """Retourne la liste des widgets actifs du tableau de bord."""
    s = get_config("dashboard_widgets")
    try:
        return json.loads(s) if s else []
    except (json.JSONDecodeError, TypeError):
        return []


def set_dashboard_widgets(widgets):
    set_config("dashboard_widgets", json.dumps(list(widgets)))


def get_dashboard_cards():
    """Retourne la liste des cartes statistiques actives."""
    s = get_config("dashboard_cards")
    try:
        return json.loads(s) if s else []
    except (json.JSONDecodeError, TypeError):
        return []


def set_dashboard_cards(cards):
    set_config("dashboard_cards", json.dumps(list(cards)))


# ─── TYPES D'INTERVENTION ─────────────────────────────────────────────────────
def get_types_intervention():
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT libelle FROM types_intervention ORDER BY ordre, libelle"
        ).fetchall()
        return [r["libelle"] for r in rows]


def add_type_intervention(libelle):
    libelle = (libelle or "").strip()
    if not libelle:
        return False
    with get_conn() as conn:
        max_ordre = conn.execute(
            "SELECT COALESCE(MAX(ordre),0) FROM types_intervention"
        ).fetchone()[0]
        try:
            conn.execute(
                "INSERT INTO types_intervention (libelle, ordre) VALUES (?, ?)",
                (libelle, max_ordre + 1))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False


def delete_type_intervention(libelle):
    with get_conn() as conn:
        conn.execute("DELETE FROM types_intervention WHERE libelle=?", (libelle,))
        conn.commit()


def update_type_intervention(ancien, nouveau):
    nouveau = (nouveau or "").strip()
    if not nouveau:
        return False
    with get_conn() as conn:
        try:
            conn.execute("UPDATE types_intervention SET libelle=? WHERE libelle=?",
                         (nouveau, ancien))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False


# ─── CLIENTS ──────────────────────────────────────────────────────────────────
def get_clients(search=""):
    with get_conn() as conn:
        if search:
            s = f"%{search}%"
            return conn.execute(
                "SELECT * FROM clients WHERE nom LIKE ? OR contact LIKE ? OR email LIKE ? "
                "OR telephone LIKE ? ORDER BY nom", (s, s, s, s)
            ).fetchall()
        return conn.execute("SELECT * FROM clients ORDER BY nom").fetchall()


def get_client(client_id):
    with get_conn() as conn:
        return conn.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()


def upsert_client(data, client_id=None):
    cid = client_id or f"C{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO clients (id,nom,contact,email,telephone,adresse)
            VALUES (:id,:nom,:contact,:email,:telephone,:adresse)
            ON CONFLICT(id) DO UPDATE SET
                nom=excluded.nom, contact=excluded.contact,
                email=excluded.email, telephone=excluded.telephone,
                adresse=excluded.adresse
        """, {"id": cid, **data})
        conn.commit()
    finally:
        conn.close()
    return cid


def delete_client(client_id):
    conn = get_conn()
    try:
        conn.execute("DELETE FROM clients WHERE id=?", (client_id,))
        conn.commit()
    finally:
        conn.close()


# ─── MOTEURS ──────────────────────────────────────────────────────────────────
def get_moteurs(search="", client_id=None):
    with get_conn() as conn:
        q = "SELECT m.*, c.nom as client_nom FROM moteurs m LEFT JOIN clients c ON m.client_id=c.id"
        params = []
        clauses = []
        if client_id:
            clauses.append("m.client_id=?")
            params.append(client_id)
        if search:
            clauses.append("(m.num_serie LIKE ? OR m.navire LIKE ? OR m.machine LIKE ? OR m.type_moteur LIKE ?)")
            s = f"%{search}%"
            params += [s, s, s, s]
        if clauses:
            q += " WHERE " + " AND ".join(clauses)
        q += " ORDER BY m.num_serie"
        return conn.execute(q, params).fetchall()


def get_moteur(moteur_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT m.*, c.nom as client_nom FROM moteurs m "
            "LEFT JOIN clients c ON m.client_id=c.id WHERE m.id=?", (moteur_id,)
        ).fetchone()


def find_moteur_by_serie(num_serie):
    with get_conn() as conn:
        return conn.execute(
            "SELECT m.*, c.nom as client_nom FROM moteurs m "
            "LEFT JOIN clients c ON m.client_id=c.id WHERE m.num_serie=?", (num_serie,)
        ).fetchone()


def find_client_by_nom(nom):
    """Recherche exacte (insensible à la casse / aux espaces) par nom client."""
    if not nom:
        return None
    nom_norm = nom.strip().lower()
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM clients WHERE LOWER(TRIM(nom))=?", (nom_norm,)
        ).fetchone()


MOTEUR_EXTRA_FIELDS = [
    "cylindree", "famille", "marque", "application", "typologie",
    "collection", "ref_constructeur", "code_affaire", "type_client",
]


def upsert_moteur(data, moteur_id=None):
    mid = moteur_id or f"M{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    # Valeurs par défaut pour tous les champs (compat anciens callers)
    defaults = {"type_moteur": "", "navire": "", "machine": "",
                "date_mise_service": "", "duree_garantie": "", "client_id": ""}
    for f in MOTEUR_EXTRA_FIELDS:
        defaults[f] = ""
    data = {**defaults, **data}
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO moteurs (id,client_id,num_serie,navire,machine,type_moteur,
                                 date_mise_service,duree_garantie,
                                 cylindree,famille,marque,application,typologie,
                                 collection,ref_constructeur,code_affaire,type_client)
            VALUES (:id,:client_id,:num_serie,:navire,:machine,:type_moteur,
                    :date_mise_service,:duree_garantie,
                    :cylindree,:famille,:marque,:application,:typologie,
                    :collection,:ref_constructeur,:code_affaire,:type_client)
            ON CONFLICT(id) DO UPDATE SET
                client_id=excluded.client_id, num_serie=excluded.num_serie,
                navire=excluded.navire, machine=excluded.machine,
                type_moteur=excluded.type_moteur,
                date_mise_service=excluded.date_mise_service,
                duree_garantie=excluded.duree_garantie,
                cylindree=excluded.cylindree, famille=excluded.famille,
                marque=excluded.marque, application=excluded.application,
                typologie=excluded.typologie, collection=excluded.collection,
                ref_constructeur=excluded.ref_constructeur,
                code_affaire=excluded.code_affaire,
                type_client=excluded.type_client
        """, {"id": mid, **data})
        conn.commit()
    finally:
        conn.close()
    return mid


def delete_moteur(moteur_id):
    conn = get_conn()
    try:
        conn.execute("DELETE FROM moteurs WHERE id=?", (moteur_id,))
        conn.commit()
    finally:
        conn.close()


def get_moteurs_garantie_expirante(jours_max=90):
    """Retourne les moteurs dont la garantie expire dans <= jours_max jours (et > 0)."""
    moteurs = get_moteurs()
    out = []
    for m in moteurs:
        statut, jours = garantie_status(m["date_mise_service"], m["duree_garantie"])
        if statut == "Active" and jours is not None and jours <= jours_max:
            out.append({"moteur": m, "jours_restants": jours})
    out.sort(key=lambda x: x["jours_restants"])
    return out


# ─── TECHNICIENS ──────────────────────────────────────────────────────────────
def get_techniciens():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM techniciens ORDER BY nom").fetchall()


def get_technicien_by_nom(nom):
    if not nom:
        return None
    with get_conn() as conn:
        return conn.execute("SELECT * FROM techniciens WHERE nom=?", (nom,)).fetchone()


def upsert_technicien(data, tech_id=None):
    tid = tech_id or f"T{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO techniciens (id,nom,email,telephone)
            VALUES (:id,:nom,:email,:telephone)
            ON CONFLICT(id) DO UPDATE SET
                nom=excluded.nom, email=excluded.email, telephone=excluded.telephone
        """, {"id": tid, **data})
        conn.commit()
    finally:
        conn.close()
    return tid


def delete_technicien(tech_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM techniciens WHERE id=?", (tech_id,))
        conn.commit()


# ─── INTERVENTIONS ────────────────────────────────────────────────────────────
def next_num_bon():
    n = int(get_config("compteur_bon") or 1)
    set_config("compteur_bon", n + 1)
    return f"BON-{datetime.now(PARIS).year}-{n:04d}"


INTERVENTION_FIELDS = [
    "client_id", "moteur_id", "type_intervention", "urgence",
    "garantie_intervention", "facturable", "interne",
    "technicien", "date_creation", "statut",
    "lieu_intervention", "nom_signataire", "email_signataire",
    "nb_heures_fct",
    "outil_diagnostic", "memoriser_avant", "memoriser_apres",
    "photos_avant", "photos_apres", "pour_information", "preconisation",
    "demande_client", "constat", "travaux", "informations",
    "description", "pieces",
    "materiels_json", "deplacements_json",
]

_INT_FIELDS = {
    "garantie_intervention", "facturable", "interne",
    "outil_diagnostic", "memoriser_avant", "memoriser_apres",
    "photos_avant", "photos_apres", "pour_information", "preconisation",
}


def _normalize_intervention_data(data):
    defaults = {f: "" for f in INTERVENTION_FIELDS}
    for f in _INT_FIELDS:
        defaults[f] = 0
    defaults["statut"]            = "En cours"
    defaults["urgence"]           = URGENCE_DEFAULT
    defaults["materiels_json"]    = "[]"
    defaults["deplacements_json"] = "{}"
    out = {**defaults, **data}
    # Sérialiser
    if not isinstance(out["materiels_json"], str):
        out["materiels_json"] = json.dumps(out["materiels_json"], ensure_ascii=False)
    if not isinstance(out["deplacements_json"], str):
        out["deplacements_json"] = json.dumps(out["deplacements_json"], ensure_ascii=False)
    # Compat
    if out["demande_client"] == "" and out["description"]:
        out["demande_client"] = out["description"]
    if out["description"] == "" and out["demande_client"]:
        out["description"] = out["demande_client"]
    # Cast int proprement
    for f in _INT_FIELDS:
        try:
            out[f] = int(out[f] or 0)
        except (ValueError, TypeError):
            out[f] = 0
    # Urgence : valeur valide ou défaut
    if out["urgence"] not in URGENCES:
        out["urgence"] = URGENCE_DEFAULT
    return out


def get_interventions(statut=None, search="", urgence=None):
    with get_conn() as conn:
        q = """
            SELECT i.*, c.nom as client_nom, c.email as client_email,
                   c.telephone as client_tel,
                   m.num_serie, m.navire, m.machine, m.type_moteur,
                   m.marque, m.ref_constructeur, m.cylindree, m.famille,
                   m.application, m.typologie, m.collection, m.code_affaire,
                   m.date_mise_service, m.duree_garantie
            FROM interventions i
            LEFT JOIN clients c ON i.client_id=c.id
            LEFT JOIN moteurs m ON i.moteur_id=m.id
        """
        params = []
        clauses = []
        if statut and statut != "Tous":
            clauses.append("i.statut=?")
            params.append(statut)
        if urgence and urgence != "Toutes":
            clauses.append("i.urgence=?")
            params.append(urgence)
        if search:
            s = f"%{search}%"
            clauses.append(
                "(i.num_bon LIKE ? OR c.nom LIKE ? OR m.num_serie LIKE ? "
                "OR m.navire LIKE ? OR i.technicien LIKE ? OR i.type_intervention LIKE ?)"
            )
            params += [s, s, s, s, s, s]
        if clauses:
            q += " WHERE " + " AND ".join(clauses)
        # Tri : urgence puis date
        q += """ ORDER BY
            CASE i.urgence
                WHEN 'Critique' THEN 1
                WHEN 'Urgente'  THEN 2
                WHEN 'Normale'  THEN 3
                ELSE 4
            END,
            i.created_at DESC"""
        return conn.execute(q, params).fetchall()


def get_intervention(inv_id=None, num_bon=None):
    with get_conn() as conn:
        sql = ("SELECT i.*, c.nom as client_nom, c.email as client_email, "
               "c.contact as client_contact, c.telephone as client_tel, "
               "c.adresse as client_adresse, "
               "m.num_serie, m.navire, m.machine, m.type_moteur, "
               "m.marque, m.ref_constructeur, m.cylindree, m.famille, "
               "m.application, m.typologie, m.collection, m.code_affaire, "
               "m.date_mise_service, m.duree_garantie "
               "FROM interventions i "
               "LEFT JOIN clients c ON i.client_id=c.id "
               "LEFT JOIN moteurs m ON i.moteur_id=m.id ")
        if inv_id:
            return conn.execute(sql + "WHERE i.id=?", (inv_id,)).fetchone()
        if num_bon:
            return conn.execute(sql + "WHERE i.num_bon=?", (num_bon,)).fetchone()


def get_interventions_for_moteur(moteur_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT i.*, c.nom as client_nom, m.num_serie "
            "FROM interventions i "
            "LEFT JOIN clients c ON i.client_id=c.id "
            "LEFT JOIN moteurs m ON i.moteur_id=m.id "
            "WHERE i.moteur_id=? ORDER BY i.created_at DESC", (moteur_id,)
        ).fetchall()


def create_intervention(data):
    iid = f"I{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    num_bon = next_num_bon()
    dossier = DOSSIERS_PATH / num_bon
    dossier.mkdir(parents=True, exist_ok=True)
    d = _normalize_intervention_data(data)
    d["id"] = iid
    d["num_bon"] = num_bon
    d["dossier_path"] = str(dossier)
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO interventions (
                id, num_bon, client_id, moteur_id, type_intervention, urgence,
                garantie_intervention, facturable, interne,
                technicien, date_creation, statut,
                lieu_intervention, nom_signataire, email_signataire, nb_heures_fct,
                outil_diagnostic, memoriser_avant, memoriser_apres,
                photos_avant, photos_apres, pour_information, preconisation,
                demande_client, constat, travaux, informations,
                description, pieces,
                materiels_json, deplacements_json, dossier_path
            ) VALUES (
                :id, :num_bon, :client_id, :moteur_id, :type_intervention, :urgence,
                :garantie_intervention, :facturable, :interne,
                :technicien, :date_creation, :statut,
                :lieu_intervention, :nom_signataire, :email_signataire, :nb_heures_fct,
                :outil_diagnostic, :memoriser_avant, :memoriser_apres,
                :photos_avant, :photos_apres, :pour_information, :preconisation,
                :demande_client, :constat, :travaux, :informations,
                :description, :pieces,
                :materiels_json, :deplacements_json, :dossier_path
            )
        """, d)
        conn.commit()
    finally:
        conn.close()
    return iid, num_bon


def update_intervention(inv_id, data):
    d = _normalize_intervention_data(data)
    d["id"] = inv_id
    conn = get_conn()
    try:
        conn.execute("""
            UPDATE interventions SET
                client_id=:client_id, moteur_id=:moteur_id,
                type_intervention=:type_intervention, urgence=:urgence,
                garantie_intervention=:garantie_intervention,
                facturable=:facturable, interne=:interne,
                technicien=:technicien,
                date_creation=:date_creation, statut=:statut,
                lieu_intervention=:lieu_intervention,
                nom_signataire=:nom_signataire, email_signataire=:email_signataire,
                nb_heures_fct=:nb_heures_fct,
                outil_diagnostic=:outil_diagnostic,
                memoriser_avant=:memoriser_avant, memoriser_apres=:memoriser_apres,
                photos_avant=:photos_avant, photos_apres=:photos_apres,
                pour_information=:pour_information, preconisation=:preconisation,
                demande_client=:demande_client, constat=:constat,
                travaux=:travaux, informations=:informations,
                description=:description, pieces=:pieces,
                materiels_json=:materiels_json, deplacements_json=:deplacements_json,
                updated_at=datetime('now')
            WHERE id=:id
        """, d)
        conn.commit()
    finally:
        conn.close()


def mark_notifie(inv_id, kind):
    col = "client_notifie" if kind == "client" else "tech_notifie"
    with get_conn() as conn:
        conn.execute(f"UPDATE interventions SET {col}=1, updated_at=datetime('now') WHERE id=?",
                     (inv_id,))
        conn.commit()


def delete_intervention(inv_id):
    conn = get_conn()
    try:
        conn.execute("DELETE FROM interventions WHERE id=?", (inv_id,))
        conn.commit()
    finally:
        conn.close()


# ─── STATS ────────────────────────────────────────────────────────────────────
def get_stats():
    """Retourne un dict de toutes les statistiques utilisables par les cartes."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT statut, COUNT(*) as n FROM interventions GROUP BY statut"
        ).fetchall()
        stats = {"En cours": 0, "Clos": 0, "Facturé": 0, "Total": 0}
        for r in rows:
            if r["statut"] in stats:
                stats[r["statut"]] = r["n"]
            stats["Total"] += r["n"]

        # Par urgence (sur "En cours" uniquement, plus pertinent)
        urg = conn.execute(
            "SELECT urgence, COUNT(*) as n FROM interventions "
            "WHERE statut='En cours' GROUP BY urgence"
        ).fetchall()
        stats["Urgentes"]  = 0
        stats["Critiques"] = 0
        for r in urg:
            if r["urgence"] == "Urgente":  stats["Urgentes"]  = r["n"]
            if r["urgence"] == "Critique": stats["Critiques"] = r["n"]

        # Classifications
        for col, key in [("garantie_intervention", "Garantie"),
                          ("facturable", "Facturables"),
                          ("interne", "Internes")]:
            stats[key] = conn.execute(
                f"SELECT COUNT(*) FROM interventions WHERE {col}=1"
            ).fetchone()[0]

        # Notifications manquantes
        stats["Non notifiés"] = conn.execute(
            "SELECT COUNT(*) FROM interventions "
            "WHERE statut='En cours' AND (client_notifie=0 OR tech_notifie=0)"
        ).fetchone()[0]

        stats["clients"]     = conn.execute("SELECT COUNT(*) FROM clients").fetchone()[0]
        stats["moteurs"]     = conn.execute("SELECT COUNT(*) FROM moteurs").fetchone()[0]
        stats["techniciens"] = conn.execute("SELECT COUNT(*) FROM techniciens").fetchone()[0]
        return stats


def get_stats_par_technicien():
    """Pour chaque technicien : nb d'interventions En cours, Clos, Facturé."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT technicien,
                   SUM(CASE WHEN statut='En cours' THEN 1 ELSE 0 END) as en_cours,
                   SUM(CASE WHEN statut='Clos'     THEN 1 ELSE 0 END) as clos,
                   SUM(CASE WHEN statut='Facturé'  THEN 1 ELSE 0 END) as facture,
                   COUNT(*) as total
            FROM interventions
            WHERE technicien <> ''
            GROUP BY technicien
            ORDER BY total DESC, technicien
        """).fetchall()
        return [dict(r) for r in rows]


def get_stats_par_type():
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT type_intervention as type,
                   COUNT(*) as total,
                   SUM(CASE WHEN statut='En cours' THEN 1 ELSE 0 END) as en_cours
            FROM interventions
            WHERE type_intervention <> ''
            GROUP BY type_intervention
            ORDER BY total DESC
        """).fetchall()
        return [dict(r) for r in rows]


def get_interventions_urgentes(limit=10):
    """Critique + Urgente, en statut En cours, triées par urgence puis date."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT i.*, c.nom as client_nom, m.num_serie, m.navire
            FROM interventions i
            LEFT JOIN clients c ON i.client_id=c.id
            LEFT JOIN moteurs m ON i.moteur_id=m.id
            WHERE i.statut='En cours' AND i.urgence IN ('Urgente','Critique')
            ORDER BY CASE i.urgence WHEN 'Critique' THEN 1 ELSE 2 END,
                     i.created_at DESC
            LIMIT ?
        """, (limit,)).fetchall()


def get_activite_recente(limit=10):
    """Derniers bons modifiés."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT i.*, c.nom as client_nom, m.num_serie
            FROM interventions i
            LEFT JOIN clients c ON i.client_id=c.id
            LEFT JOIN moteurs m ON i.moteur_id=m.id
            ORDER BY i.updated_at DESC
            LIMIT ?
        """, (limit,)).fetchall()


def get_non_notifies(limit=10):
    """Bons En cours sans notification client OU technicien."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT i.*, c.nom as client_nom, m.num_serie
            FROM interventions i
            LEFT JOIN clients c ON i.client_id=c.id
            LEFT JOIN moteurs m ON i.moteur_id=m.id
            WHERE i.statut='En cours' AND (i.client_notifie=0 OR i.tech_notifie=0)
            ORDER BY i.created_at DESC
            LIMIT ?
        """, (limit,)).fetchall()


# ─── GARANTIE ─────────────────────────────────────────────────────────────────
def garantie_status(date_mise_service, duree_garantie_mois):
    s = (date_mise_service or "").strip()
    try:
        d = datetime.strptime(s, "%d/%m/%Y").date()
    except (ValueError, TypeError):
        return ("—", None)
    try:
        mois = int(str(duree_garantie_mois).strip() or 0)
    except (ValueError, TypeError):
        return ("—", None)
    if mois <= 0:
        return ("—", None)
    annee_fin = d.year + (d.month - 1 + mois) // 12
    mois_fin  = (d.month - 1 + mois) % 12 + 1
    bissextile = (annee_fin % 4 == 0 and (annee_fin % 100 != 0 or annee_fin % 400 == 0))
    days_in_month = [31, 29 if bissextile else 28, 31, 30, 31, 30,
                     31, 31, 30, 31, 30, 31]
    jour_fin = min(d.day, days_in_month[mois_fin - 1])
    fin = _date(annee_fin, mois_fin, jour_fin)
    today = datetime.now(PARIS).date()
    diff = (fin - today).days
    if diff >= 0:
        return ("Active", diff)
    return ("Expirée", -diff)


# ─── EMAIL VALIDATION (informative, non bloquante) ────────────────────────────
def email_looks_valid(s):
    """
    Validation d'email TRÈS PERMISSIVE et purement informative.
    Doit contenir un @ et au moins un . après. Pas de regex stricte —
    l'utilisateur reste libre de saisir ce qu'il veut.
    """
    if not s:
        return True  # vide = valide (pas obligatoire)
    s = s.strip()
    if "@" not in s:
        return False
    local, _, domain = s.rpartition("@")
    return bool(local) and "." in domain
