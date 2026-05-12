"""
EMS – Import CSV du parc clients & moteurs
Lit un fichier CSV, mappe les colonnes vers les champs DB, et déduplique par N° série.
"""

import csv
import io
from pathlib import Path

import database as db


# Mapping par défaut : nom de colonne CSV → champ DB
# Plusieurs alias possibles par champ (essai dans l'ordre, insensible casse/espaces)
DEFAULT_MAPPING = {
    # Client
    "nom_client":        ["Tiers", "Client", "Nom client", "Société"],
    "type_client":       ["Type Client", "Type client", "Catégorie client"],
    # Moteur
    "navire":            ["Machine / Engin", "Machine/Engin", "Machine", "Navire", "Engin"],
    "num_serie":         ["N° Série", "N° Serie", "Numéro de série", "Num série", "N°Série"],
    "num_moteur":        ["N° Moteur", "Numéro moteur"],
    "cylindree":         ["Cylindrée", "Cylindree", "Cyl."],
    "famille":           ["Famille"],
    "ref_constructeur":  ["Ref Constructeur", "Réf. Constructeur", "Ref constructeur"],
    "application":       ["Application"],
    "type_moteur":       ["Type"],
    "typologie":         ["Typologie"],
    "marque":            ["Marque"],
    "collection":        ["Collection"],
    "code_affaire":      ["Code Affaire", "Code affaire", "N° Affaire"],
}


def _norm(s):
    """Normalisation pour matching de colonnes."""
    return (s or "").strip().lower().replace("°", "").replace(" ", "")


def detect_columns(headers):
    """
    Auto-détecte le mapping en cherchant chaque champ DB dans les en-têtes CSV.
    Retourne un dict {field: column_index_or_None}.
    """
    norm_headers = [_norm(h) for h in headers]
    mapping = {}
    for field, aliases in DEFAULT_MAPPING.items():
        idx = None
        for alias in aliases:
            target = _norm(alias)
            for i, h in enumerate(norm_headers):
                if h == target:
                    idx = i
                    break
            if idx is not None:
                break
        mapping[field] = idx
    return mapping


def read_csv_preview(path, max_rows=15, delimiter=None, encoding=None):
    """
    Lit le CSV et retourne (headers, rows_preview, total_lines, detected_mapping, used_delim, used_enc).
    Détection auto du séparateur et de l'encodage si non fournis.
    """
    path = Path(path)
    raw = path.read_bytes()

    # Détection encodage
    if encoding is None:
        for enc in ("utf-8-sig", "utf-8", "cp1252", "iso-8859-1", "latin-1"):
            try:
                text = raw.decode(enc)
                encoding = enc
                break
            except UnicodeDecodeError:
                continue
        else:
            text = raw.decode("utf-8", errors="replace")
            encoding = "utf-8"
    else:
        text = raw.decode(encoding, errors="replace")

    # Détection séparateur
    if delimiter is None:
        sample = "\n".join(text.splitlines()[:5])
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            # Fallback : compter les occurrences
            counts = {sep: sample.count(sep) for sep in [";", ",", "\t", "|"]}
            delimiter = max(counts, key=counts.get) if counts else ","

    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)
    if not rows:
        return [], [], 0, {}, delimiter, encoding

    headers = rows[0]
    data = rows[1:]
    preview = data[:max_rows]
    mapping = detect_columns(headers)
    return headers, preview, len(data), mapping, delimiter, encoding


def _cell(row, idx):
    """Récupère une cellule en gérant les index None ou hors limites."""
    if idx is None or idx < 0 or idx >= len(row):
        return ""
    v = row[idx]
    return v.strip() if v else ""


def import_rows(path, mapping, delimiter, encoding,
                skip_empty_serie=False, dry_run=False, progress_cb=None):
    """
    Importe les lignes du CSV.

    Stratégie de dédoublonnage :
      - Par N° série (col `num_serie`) : un seul moteur créé par N° série unique.
      - Si N° série vide ET skip_empty_serie=True : ligne ignorée.
      - Si N° série vide ET skip_empty_serie=False : moteur quand même créé
        avec un N° généré "AUTO-NNNN".
      - Clients dédoublonnés par nom (insensible casse/espaces).

    Retourne un dict avec les stats : {
        'clients_crees', 'clients_existants',
        'moteurs_crees', 'moteurs_mis_a_jour', 'moteurs_doublons',
        'lignes_ignorees', 'erreurs'  (liste de tuples (n_ligne, msg))
    }
    """
    stats = {
        "clients_crees": 0, "clients_existants": 0,
        "moteurs_crees": 0, "moteurs_mis_a_jour": 0, "moteurs_doublons": 0,
        "lignes_ignorees": 0, "lignes_total": 0,
        "erreurs": [],
    }

    path = Path(path)
    text = path.read_bytes().decode(encoding, errors="replace")
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)
    if not rows:
        return stats
    data = rows[1:]
    stats["lignes_total"] = len(data)

    # Caches en mémoire pour cette import
    clients_cache = {}    # nom_normalisé → client_id
    series_vues = set()    # N° série déjà traités dans CE batch (dédoublonnage)
    auto_counter = 1

    for i, row in enumerate(data, start=2):  # ligne 2 = première ligne de données
        if progress_cb:
            progress_cb(i - 1, len(data))

        nom_client  = _cell(row, mapping.get("nom_client"))
        num_serie   = _cell(row, mapping.get("num_serie"))

        # Ligne complètement vide ?
        if not any(_cell(row, idx) for idx in mapping.values() if idx is not None):
            stats["lignes_ignorees"] += 1
            continue

        # N° série vide
        if not num_serie:
            if skip_empty_serie:
                stats["lignes_ignorees"] += 1
                continue
            # Générer un N° série automatique unique
            while True:
                gen = f"AUTO-{auto_counter:04d}"
                auto_counter += 1
                if gen not in series_vues and not db.find_moteur_by_serie(gen):
                    num_serie = gen
                    break

        # Dédoublonnage strict par N° série dans CE batch
        if num_serie in series_vues:
            stats["moteurs_doublons"] += 1
            continue
        series_vues.add(num_serie)

        # Récupérer ou créer le client
        client_id = ""
        if nom_client:
            nom_norm = nom_client.strip().lower()
            if nom_norm in clients_cache:
                client_id = clients_cache[nom_norm]
            else:
                existant = db.find_client_by_nom(nom_client)
                if existant:
                    client_id = existant["id"]
                    stats["clients_existants"] += 1
                else:
                    if not dry_run:
                        client_id = db.upsert_client({
                            "nom": nom_client.strip(),
                            "contact": "", "email": "",
                            "telephone": "", "adresse": "",
                        })
                    else:
                        client_id = f"DRY-{nom_norm}"
                    stats["clients_crees"] += 1
                clients_cache[nom_norm] = client_id

        # Préparer les données moteur
        moteur_data = {
            "client_id":        client_id,
            "num_serie":        num_serie,
            "navire":           _cell(row, mapping.get("navire")),
            "machine":          _cell(row, mapping.get("ref_constructeur")) or _cell(row, mapping.get("type_moteur")),
            "type_moteur":      _cell(row, mapping.get("type_moteur")),
            "date_mise_service":"",
            "duree_garantie":   "",
            "cylindree":        _cell(row, mapping.get("cylindree")),
            "famille":          _cell(row, mapping.get("famille")),
            "marque":           _cell(row, mapping.get("marque")),
            "application":      _cell(row, mapping.get("application")),
            "typologie":        _cell(row, mapping.get("typologie")),
            "collection":       _cell(row, mapping.get("collection")),
            "ref_constructeur": _cell(row, mapping.get("ref_constructeur")),
            "code_affaire":     _cell(row, mapping.get("code_affaire")),
            "type_client":      _cell(row, mapping.get("type_client")),
        }

        # Existe déjà en base ? → mise à jour, sinon création
        try:
            existant = db.find_moteur_by_serie(num_serie)
            if existant:
                if not dry_run:
                    db.upsert_moteur(moteur_data, moteur_id=existant["id"])
                stats["moteurs_mis_a_jour"] += 1
            else:
                if not dry_run:
                    db.upsert_moteur(moteur_data)
                stats["moteurs_crees"] += 1
        except Exception as e:
            stats["erreurs"].append((i, str(e)))

    return stats
