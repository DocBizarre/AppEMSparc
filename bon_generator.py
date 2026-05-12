"""
EMS – Génération des bons d'intervention HTML (v2)
Layout fidèle au modèle papier officiel EMS (cahier des charges).
"""

import base64
import json
import os
import platform
import subprocess
from datetime import datetime
from pathlib import Path

DOSSIERS_PATH = Path(__file__).parent / "dossiers"
LOGO_PATH     = Path(__file__).parent / "assets" / "logo_ems.png"

# Logo embarqué (fallback si assets/logo_ems.png absent)
try:
    from logo_data import LOGO_EMS_B64
except ImportError:
    LOGO_EMS_B64 = ""

# Les 4 types officiels qui apparaissent comme cases à cocher dans l'en-tête
TYPES_HEADER = ["Entretien", "Dépannage", "Diagnostic", "Garantie"]


def _g(obj, key, default=""):
    """Accès tolérant à un sqlite3.Row, dict ou None."""
    if obj is None:
        return default
    try:
        v = obj[key]
        return v if v is not None else default
    except (KeyError, IndexError):
        return default


def _esc(s):
    """Échappement HTML minimaliste pour valeurs utilisateur."""
    if s is None:
        return ""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))


def _logo_data_uri():
    """
    Retourne le logo en data:URI.
    1. Essaie d'abord le fichier assets/logo_ems.png (si présent)
    2. Sinon fallback sur le logo embarqué dans logo_data.py
    3. En dernier recours : chaîne vide (le code utilise alors un texte fallback)
    """
    # 1. Fichier sur disque (priorité, permet de personnaliser le logo)
    if LOGO_PATH.is_file():
        try:
            b = LOGO_PATH.read_bytes()
            return "data:image/png;base64," + base64.b64encode(b).decode("ascii")
        except OSError:
            pass
    # 2. Logo embarqué (fonctionne même si assets/ absent)
    if LOGO_EMS_B64:
        return "data:image/png;base64," + LOGO_EMS_B64
    # 3. Aucun logo disponible
    return ""


def _check(cond):
    """Case à cocher : ☒ (cochée) ou ☐ (vide). Robuste aux 0/1, '0'/'1', True/False."""
    try:
        return "☒" if int(cond) else "☐"
    except (ValueError, TypeError):
        return "☒" if bool(cond) else "☐"


def generer_bon_html(inv, client=None, moteur=None):
    """
    Génère le HTML du bon d'intervention au format EMS officiel.
    inv    : sqlite3.Row ou dict de l'intervention (avec jointures clients/moteurs OK)
    client : optionnel si déjà jointuré
    moteur : optionnel si déjà jointuré
    """
    # ── Identification du bon ─────────────────────────────────────────────────
    num_bon  = _g(inv, "num_bon")
    date_i   = _g(inv, "date_creation")
    technicien = _g(inv, "technicien")
    statut   = _g(inv, "statut", "En cours")
    urgence  = _g(inv, "urgence", "Normale")

    # Type d'intervention : on coche la case si le libellé correspond à un des 4
    type_inv = _g(inv, "type_intervention")

    # Classifications
    cls_garantie = _g(inv, "garantie_intervention", 0)
    cls_factur   = _g(inv, "facturable", 0)
    cls_interne  = _g(inv, "interne", 0)

    # ── Client et signataire ──────────────────────────────────────────────────
    c_nom     = _g(client, "nom")       or _g(inv, "client_nom")
    c_contact = _g(client, "contact")   or _g(inv, "client_contact")
    c_email   = _g(client, "email")     or _g(inv, "client_email")
    c_tel     = _g(client, "telephone") or _g(inv, "client_tel")
    c_adresse = _g(client, "adresse")   or _g(inv, "client_adresse")

    lieu          = _g(inv, "lieu_intervention")
    nom_signataire = _g(inv, "nom_signataire") or c_contact
    email_signataire = _g(inv, "email_signataire") or c_email

    # ── Équipement ────────────────────────────────────────────────────────────
    navire     = _g(moteur, "navire")      or _g(inv, "navire")
    machine    = _g(moteur, "machine")     or _g(inv, "machine")
    type_mot   = _g(moteur, "type_moteur") or _g(inv, "type_moteur")
    num_serie  = _g(moteur, "num_serie")   or _g(inv, "num_serie")
    date_svc   = _g(moteur, "date_mise_service") or _g(inv, "date_mise_service")
    garantie   = _g(moteur, "duree_garantie")    or _g(inv, "duree_garantie")
    nb_heures  = _g(inv, "nb_heures_fct")
    marque     = _g(moteur, "marque")            or _g(inv, "marque")
    ref_const  = _g(moteur, "ref_constructeur")   or _g(inv, "ref_constructeur")
    # Type complet affiché : "MARQUE Référence" en priorité
    # (ex: "BAUDOUIN 12M26.3"), sinon fallback sur type_moteur / machine
    tech_ref = ref_const or type_mot or machine
    type_complet = f"{marque} {tech_ref}".strip() if marque else tech_ref

    # ── Options ───────────────────────────────────────────────────────────────
    opt_diag   = _g(inv, "outil_diagnostic", 0)
    mem_avant  = _g(inv, "memoriser_avant", 0)
    mem_apres  = _g(inv, "memoriser_apres", 0)
    ph_avant   = _g(inv, "photos_avant", 0)
    ph_apres   = _g(inv, "photos_apres", 0)
    pour_info  = _g(inv, "pour_information", 0)
    preco      = _g(inv, "preconisation", 0)

    # ── Zones de texte ────────────────────────────────────────────────────────
    demande_client = _g(inv, "demande_client") or _g(inv, "description")
    constat        = _g(inv, "constat")
    travaux        = _g(inv, "travaux")
    informations   = _g(inv, "informations")

    # ── Tableaux JSON ─────────────────────────────────────────────────────────
    try:
        materiels = json.loads(_g(inv, "materiels_json", "[]") or "[]")
    except (json.JSONDecodeError, TypeError):
        materiels = []
    try:
        depl = json.loads(_g(inv, "deplacements_json", "{}") or "{}")
    except (json.JSONDecodeError, TypeError):
        depl = {}

    # Compatibilité : ancien champ "pieces" texte libre → injecté en première ligne
    legacy_pieces = _g(inv, "pieces")
    if legacy_pieces and not materiels:
        materiels = [{"qte": "", "ref": "", "designation": legacy_pieces}]

    # ── Construction des lignes du tableau Matériels (au moins 5 lignes) ──────
    n_min_mat = 5
    mat_rows = []
    for m in materiels:
        mat_rows.append((
            _esc(m.get("qte", "")),
            _esc(m.get("ref", "")),
            _esc(m.get("designation", ""))
        ))
    while len(mat_rows) < n_min_mat:
        mat_rows.append(("", "", ""))

    mat_html = "\n".join(
        f"      <tr><td>{q}</td><td>{r}</td><td>{d}</td></tr>"
        for q, r, d in mat_rows
    )

    # ── Construction du tableau Déplacements ──────────────────────────────────
    def d(k):
        return _esc(depl.get(k, ""))

    # ── Statut visuel (cercle coloré) ─────────────────────────────────────────
    statut_cls = {"En cours": "ec", "Clos": "clos", "Facturé": "fact"}.get(statut, "ec")
    urg_cls = {"Critique": "crit", "Urgente": "urg", "Normale": "norm"}.get(urgence, "norm")

    # ── Cases à cocher pour les 4 types officiels (en-tête) ───────────────────
    type_cases = ""
    for t in TYPES_HEADER:
        coche = "☒" if t.lower() == type_inv.lower() else "☐"
        type_cases += f'<div class="type-row"><span class="cb">{coche}</span> {t}</div>\n'

    # Si type_inv est différent des 4, l'afficher en bas du bloc
    type_other = ""
    if type_inv and type_inv not in TYPES_HEADER:
        type_other = f'<div class="type-row"><span class="cb">☒</span> {_esc(type_inv)}</div>'

    logo_uri = _logo_data_uri()
    logo_html = f'<img src="{logo_uri}" alt="EMS">' if logo_uri else '<div class="logo-fallback">EMS</div>'

    # ── HTML final ────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Bon d'intervention {num_bon}</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: Arial, Helvetica, sans-serif; font-size: 10.5px; color: #000;
       background: #fff; padding: 16px 20px; }}
table {{ border-collapse: collapse; width: 100%; }}
td, th {{ border: 1px solid #000; padding: 3px 5px; vertical-align: top; }}

/* En-tête en 3 colonnes */
.header {{ display: grid; grid-template-columns: 170px 1fr 200px; gap: 0;
           border-bottom: 2px solid #000; padding-bottom: 6px; margin-bottom: 8px; }}
.header-left {{ display: flex; flex-direction: column; align-items: center;
                justify-content: center; padding: 4px; }}
.header-left img {{ max-width: 150px; max-height: 80px; }}
.header-left .logo-fallback {{ font-size: 28px; font-weight: bold; color: #003366;
                                letter-spacing: 3px; }}
.header-left .slogan {{ font-size: 8px; color: #555; font-style: italic;
                        margin-top: 2px; text-align: center; }}
.header-mid {{ text-align: center; padding: 4px; }}
.header-mid h1 {{ font-size: 16px; font-weight: bold; margin-bottom: 4px; }}
.header-mid .info {{ font-size: 9.5px; line-height: 1.4; }}
.header-right {{ padding: 4px 8px; }}
.header-right .titre {{ font-weight: bold; font-size: 11px; margin-bottom: 4px; }}
.type-row {{ font-size: 10.5px; padding: 1px 0; }}
.type-row .cb {{ font-family: "Segoe UI Symbol", Arial, sans-serif;
                 font-size: 13px; margin-right: 4px; }}

/* Référence du bon (encadré en haut à gauche) */
.ref-box {{ position: absolute; top: 10px; left: 18px;
            background: #003366; color: #fff; padding: 4px 10px;
            font-size: 11px; font-weight: bold; border-radius: 4px;
            z-index: 10; }}
.statut-badge {{ display: inline-block; padding: 1px 8px; margin-left: 6px;
                 border-radius: 8px; font-size: 9px; font-weight: bold; }}
.statut-ec   {{ background: #fff3cd; color: #856404; }}
.statut-clos {{ background: #d1e7dd; color: #0f5132; }}
.statut-fact {{ background: #cfe2ff; color: #084298; }}
.urg-badge   {{ display: inline-block; padding: 1px 8px; margin-left: 4px;
                 border-radius: 8px; font-size: 9px; font-weight: bold; }}
.urg-norm    {{ background: #e9ecef; color: #495057; }}
.urg-urg     {{ background: #ffe5cc; color: #cc4400; }}
.urg-crit    {{ background: #f8d7da; color: #721c24; }}
.classif {{ font-size: 9.5px; margin-top: 4px; }}
.classif .cb {{ font-family: "Segoe UI Symbol", Arial, sans-serif; font-size: 12px; }}
.classif .pill {{ display: inline-block; padding: 1px 6px; margin-right: 4px;
                   background: #f0f0f0; border-radius: 8px; font-weight: bold; }}
.classif .pill.on {{ background: #003366; color: #fff; }}

/* Bloc info client/équipement (2 colonnes) */
.bloc-info {{ width: 100%; margin-top: 4px; }}
.bloc-info td {{ font-size: 10px; height: 18px; }}
.bloc-info .lbl {{ background: #f5f5f5; font-weight: bold; width: 28%; }}

/* Sections texte */
.section-title {{ font-weight: bold; font-size: 11px; text-decoration: underline;
                  margin: 10px 0 3px; }}
.zone-texte {{ border: 1px solid #000; min-height: 50px; padding: 4px 6px;
               white-space: pre-wrap; font-size: 10px; line-height: 1.5; }}
.zone-grande {{ min-height: 70px; }}

/* Ligne d'options (cases à cocher inline) */
.options {{ margin: 4px 0; font-size: 10.5px; }}
.options .opt {{ display: inline-block; margin-right: 16px; }}
.options .cb {{ font-family: "Segoe UI Symbol", Arial, sans-serif;
                font-size: 13px; margin-right: 3px; }}
.options strong {{ font-size: 10.5px; }}

/* Tableau matériels */
.materiels {{ margin-top: 4px; }}
.materiels th {{ background: #003366; color: #fff; font-size: 10px;
                 text-align: left; padding: 4px 6px; }}
.materiels td {{ height: 22px; font-size: 10px; }}
.materiels .qte {{ width: 60px; text-align: center; }}
.materiels .ref {{ width: 130px; }}

/* Bloc info / préco à droite des matériels */
.split {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px;
          margin-top: 6px; }}
.info-box {{ border: 1px solid #000; padding: 4px 6px; min-height: 90px;
             font-size: 10px; line-height: 1.4; }}
.info-box .head {{ font-size: 10.5px; margin-bottom: 4px; }}
.info-box .head .opt {{ display: inline-block; margin-right: 18px; }}
.info-box .body {{ white-space: pre-wrap; }}

/* Tableau déplacements */
.depl {{ width: 100%; margin-top: 4px; }}
.depl th {{ background: #003366; color: #fff; font-size: 10px;
            padding: 4px 6px; text-align: left; }}
.depl td {{ height: 18px; font-size: 10px; }}
.depl .lbl {{ background: #f5f5f5; font-weight: bold; width: 25%; }}
.depl .val {{ width: 25%; }}

/* Signatures */
.signatures {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px;
               margin-top: 16px; }}
.sign-box {{ border: 1px solid #000; padding: 6px; min-height: 70px;
             font-size: 10px; }}
.sign-box .head {{ font-weight: bold; margin-bottom: 4px; }}
.sign-box .lab  {{ font-size: 9px; color: #555; margin-top: 30px;
                   border-top: 1px dotted #999; padding-top: 3px; }}

/* Pied de page */
.footer {{ border-top: 1px solid #000; margin-top: 14px; padding-top: 6px;
           font-size: 8.5px; color: #444; text-align: center; line-height: 1.6; }}
.footer strong {{ color: #003366; }}

/* Bouton imprimer */
.print-btn {{ position: fixed; top: 10px; right: 10px; background: #003366;
              color: #fff; border: 0; padding: 8px 16px; border-radius: 5px;
              font-size: 11px; cursor: pointer; font-family: Arial; }}
@media print {{
  .print-btn {{ display: none; }}
  body {{ padding: 8mm; }}
  .ref-box {{ position: static; float: left; margin-bottom: 4px; }}
}}
</style>
</head>
<body>
<button class="print-btn" onclick="window.print()">🖨 Imprimer / PDF</button>

<div class="ref-box">{_esc(num_bon)}
  <span class="statut-badge statut-{statut_cls}">{_esc(statut)}</span>
  <span class="urg-badge urg-{urg_cls}">⚡ {_esc(urgence)}</span>
</div>

<!-- ═══ EN-TÊTE ═══ -->
<div class="header">
  <div class="header-left">
    {logo_html}
    <div class="slogan">L'application motorisée,<br>au cœur de vos énergies</div>
  </div>
  <div class="header-mid">
    <h1>BON D'INTERVENTION</h1>
    <div class="info">
      9bis avenue Louis Martin – 35400 Saint Malo<br>
      Tél : 02.99.19.01.99 &nbsp;–&nbsp; Fax : 02.99.81.11.75<br>
      Courriel : service.technique@emeraudemoteurs.com<br>
      Siret 431 976 729 00027 &nbsp;|&nbsp; TVA intra FR 14 431 976 729
    </div>
  </div>
  <div class="header-right">
    <div class="titre">Type d'intervention :</div>
    {type_cases}
    {type_other}
  </div>
</div>

<!-- ═══ BLOC CLIENT / ÉQUIPEMENT (2 COLONNES) ═══ -->
<table class="bloc-info">
  <tr>
    <td class="lbl">Informations société<br>Signature</td>
    <td>{_esc(c_nom)}<br><small>{_esc(c_adresse)}</small></td>
    <td class="lbl">Nom bateau / type machine</td>
    <td>{_esc(navire)}</td>
  </tr>
  <tr>
    <td class="lbl">Lieu de l'intervention</td>
    <td>{_esc(lieu)}</td>
    <td class="lbl">Type moteur / inverseur</td>
    <td>{_esc(type_complet)}</td>
  </tr>
  <tr>
    <td class="lbl">Nom du signataire</td>
    <td>{_esc(nom_signataire)}</td>
    <td class="lbl">N° de série</td>
    <td><strong>{_esc(num_serie)}</strong></td>
  </tr>
  <tr>
    <td class="lbl">Courriel du signataire</td>
    <td>{_esc(email_signataire)}</td>
    <td class="lbl">Nb heures de fonctionnement</td>
    <td>{_esc(nb_heures)}</td>
  </tr>
  <tr>
    <td class="lbl">Téléphone</td>
    <td>{_esc(c_tel)}</td>
    <td class="lbl">Date de mise en service</td>
    <td>{_esc(date_svc)}</td>
  </tr>
  <tr>
    <td class="lbl">Date de l'intervention</td>
    <td>{_esc(date_i)}</td>
    <td class="lbl">Garantie</td>
    <td>{_esc(garantie)} mois</td>
  </tr>
  <tr>
    <td class="lbl">Technicien EMS</td>
    <td><strong>{_esc(technicien)}</strong></td>
    <td class="lbl">Date de l'intervention</td>
    <td>{_esc(date_i)}</td>
  </tr>
</table>

<!-- Classifications -->
<div class="classif">
  <span class="pill {'on' if cls_garantie else ''}">{_check(cls_garantie)} Garantie</span>
  <span class="pill {'on' if cls_factur else ''}">{_check(cls_factur)} Facturable</span>
  <span class="pill {'on' if cls_interne else ''}">{_check(cls_interne)} Interne</span>
</div>

<!-- ═══ DEMANDE DU CLIENT ═══ -->
<div class="section-title">DEMANDE DU CLIENT :</div>
<div class="zone-texte">{_esc(demande_client)}</div>

<!-- ═══ OPTIONS / DIAGNOSTIC ═══ -->
<div class="options">
  <span class="opt"><span class="cb">{_check(opt_diag)}</span> Utilisation de l'Outil de diagnostic</span>
  <span class="opt"><strong>Mémoriser les données :</strong>
    <span class="cb">{_check(mem_avant)}</span> avant
    <span class="cb">{_check(mem_apres)}</span> après</span>
  <span class="opt"><strong>PHOTOS :</strong>
    <span class="cb">{_check(ph_avant)}</span> avant
    <span class="cb">{_check(ph_apres)}</span> après</span>
</div>

<!-- ═══ CONSTAT AVANT INTERVENTION ═══ -->
<div class="section-title">CONSTAT AVANT INTERVENTION :</div>
<div class="zone-texte zone-grande">{_esc(constat)}</div>

<!-- ═══ TRAVAUX ═══ -->
<div class="section-title">TRAVAUX :</div>
<div class="zone-texte zone-grande">{_esc(travaux)}</div>

<!-- ═══ MATÉRIELS UTILISÉS + INFORMATIONS ═══ -->
<div class="section-title">MATÉRIELS UTILISÉS</div>
<div class="split">
  <table class="materiels">
    <thead>
      <tr>
        <th class="qte">QUANTITÉ</th>
        <th class="ref">RÉFÉRENCE</th>
        <th>DÉSIGNATION</th>
      </tr>
    </thead>
    <tbody>
{mat_html}
    </tbody>
  </table>
  <div class="info-box">
    <div class="head">
      <span class="opt"><span class="cb" style="font-family:'Segoe UI Symbol',Arial;font-size:13px;">{_check(pour_info)}</span> Pour information</span>
      <span class="opt"><span class="cb" style="font-family:'Segoe UI Symbol',Arial;font-size:13px;">{_check(preco)}</span> Préconisation</span>
    </div>
    <div class="body">{_esc(informations)}</div>
  </div>
</div>

<!-- ═══ DÉPLACEMENTS ═══ -->
<div class="section-title">DÉPLACEMENTS</div>
<table class="depl">
  <tr>
    <td class="lbl">Temps de trajet aller</td>      <td class="val">{d("trajet_aller")}</td>
    <td class="lbl">Frais de repas</td>             <td class="val">{d("frais_repas")}</td>
  </tr>
  <tr>
    <td class="lbl">Heure début intervention matin</td>      <td class="val">{d("heure_debut_matin")}</td>
    <td class="lbl">Frais d'hôtel</td>                       <td class="val">{d("frais_hotel")}</td>
  </tr>
  <tr>
    <td class="lbl">Heure fin intervention matin</td>        <td class="val">{d("heure_fin_matin")}</td>
    <td class="lbl">Frais de péages</td>                     <td class="val">{d("frais_peages")}</td>
  </tr>
  <tr>
    <td class="lbl">Heure début intervention après-midi</td> <td class="val">{d("heure_debut_apres")}</td>
    <td class="lbl">Temps de préparation avant intervention</td> <td class="val">{d("temps_preparation")}</td>
  </tr>
  <tr>
    <td class="lbl">Heure fin intervention après-midi</td>   <td class="val">{d("heure_fin_apres")}</td>
    <td class="lbl">Temps de rangement après intervention</td>   <td class="val">{d("temps_rangement")}</td>
  </tr>
  <tr>
    <td class="lbl">Temps de trajet retour</td>     <td class="val">{d("trajet_retour")}</td>
    <td class="lbl"></td>                            <td class="val"></td>
  </tr>
</table>

<!-- ═══ SIGNATURES ═══ -->
<div class="signatures">
  <div class="sign-box">
    <div class="head">Signature Client (Nom &amp; date) :</div>
    <div class="lab">Bon pour accord des travaux réalisés</div>
  </div>
  <div class="sign-box">
    <div class="head">Signature Technicien EMS :</div>
    <div class="lab">{_esc(technicien)}</div>
  </div>
</div>

<!-- ═══ PIED DE PAGE ═══ -->
<div class="footer">
  <strong>Emeraude Moteurs Systèmes</strong> – Constructeur de groupe de puissance<br>
  9bis avenue Louis Martin – 35400 Saint Malo &nbsp;|&nbsp;
  9 Rue d'Armorique – 35540 Miniac Morvan<br>
  Tél : 02.99.19.01.99 &nbsp;|&nbsp; Fax : 02.99.81.11.75 &nbsp;|&nbsp;
  <strong>www.emeraudemoteurs.com</strong><br>
  <em style="color:#888;">Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</em>
</div>
</body>
</html>"""
    return html


def sauvegarder_bon(inv):
    """Sauvegarde le HTML dans le dossier de l'intervention, retourne le path."""
    num_bon = _g(inv, "num_bon")
    if not num_bon:
        raise ValueError("Le bon n'a pas de num_bon")
    dossier = DOSSIERS_PATH / num_bon
    dossier.mkdir(parents=True, exist_ok=True)
    html = generer_bon_html(inv)
    path = dossier / f"{num_bon}.html"
    path.write_text(html, encoding="utf-8")
    return path


def ouvrir_fichier(path):
    p = str(path)
    if platform.system() == "Windows":
        os.startfile(p)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])
