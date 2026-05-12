"""
EMS – Notifications par email
Génère des brouillons d'email via mailto: (ouverture du client mail par défaut).
Aucune configuration SMTP nécessaire — fonctionne sur Windows / macOS / Linux.

Le HTML du bon est sauvegardé dans le dossier de l'intervention pour permettre
au technicien de le joindre manuellement à l'email (mailto: ne supporte pas
les pièces jointes de manière fiable cross-platform).
"""

from urllib.parse import quote
import webbrowser
from datetime import datetime


def _esc(s):
    """Encode pour query string mailto."""
    return quote(str(s or ""), safe="")


def _build_mailto(to="", cc="", subject="", body=""):
    parts = []
    if cc:
        parts.append(f"cc={_esc(cc)}")
    if subject:
        parts.append(f"subject={_esc(subject)}")
    if body:
        parts.append(f"body={_esc(body)}")
    qs = "&".join(parts)
    return f"mailto:{_esc(to)}" + (f"?{qs}" if qs else "")


# ─── Templates ────────────────────────────────────────────────────────────────
TEMPLATE_CLIENT = """Bonjour{contact_line},

Nous accusons réception de votre demande d'intervention.

Référence du bon : {num_bon}
Équipement : {machine}{navire_line}
N° de série : {num_serie}
Type d'intervention : {type_intervention}
Date prévue : {date_creation}

Votre dossier est désormais pris en charge par notre service technique. {technicien_line}vous contactera prochainement pour convenir des modalités d'intervention.

Pour toute question, n'hésitez pas à nous contacter au 02.99.19.01.99.

Cordialement,

L'équipe EMS – Emeraude Moteurs Systèmes
9bis avenue Louis Martin – 35400 Saint Malo
Tél : 02.99.19.01.99
www.emeraudemoteurs.com
"""

TEMPLATE_TECHNICIEN = """Bonjour {technicien},

Une nouvelle intervention vous est assignée.

═══════════════════════════════════════
Référence : {num_bon}
Statut    : {statut}
Date      : {date_creation}
═══════════════════════════════════════

CLIENT
  Nom      : {client_nom}
  Contact  : {contact}
  Tél      : {tel}
  Email    : {email}
  Adresse  : {adresse}

ÉQUIPEMENT
  Navire/Site : {navire}
  Machine     : {machine}
  N° de série : {num_serie}
  Mise en service : {date_mise_service}
  Garantie    : {duree_garantie} mois

INTERVENTION
  Type        : {type_intervention}
  Description : {description}

Bon HTML à joindre :
{bon_path}

Bon courage,
EMS
"""


def _safe(d, key, default=""):
    """Accès tolérant pour dict ou sqlite3.Row."""
    if d is None:
        return default
    try:
        v = d[key]
        return v if v is not None else default
    except (KeyError, IndexError):
        return default


def email_client(inv, client, moteur, bon_path=""):
    """
    Ouvre le client mail avec un brouillon de notification client.
    inv, client, moteur : sqlite3.Row ou dict.
    """
    contact  = _safe(client, "contact").strip()
    nom      = _safe(client, "nom")
    email    = _safe(client, "email")
    machine  = _safe(moteur, "machine") or _safe(inv, "machine")
    navire   = _safe(moteur, "navire")  or _safe(inv, "navire")
    num_bon  = _safe(inv, "num_bon")

    contact_line = f" {contact}" if contact else f" {nom}"
    navire_line  = f" – {navire}" if navire else ""
    technicien   = _safe(inv, "technicien")
    technicien_line = f"Notre technicien {technicien} " if technicien else "Un technicien "

    subject = f"[EMS] Prise en charge de votre intervention – {num_bon}"
    body = TEMPLATE_CLIENT.format(
        contact_line=contact_line,
        num_bon=num_bon,
        machine=machine or "—",
        navire_line=navire_line,
        num_serie=_safe(moteur, "num_serie") or _safe(inv, "num_serie"),
        type_intervention=_safe(inv, "type_intervention"),
        date_creation=_safe(inv, "date_creation"),
        technicien_line=technicien_line,
    )

    url = _build_mailto(to=email, subject=subject, body=body)
    webbrowser.open(url)
    return url


def email_technicien(inv, client, moteur, technicien_email="", bon_path=""):
    """
    Ouvre le client mail avec un brouillon d'assignation au technicien.
    technicien_email : adresse du technicien (peut être vide → laissée à compléter).
    """
    num_bon = _safe(inv, "num_bon")
    subject = f"[EMS] Nouveau bon assigné – {num_bon} – {_safe(inv, 'type_intervention')}"
    body = TEMPLATE_TECHNICIEN.format(
        technicien=_safe(inv, "technicien"),
        num_bon=num_bon,
        statut=_safe(inv, "statut", "En cours"),
        date_creation=_safe(inv, "date_creation"),
        client_nom=_safe(client, "nom") or _safe(inv, "client_nom"),
        contact=_safe(client, "contact"),
        tel=_safe(client, "telephone"),
        email=_safe(client, "email"),
        adresse=_safe(client, "adresse"),
        navire=_safe(moteur, "navire") or _safe(inv, "navire"),
        machine=_safe(moteur, "machine") or _safe(inv, "machine"),
        num_serie=_safe(moteur, "num_serie") or _safe(inv, "num_serie"),
        date_mise_service=_safe(moteur, "date_mise_service"),
        duree_garantie=_safe(moteur, "duree_garantie"),
        type_intervention=_safe(inv, "type_intervention"),
        description=_safe(inv, "description"),
        bon_path=bon_path or "(à générer)",
    )
    url = _build_mailto(to=technicien_email, subject=subject, body=body)
    webbrowser.open(url)
    return url
