# EMS – Outil de suivi des interventions

**Emeraude Moteurs Systèmes** – Gestion des demandes clients & bons d'intervention
Version 1.8 – Indice A | Auteur : Vincent BIGOT | Référence cahier des charges : 30/04/2026

---

## Nouveautés v1.8

### Tableau de bord modulable
Configurable via le bouton **⚙️ Configurer** :
- **8 widgets** au choix : cartes statistiques, interventions urgentes, activité récente, garanties expirantes, charge par technicien, répartition par type, bons non notifiés, classifications
- **Ordre personnalisable** : flèches ▲▼ pour réorganiser
- **13 cartes statistiques** au choix : En cours / Clos / Facturé / Total / Urgentes / Critiques / Garantie / Facturables / Internes / Non notifiés / Clients / Moteurs / Techniciens
- **Préférences persistantes** : sauvegardées dans la base de données

### Critère d'urgence
Trois niveaux : **Normale** / **Urgente** / **Critique**
- Affichage avec badge coloré sur le bon HTML
- Surlignage rouge/orange des lignes dans tous les tableaux
- Tri automatique : les urgentes/critiques remontent en haut
- Filtre dédié dans l'onglet Interventions

### Classifications (3 cases indépendantes)
- **☐ Garantie** : auto-cochée si moteur sous garantie active
- **☐ Facturable**
- **☐ Interne**
Apparaissent sous forme de pills colorées sur le bon HTML.

### Heure de Paris partout
- Tous les timestamps affichés sont en heure Europe/Paris
- Gestion automatique de l'heure d'été/d'hiver (DST)
- Format court `JJ/MM HH:MM` dans les tableaux

### Notifications mieux mises en évidence
- Colonne dédiée **Notif.** : `📧 ✓ / 🔧 ✓` (notifié) ou `📧 ○ / 🔧 ○` (en attente)
- Distingue clairement client et technicien
- Widget dédié "Bons non notifiés" sur le tableau de bord

### Email plus permissif
- La validation des emails est désormais **informative et non bloquante**
- Format douteux → simple confirmation, pas de refus
- Coloration orange (warning) au lieu de rouge (erreur)

### Onglet Techniciens amélioré
- Colonne "Interventions" : compteur du nombre de bons par technicien
- Alignement des colonnes corrigé (cadrage à gauche cohérent)

---

## Conformité au cahier des charges

| # | Exigence | Implémentation |
|---|---|---|
| 1 | Préremplir des bons d'intervention | ✅ Formulaire complet conforme au modèle papier EMS |
| 2 | Suivi des interventions | ✅ Tableau de bord modulable + historique |
| 3 | Prévenir le client de sa prise en charge | ✅ Bouton 📧 *Prévenir client* (mailto) |
| 4 | Prévenir le technicien et éditer les bons | ✅ Bouton 📧 *Prévenir technicien* + bon HTML |
| 5 | Tableau de bord modifiable | ✅ Modulable v1.8 |
| 6 | Bases : clients/navires/machines/N° série/dates/garanties/types | ✅ |
| 7 | Ajouter facilement | ✅ Boutons ➕ partout, ajout inline |
| 8 | Recherche directe par N° série (style garage) | ✅ Barre dédiée sidebar |
| 9 | Interface simple et ergonomique | ✅ |
| 10 | Dossier automatique pour photos/documents | ✅ Bouton 📎 Ajouter fichier(s) |
| 11 | Types d'intervention configurables | ✅ |

Le bon HTML reproduit fidèlement le modèle papier EMS (logo, slogan, en-tête, cases à cocher Entretien/Dépannage/Diagnostic/Garantie, tableau matériels, tableau déplacements, signatures).

---

## Structure du projet

```
EMS_Interventions/
├── main.py              ← Interface Tkinter (tableau de bord modulable)
├── database.py          ← SQLite + heure Paris + migration douce
├── bon_generator.py     ← Génération HTML conforme modèle EMS
├── mailer.py            ← Notifications client/technicien
├── assets/
│   └── logo_ems.png     ← Logo officiel EMS
├── data/
│   └── ems.db           ← Base SQLite
├── dossiers/
│   └── BON-2026-0001/   ← Un dossier par bon
├── README.md
└── requirements.txt
```

---

## Installation

- **Python 3.8 ou supérieur**
- **tkinter** (inclus avec Python sur Windows et macOS)

### Linux (si tkinter absent) :
```bash
sudo apt install python3-tk        # Debian/Ubuntu
sudo dnf install python3-tkinter   # Fedora
```

Aucune dépendance externe — uniquement la bibliothèque standard Python.

---

## Lancement

```bash
python main.py
```

Windows : double-clic sur `Lancer_EMS.bat`
macOS/Linux : `./lancer_ems.sh`

---

## Migration depuis v1.5 / v1.6 / v1.7

Automatique au premier lancement. Toutes les nouvelles colonnes/tables sont ajoutées sans perte de données. Le compteur de bons et les types d'intervention sont préservés.

---

## Informations société

**Emeraude Moteurs Systèmes**
9bis avenue Louis Martin – 35400 Saint Malo
9 Rue d'Armorique – 35540 Miniac Morvan
Tél : 02.99.19.01.99 | Fax : 02.99.81.11.75
Courriel : service.technique@emeraudemoteurs.com
Siret 431 976 729 00027 | TVA intra FR 14 431 976 729
www.emeraudemoteurs.com
