#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMS – Application Bons d'intervention.

Création et gestion des bons d'intervention (tableau de bord,
interventions, nouveau bon). La création rapide d'un client et/ou
d'un moteur reste disponible directement depuis le formulaire de bon.

Le parc (saisie des clients / moteurs / techniciens) est géré par
l'application séparée « Gestion de parc » (app_parc.py).

Lancement :  python app_bons.py
"""

from main import AppEMS

if __name__ == "__main__":
    AppEMS(mode="bons").mainloop()
