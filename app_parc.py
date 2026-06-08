#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMS – Application Gestion de parc.

Saisie et gestion du référentiel : clients, moteurs, techniciens.
Ces données sont partagées avec l'application « Bons d'intervention »
(même base de données).

Lancement :  python app_parc.py
"""

from main import AppEMS

if __name__ == "__main__":
    AppEMS(mode="parc").mainloop()
