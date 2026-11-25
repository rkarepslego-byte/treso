"""Reprise de l'ancien classeur de trésorerie du client (un onglet par mois).

Structure reconnue (onglets « NOV26 », « DEC26 »…) : ligne d'en-tête contenant
« Date | Tiers | Libellé | Type | Périodicité | BANQUE | Recettes Prévues | Recettes Réalisées |
Dépenses Prévues | Dep récurrente | Debité | Dépenses Réalisés ». La colonne Date contient le
jour du mois (nombre, ou date « 1900 » saisie comme jour).

Produit :
  * les charges fixes (lignes avec une périodicité et une dépense récurrente) ;
  * les prévisions ponctuelles du mois (dépenses / recettes prévues sans périodicité) ;
  * la liste des lignes non reprises (sans jour, sans banque…), à vérifier.

Usage : python -m outils.migrer_treso ancien.xlsx DEC26 --annee 2026 --mois 12
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass, field
from datetime import date, datetime

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RACINE not in sys.path:
    sys.path.insert(0, RACINE)

from reference import moteur as m  # noqa: E402

# Codes de l'ancien classeur (onglet DONNEES) -> catégories de l'outil
CATEGORIES_CLIENT = [
    ("CLT", "Clients", "Recette"), ("CB", "Carte bancaire", "Dépense"), ("FOUR", "Fournisseurs", "Dépense"),
    ("FG", "Frais généraux", "Dépense"), ("PRÊT/LOA", "Prêts / Leasing", "Dépense"),
    ("SAL", "Salaires et charges sociales", "Dépense"), ("INT", "Intérêts / Placements", "Les deux"),
    ("BQ", "Frais bancaires", "Dépense"), ("IMP", "Impôts et taxes", "Dépense"), ("DIV", "Divers", "Les deux"),
]
BANQUES_CLIENT = {"BPGO": "BPGO", "CIC": "CIC"}
FREQUENCES = {"MENSUEL": "Mensuelle", "MENSUELLE": "Mensuelle", "TRIMESTRE": "Trimestrielle",
              "TRIMESTRIEL": "Trimestrielle", "TRIMESTRIELLE": "Trimestrielle", "TRIM": "Trimestrielle",
              "SEMESTRE": "Semestrielle", "SEMESTRIEL": "Semestrielle", "SEMESTRIELLE": "Semestrielle",
              "ANNUEL": "Annuelle", "ANNUELLE": "Annuelle"}


def categorie(code) -> str:
    c = (str(code).strip().upper() if code else "")
    for k, lib, _t in CATEGORIES_CLIENT:
        if c == k:
            return lib
    return "Divers"


@dataclass
class Reprise:
    charges: list = field(default_factory=list)
    previsions: list = field(default_factory=list)
    ignorees: list = field(default_factory=list)


def _jour(v):
    if isinstance(v, datetime):
        return v.day if v.year == 1900 else None
    if isinstance(v, (int, float)) and 1 <= v <= 31:
        return int(v)
    return None


def _montant(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def lire_mois(chemin: str, onglet: str, annee: int, mois: int) -> Reprise:
    import calendar
    import openpyxl
    wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
    ws = wb[onglet]
    lignes = [list(r) for r in ws.iter_rows(max_col=14, values_only=True)]
    wb.close()
    i0 = next(i for i, l in enumerate(lignes)
              if l and len(l) > 6 and str(l[1] or "").strip().lower().startswith("date") and str(l[6] or "").strip().upper() == "BANQUE")
    rep = Reprise()
    dernier_jour = calendar.monthrange(annee, mois)[1]
    for num, l in enumerate(lignes[i0 + 1:], start=i0 + 2):
        l = (l + [None] * 14)[:14]
        _, d, tiers, lib, type_, period, banque, rp, rr, dp, rec, debite, dr, _s = l
        tiers = str(tiers or "").strip()
        lib = str(lib or "").strip()
        if lib and _est_nombre(lib):
            lib = ""
        libelle = " - ".join(x for x in (tiers, lib) if x)
        recette, depense_rec, depense = _montant(rp), _montant(rec), _montant(dp)
        if not libelle or (recette == 0 and depense_rec == 0 and depense == 0):
            continue
        jour = _jour(d)
        nom_banque = BANQUES_CLIENT.get(str(banque or "").strip().upper())
        if libelle.upper().startswith(("SAISIR", "FABIEN", "A PAYER")):
            continue
        if jour is None or nom_banque is None:
            rep.ignorees.append((num, libelle, recette or depense_rec or depense,
                                 "jour du mois manquant" if jour is None else "banque manquante"))
            continue
        dt = date(annee, mois, min(jour, dernier_jour))
        freq = FREQUENCES.get(str(period or "").strip().upper().replace("?", "").strip())
        if freq and depense_rec:
            rep.charges.append(m.Charge("", libelle, nom_banque, categorie(type_), m.TYPE_DEPENSE, depense_rec, dt, freq,
                                        commentaire=f"Repris de l'onglet {onglet} (ligne {num})"))
        elif recette:
            rep.previsions.append(m.Prevision("", dt, nom_banque, libelle, categorie(type_) if type_ else "Clients",
                                              m.TYPE_RECETTE, recette, f"Repris de {onglet}", origine="Import"))
        else:
            rep.previsions.append(m.Prevision("", dt, nom_banque, libelle, categorie(type_), m.TYPE_DEPENSE,
                                              depense or depense_rec, f"Repris de {onglet}", origine="Import"))
    return rep


def _est_nombre(s):
    try:
        float(str(s).replace(",", ".").replace(" ", ""))
        return True
    except ValueError:
        return False


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("classeur")
    ap.add_argument("onglet")
    ap.add_argument("--annee", type=int, required=True)
    ap.add_argument("--mois", type=int, required=True)
    a = ap.parse_args(argv)
    rep = lire_mois(a.classeur, a.onglet, a.annee, a.mois)
    print(f"{len(rep.charges)} charge(s) fixe(s) :")
    for c in rep.charges:
        print(f"  {c.premiere:%d/%m} {c.banque:<5} {c.frequence:<13} {c.montant:>10.2f}  {c.categorie:<28} {c.libelle}")
    print(f"{len(rep.previsions)} prévision(s) ponctuelle(s) :")
    for p in rep.previsions:
        print(f"  {p.date:%d/%m} {p.banque:<5} {p.type:<8} {p.montant:>10.2f}  {p.libelle}")
    print(f"{len(rep.ignorees)} ligne(s) non reprise(s) :")
    for num, lib, mt, raison in rep.ignorees:
        print(f"  ligne {num} : {lib} ({mt:.2f}) - {raison}")


if __name__ == "__main__":
    main()
