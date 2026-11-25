"""Jeu de données fictives de démonstration et de test.

Toutes les dates sont calculées par rapport à une date de référence (par défaut la
date du jour de la génération) afin que la démonstration reste parlante.

Le jeu couvre volontairement les cas difficiles du cahier des charges :
  * paiements prévus puis débités (libellés bancaires différents des libellés prévus)
  * dépenses bancaires non prévues
  * deux paiements identiques légitimes le même jour (péages, Banque 2)
  * virement interne Banque 1 -> Banque 2
  * charges fixes mensuelles / trimestrielles / annuelles
  * deux prévisions identiques candidates pour une même opération (correspondance incertaine)
  * deux relevés Banque 1 qui se chevauchent
  * une opération antérieure à la date du solde initial (ignorée)
"""

from __future__ import annotations

import csv
import os
from datetime import date, timedelta

from reference import moteur as m

NOM_B1, NOM_B2 = "Banque Alpha", "Banque Beta"


def _jour(ref: date, d: int) -> date:
    return ref + timedelta(days=d)


def _jour_du_mois(base: date, jour: int) -> date:
    return date(base.year, base.month, min(jour, 28))


def parametres_demo(ref: date) -> m.Parametres:
    d0 = ref - timedelta(days=35)
    return m.Parametres(
        banques=[m.Banque("B1", NOM_B1, 48000.00, d0), m.Banque("B2", NOM_B2, 8000.00, d0)],
        date_ref=ref,
    )


def etat_initial(ref: date) -> m.Etat:
    """Prévisions et charges fixes telles qu'un utilisateur les aurait saisies."""
    params = parametres_demo(ref)
    d0 = params.banques[0].date_solde
    etat = m.Etat(params)
    debut = d0 - timedelta(days=60)

    etat.charges = [
        m.Charge("CF-001", "Loyer bureaux", NOM_B1, "Loyer", m.TYPE_DEPENSE, 2400.00, _jour_du_mois(debut, 5), "Mensuelle"),
        m.Charge("CF-002", "Salaires", NOM_B1, "Salaires", m.TYPE_DEPENSE, 18500.00, _jour_du_mois(debut, 27), "Mensuelle"),
        m.Charge("CF-003", "URSSAF", NOM_B1, "Charges sociales", m.TYPE_DEPENSE, 6200.00, _jour_du_mois(debut, 15), "Mensuelle"),
        m.Charge("CF-004", "Abonnement Orange Pro", NOM_B2, "Télécom / Internet", m.TYPE_DEPENSE, 89.90,
                 _jour_du_mois(debut, 10), "Mensuelle"),
        m.Charge("CF-005", "Assurance multirisque", NOM_B2, "Assurances", m.TYPE_DEPENSE, 1350.00,
                 _jour_du_mois(ref + timedelta(days=20), 8), "Trimestrielle"),
        m.Charge("CF-006", "Crédit véhicule", NOM_B2, "Emprunts", m.TYPE_DEPENSE, 450.00, _jour_du_mois(debut, 20),
                 "Mensuelle", fin=m.edate(_jour_du_mois(debut, 20), 14)),
        m.Charge("CF-007", "Licences logicielles", NOM_B1, "Informatique", m.TYPE_DEPENSE, 1200.00,
                 _jour_du_mois(ref + timedelta(days=100), 3), "Annuelle"),
        m.Charge("CF-008", "Loyer perçu sous-location", NOM_B2, "Autres recettes", m.TYPE_RECETTE, 650.00,
                 _jour_du_mois(debut, 2), "Mensuelle"),
    ]

    # Historique : l'outil a été mis en service à la date du solde initial
    params.date_ref = d0 + timedelta(days=1)
    m.generer_echeances(etat)
    params.date_ref = ref

    def prev(dj, banque, lib, cat, type_, montant, com=""):
        p = m.Prevision(m.prochain_id("P-", [x.id for x in etat.previsions], 6), _jour(ref, dj), banque,
                        lib, cat, type_, montant, com)
        etat.previsions.append(p)
        return p

    prev(-10, NOM_B1, "Facture F2026-118 Client Dupont", "Ventes / Clients", m.TYPE_RECETTE, 12000.00)
    prev(-6, NOM_B1, "Facture imprimerie Lemoine", "Fournisseurs", m.TYPE_DEPENSE, 1850.00)
    prev(-9, NOM_B2, "Carburant flotte", "Véhicules / Carburant", m.TYPE_DEPENSE, 120.00)
    prev(-5, NOM_B2, "Carburant flotte", "Véhicules / Carburant", m.TYPE_DEPENSE, 120.00)
    prev(-3, NOM_B1, "Facture F2026-119 Client Morel", "Ventes / Clients", m.TYPE_RECETTE, 3600.00,
         "Relancé le " + _jour(ref, -12).strftime("%d/%m"))
    prev(12, NOM_B1, "Facture F2026-121 Martin SA", "Ventes / Clients", m.TYPE_RECETTE, 8500.00)
    prev(18, NOM_B1, "TVA à décaisser", "TVA", m.TYPE_DEPENSE, 7400.00)
    prev(25, NOM_B1, "Acompte projet Leroy", "Ventes / Clients", m.TYPE_RECETTE, 15000.00)
    prev(30, NOM_B2, "Achat matériel informatique", "Informatique", m.TYPE_DEPENSE, 3200.00)
    prev(40, NOM_B1, "Honoraires expert-comptable", "Honoraires", m.TYPE_DEPENSE, 1500.00)
    prev(50, NOM_B1, "Solde projet Bernard", "Ventes / Clients", m.TYPE_RECETTE, 22000.00)
    prev(60, NOM_B1, "Facture Client Garnier", "Ventes / Clients", m.TYPE_RECETTE, 18000.00)
    prev(75, NOM_B1, "Facture Client Durand", "Ventes / Clients", m.TYPE_RECETTE, 4800.00)
    prev(85, NOM_B1, "Facture Client Petit", "Ventes / Clients", m.TYPE_RECETTE, 9500.00)
    m.ajouter_virement(etat, _jour(ref, -4), NOM_B1, NOM_B2, 10000.00, "Virement de trésorerie vers Beta")
    m.ajouter_virement(etat, _jour(ref, 35), NOM_B1, NOM_B2, 5000.00, "Virement de trésorerie vers Beta")
    m.generer_echeances(etat)
    return etat


def _occurrences(etat: m.Etat, ref_charge: str, debut: date, fin: date) -> list[date]:
    return sorted(p.date for p in etat.previsions if p.ref_charge == ref_charge and debut <= p.date <= fin)


def releves_demo(etat: m.Etat) -> tuple[dict, dict]:
    """Relevés bancaires fictifs : (lignes B1 par fichier CSV, lignes B2 par fichier XLSX)."""
    ref = etat.params.date_ref
    d0 = etat.params.banques[0].date_solde
    hier = ref - timedelta(days=1)

    ops_b1: list[tuple[date, str, float]] = [
        (d0, "PRLV SEPA ANCIEN PRELEVEMENT DEJA DANS LE SOLDE", -75.00),  # antérieure : ignorée
    ]
    for d in _occurrences(etat, "CF-001", d0 + timedelta(days=1), hier):
        ops_b1.append((d + timedelta(days=1), "PRLV SEPA SCI LES TILLEULS LOYER", -2400.00))
    for d in _occurrences(etat, "CF-002", d0 + timedelta(days=1), hier):
        ops_b1.append((d, "VIR SEPA SALAIRES DU MOIS", -18500.00))
    for d in _occurrences(etat, "CF-003", d0 + timedelta(days=1), hier):
        ops_b1.append((d + timedelta(days=2), "PRLV SEPA URSSAF ILE DE FRANCE", -6200.00))
    ops_b1 += [
        (_jour(ref, -9), "VIR SEPA RECU DUPONT ET FILS FACT F2026-118", 12000.00),
        (_jour(ref, -5), "VIR SEPA LEMOINE IMPRESSION", -1850.00),
        (_jour(ref, -4), "VIR SEPA VERS BANQUE BETA TRESORERIE", -10000.00),
        (_jour(ref, -20), "CB AMAZON EU SARL", -149.90),
        (_jour(ref, -16), "FRAIS TENUE DE COMPTE", -18.50),
        (_jour(ref, -13), "CB RESTAURANT LE MARAIS", -86.40),
        (_jour(ref, -2), "VIR SEPA RECU CLIENT MOREL", 3500.00),  # montant différent de la prévision (3 600)
    ]
    ops_b1.sort(key=lambda x: x[0])

    # Deux relevés qui se chevauchent : partie 1 jusqu'à J-12, partie 2 depuis J-20
    def fmt(v):
        s = f"{abs(v):,.2f}".replace(",", " ").replace(".", ",")
        return ("-" if v < 0 else "") + s

    def csv_rows(ops, du, au):
        rows = [["BANQUE ALPHA - Export des opérations"], ["Compte courant n° 00012345678"],
                [f"Période du {du:%d/%m/%Y} au {au:%d/%m/%Y}"],
                ["Date opération", "Date valeur", "Libellé", "Montant (EUR)"]]
        for d, lib, mt in ops:
            if du <= d <= au:
                rows.append([d.strftime("%d/%m/%Y"), (d + timedelta(days=1)).strftime("%d/%m/%Y"), lib, fmt(mt)])
        rows.append(["", "", f"Solde au {au:%d/%m/%Y}", ""])
        return rows

    fichiers_b1 = {
        "releve_alpha_partie1.csv": csv_rows(ops_b1, d0 - timedelta(days=3), _jour(ref, -12)),
        "releve_alpha_partie2.csv": csv_rows(ops_b1, _jour(ref, -20), hier),
    }

    ops_b2: list[tuple[date, str, str, float | None, float | None, str]] = []
    n = 1000
    for d in _occurrences(etat, "CF-004", d0 + timedelta(days=1), hier):
        ops_b2.append((d, "PRLV SEPA", "ORANGE SA FACTURE MOBILE", 89.90, None, ""))
    for d in _occurrences(etat, "CF-006", d0 + timedelta(days=1), hier):
        ops_b2.append((d + timedelta(days=1), "ECHEANCE PRET", "CREDIT AUTO N 45A12", 450.00, None, ""))
    for d in _occurrences(etat, "CF-008", d0 + timedelta(days=1), hier):
        ops_b2.append((d, "VIR SEPA RECU", "SOUS LOCATION BUREAU 2", None, 650.00, ""))
    ops_b2 += [
        (_jour(ref, -3), "VIR SEPA RECU", "BANQUE ALPHA TRESORERIE", None, 10000.00, ""),
        (_jour(ref, -7), "CARTE", "TOTALENERGIES STATION A6", 120.00, None, ""),
        (_jour(ref, -11), "CARTE", "PEAGE APRR", 14.20, None, ""),
        (_jour(ref, -11), "CARTE", "PEAGE APRR", 14.20, None, ""),
        (_jour(ref, -8), "COMMISSION", "COTISATION CARTE BUSINESS", 45.00, None, ""),
        (_jour(ref, -15), "VIR SEPA EMIS", "SARL DECO PLUS ACOMPTE", 980.00, None, ""),
    ]
    ops_b2.sort(key=lambda x: (x[0], x[2]))
    lignes_b2 = [["Date", "Libellé opération", "Détail", "Débit", "Crédit", "N° opération"]]
    for d, lib, det, deb, cre, _ in ops_b2:
        n += 1
        lignes_b2.append([d, lib, det, deb, cre, f"OP{d:%y%m%d}{n}"])
    return fichiers_b1, {"releve_beta.xlsx": lignes_b2}


def ecrire_releves(dossier_racine: str, fichiers_b1: dict, fichiers_b2: dict) -> None:
    import xlsxwriter

    d1 = os.path.join(dossier_racine, "Imports", "Banque1")
    d2 = os.path.join(dossier_racine, "Imports", "Banque2")
    os.makedirs(d1, exist_ok=True)
    os.makedirs(d2, exist_ok=True)
    for nom, lignes in fichiers_b1.items():
        with open(os.path.join(d1, nom), "w", encoding="cp1252", newline="") as f:
            w = csv.writer(f, delimiter=";", lineterminator="\r\n")
            for r in lignes:
                w.writerow(r)
    for nom, lignes in fichiers_b2.items():
        wb = xlsxwriter.Workbook(os.path.join(d2, nom))
        ws = wb.add_worksheet("Opérations")
        fdate = wb.add_format({"num_format": "dd/mm/yyyy"})
        fnum = wb.add_format({"num_format": "#,##0.00"})
        entete = wb.add_format({"bold": True})
        for i, r in enumerate(lignes):
            for j, v in enumerate(r):
                if i == 0:
                    ws.write_string(i, j, v, entete)
                elif isinstance(v, date):
                    ws.write_datetime(i, j, v, fdate)
                elif isinstance(v, float):
                    ws.write_number(i, j, v, fnum)
                elif v is not None:
                    ws.write_string(i, j, str(v))
        ws.set_column(0, 0, 12)
        ws.set_column(1, 2, 28)
        ws.set_column(3, 5, 14)
        wb.close()


def lignes_brutes_depuis_fichiers(dossier_racine: str, params_b1: dict, params_b2: dict) -> tuple[dict, dict]:
    """Relit les fichiers écrits (comme le ferait Power Query) pour les tests."""
    res = []
    for code, p in (("Banque1", params_b1), ("Banque2", params_b2)):
        d = os.path.join(dossier_racine, "Imports", code)
        res.append({f: m.lire_fichier(os.path.join(d, f), p) for f in sorted(os.listdir(d))
                    if os.path.splitext(f)[1].lower() in p["Extensions"]})
    return res[0], res[1]
