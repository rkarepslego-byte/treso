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
from datetime import date, datetime, timedelta

from reference import moteur as m

NOM_B1, NOM_B2 = "BPGO", "CIC"


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
    from build.schema import REGLES_DEFAUT
    params = parametres_demo(ref)
    d0 = params.banques[0].date_solde
    etat = m.Etat(params)
    etat.regles = [m.Regle(k, c, a) for k, c, a in REGLES_DEFAUT]
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
    m.ajouter_virement(etat, _jour(ref, -4), NOM_B1, NOM_B2, 10000.00, "Virement de trésorerie vers CIC")
    m.ajouter_virement(etat, _jour(ref, 35), NOM_B1, NOM_B2, 5000.00, "Virement de trésorerie vers CIC")
    m.generer_echeances(etat)
    return etat


def _occurrences(etat: m.Etat, ref_charge: str, debut: date, fin: date) -> list[date]:
    return sorted(p.date for p in etat.previsions if p.ref_charge == ref_charge and debut <= p.date <= fin)


# ---------------------------------------------------------------------------
# Formats réels des exports (BPGO : CSV ; CIC : classeur .xlsx)
# ---------------------------------------------------------------------------
FEUILLE_CIC = "Cpt 14270 00020886002"
ENTETE_BPGO = ["Date comptable", "Libelle simplifie", "Reference", "Informations complementaires", "Type operation",
               "Debit", "Credit", "Date operation", "Date de valeur", "Pointage"]


def _fmt_bpgo(v: float) -> str:
    return ("-" if v < 0 else "+") + f"{abs(v):.2f}".replace(".", ",")


def _reference_bpgo(d: date, libelle: str, montant: float) -> str:
    """Référence d'opération stable (les paiements carte n'en ont pas, comme dans les vrais exports)."""
    if libelle.upper().startswith(("CB ", "CARTE", "PAIEMENT")):
        return ""
    import hashlib
    h = hashlib.sha1(f"{d}|{libelle}|{montant:.2f}".encode()).hexdigest().upper()
    return h[:7]


def _type_bpgo(libelle: str, montant: float) -> str:
    l = libelle.upper()
    if l.startswith("VIR"):
        return "Virement" if montant > 0 else "Remise virement"
    if l.startswith("PRLV"):
        return "Prelevement SDD"
    if l.startswith(("CB", "CARTE", "PAIEMENT")):
        return "Paiement CB"
    return "Frais et extournes"


def lignes_csv_bpgo(ops, du: date, au: date, solde_debut: float = 0.0) -> list[list]:
    """Lignes d'un export CSV BPGO : ops = [(date, libellé, montant[, informations])]."""
    retenues = [o for o in ops if du <= o[0] <= au]
    solde_fin = solde_debut + sum(o[2] for o in retenues)
    rows = [[f"Code de la banque : 13807", "Code de l'agence : 00000", f"Date de début de téléchargement : {du:%d/%m/%Y}",
             f"Date de fin de téléchargement : {au:%d/%m/%Y}"],
            ["Numéro de compte : 00000000000", "Nom du compte : Compte Courant", "Devise : EUR"],
            [],
            ["Solde en fin de période", "", "", f"{solde_fin:.2f}".replace(".", ",")],
            list(ENTETE_BPGO)]
    for o in sorted(retenues, key=lambda x: x[0], reverse=True):
        d, lib, mt = o[0], o[1], o[2]
        info = o[3] if len(o) > 3 else ""
        rows.append([f"{d:%d/%m/%Y}", lib, _reference_bpgo(d, lib, mt), info, _type_bpgo(lib, mt),
                     _fmt_bpgo(mt) if mt < 0 else "", _fmt_bpgo(mt) if mt > 0 else "", f"{d:%d/%m/%Y}",
                     f"{d:%d/%m/%Y}", "Non"])
    rows.append(["Solde en début de période", "", "", f"{solde_debut:.2f}".replace(".", ",")])
    return rows


def lignes_xlsx_cic(ops, au: date, solde_fin: float = 0.0) -> list[list]:
    """Lignes de l'onglet du compte d'un export CIC : ops = [(date, libellé, débit, crédit)]."""
    rows = [[f"Situation de votre compte C/C Contrat Pro Global EXEMPLE (EUR) au {au:%d/%m/%Y}"],
            ["R.I.B. : 00000 00000 00000000000"],
            [None, None, None, "Solde initial : ", None, None, "EUR"],
            ["Liste de vos comptes"],
            ["Date", "Valeur", "Libellé", "Débit", "Crédit", "Solde", "Dev"]]
    for d, lib, deb, cre in sorted(ops, key=lambda x: (x[0], x[1])):
        rows.append([d, d, lib, -abs(deb) if deb else None, abs(cre) if cre else None, None, "EUR"])
    rows += [[None] * 7, [None, None, None, f"Solde au {au:%d/%m/%Y} : ", None, solde_fin, "EUR"],
             ["Liste de vos comptes"]]
    return rows


def releves_demo(etat: m.Etat) -> tuple[dict, dict]:
    """Relevés bancaires fictifs, aux formats réels : (lignes BPGO par fichier CSV, lignes CIC par fichier XLSX)."""
    ref = etat.params.date_ref
    d0 = etat.params.banques[0].date_solde
    hier = ref - timedelta(days=1)

    ops_b1: list[tuple] = [
        (d0, "PRLV SEPA ANCIEN PRELEVEMENT DEJA DANS LE SOLDE", -75.00),  # antérieure : ignorée
    ]
    for d in _occurrences(etat, "CF-001", d0 + timedelta(days=1), hier):
        ops_b1.append((d + timedelta(days=1), "PRLV SEPA SCI LES TILLEULS", -2400.00, "LOYER BUREAUX"))
    for d in _occurrences(etat, "CF-002", d0 + timedelta(days=1), hier):
        ops_b1.append((d, "VIR SALAIRES DU MOIS", -18500.00, "Virement de Societe Exemple"))
    for d in _occurrences(etat, "CF-003", d0 + timedelta(days=1), hier):
        ops_b1.append((d + timedelta(days=2), "PRLV SEPA URSSAF PAYS DE LOIRE", -6200.00))
    ops_b1 += [
        (_jour(ref, -9), "VIR INST DUPONT ET FILS", 12000.00, "FACT F2026-118"),
        (_jour(ref, -5), "VIR LEMOINE IMPRESSION", -1850.00, "Virement de Societe Exemple"),
        (_jour(ref, -4), "VIR VERS CIC TRESORERIE", -10000.00),
        (_jour(ref, -20), "CB AMAZON EU SARL", -149.90),
        (_jour(ref, -16), "FRAIS TENUE DE COMPTE", -18.50),
        (_jour(ref, -13), "CB RESTAURANT LE MARAIS", -86.40),
        (_jour(ref, -2), "VIR INST CLIENT MOREL", 3500.00),  # montant différent de la prévision (3 600)
        (_jour(ref, -14), "VIR VERS CIC APPRO", -2000.00),   # virement interne non prévu
    ]
    ops_b1.sort(key=lambda x: x[0])
    solde_b1 = etat.params.banques[0].solde_initial
    fichiers_b1 = {
        "operations_partie1.csv": lignes_csv_bpgo(ops_b1, d0 - timedelta(days=3), _jour(ref, -12), solde_b1),
        "operations_partie2.csv": lignes_csv_bpgo(ops_b1, _jour(ref, -20), hier, solde_b1),
    }

    ops_b2: list[tuple] = []
    for d in _occurrences(etat, "CF-004", d0 + timedelta(days=1), hier):
        ops_b2.append((d, "PRLV SEPA ORANGE SA FACTURE MOBILE", 89.90, None))
    for d in _occurrences(etat, "CF-006", d0 + timedelta(days=1), hier):
        ops_b2.append((d + timedelta(days=1), "ECHEANCE PRET CREDIT AUTO N 45A12", 450.00, None))
    for d in _occurrences(etat, "CF-008", d0 + timedelta(days=1), hier):
        ops_b2.append((d, "VIR SEPA RECU SOUS LOCATION BUREAU 2", None, 650.00))
    ops_b2 += [
        (_jour(ref, -3), "VIR BQ POP BPGO TRESORERIE", None, 10000.00),
        (_jour(ref, -7), "PAIEMENT CB TOTALENERGIES STATION A6", 120.00, None),
        (_jour(ref, -11), "PAIEMENT CB PEAGE APRR", 14.20, None),
        (_jour(ref, -11), "PAIEMENT CB PEAGE APRR", 14.20, None),
        (_jour(ref, -8), "COMMISSION COTISATION CARTE BUSINESS", 45.00, None),
        (_jour(ref, -15), "VIR SARL DECO PLUS ACOMPTE", 980.00, None),
        (_jour(ref, -13), "VIR BQ POP BPGO APPRO", None, 2000.00),
    ]
    solde_b2 = etat.params.banques[1].solde_initial + sum((c or 0) - (d or 0) for dt, _l, d, c in ops_b2 if dt > d0)
    return fichiers_b1, {"comptes_1.xlsx": lignes_xlsx_cic(ops_b2, hier, solde_b2)}


def ecrire_releves(dossier_racine: str, fichiers_b1: dict, fichiers_b2: dict) -> None:
    import xlsxwriter

    d1 = os.path.join(dossier_racine, "Imports", "Banque1")
    d2 = os.path.join(dossier_racine, "Imports", "Banque2")
    os.makedirs(d1, exist_ok=True)
    os.makedirs(d2, exist_ok=True)
    for nom, lignes in fichiers_b1.items():
        with open(os.path.join(d1, nom), "w", encoding="cp1252", newline="") as f:
            w = csv.writer(f, delimiter=";", lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
            for r in lignes:
                w.writerow(r)
    for nom, lignes in fichiers_b2.items():
        ecrire_xlsx_cic(os.path.join(d2, nom), lignes)


def ecrire_xlsx_cic(chemin: str, lignes: list[list]) -> None:
    import xlsxwriter
    wb = xlsxwriter.Workbook(chemin)
    recap = wb.add_worksheet("Vos comptes")
    recap.write(0, 0, "Votre situation financière")
    for j, h in enumerate(["Compte", "R.I.B.", "Solde", "Dev"]):
        recap.write(2, j, h)
    recap.write_row(3, 0, ["C/C Contrat Pro Global EXEMPLE", "00000 00000 00000000000", lignes[-2][5], "EUR"])
    ws = wb.add_worksheet(FEUILLE_CIC)
    fdate = wb.add_format({"num_format": "dd/mm/yyyy"})
    for i, r in enumerate(lignes):
        for j, v in enumerate(r):
            if isinstance(v, date):
                ws.write_datetime(i, j, datetime(v.year, v.month, v.day), fdate)
            elif isinstance(v, (int, float)):
                ws.write_number(i, j, v)
            elif v is not None:
                ws.write_string(i, j, str(v))
    wb.close()


def lignes_brutes_depuis_fichiers(dossier_racine: str, params_b1: dict, params_b2: dict) -> tuple[dict, dict]:
    """Relit les fichiers écrits (comme le ferait Power Query) pour les tests."""
    res = []
    for code, p in (("Banque1", params_b1), ("Banque2", params_b2)):
        d = os.path.join(dossier_racine, "Imports", code)
        res.append({f: m.lire_fichier(os.path.join(d, f), p) for f in sorted(os.listdir(d))
                    if os.path.splitext(f)[1].lower() in p["Extensions"]})
    return res[0], res[1]
