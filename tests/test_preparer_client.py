"""Préparation du classeur client (outils/preparer_client.py) sur des données FICTIVES aux formats
réels : soldes de départ, calage des charges sur l'historique, salaires individuels, doublons,
règles apprises et simulation de la première actualisation."""

import csv
import os
from datetime import date, datetime

import xlsxwriter

from build import donnees_demo as dd
from outils.preparer_client import preparer
from reference import moteur as m

DEPART, REF = date(2026, 11, 24), date(2026, 11, 25)


def _fichiers(tmp):
    ops_b1 = []
    for mois in range(1, 12):
        ops_b1.append((date(2026, mois, 5), "VIR SCI DES TILLEULS LOYER", -1000.00))
        ops_b1.append((date(2026, mois, 24), "COTIS RYTHMEO XCCNV352 CONTRAT", -15.09))
    for d in (date(2026, 1, 7), date(2026, 4, 7), date(2026, 7, 7), date(2026, 10, 6)):
        ops_b1.append((d, "FRAIS ARRETE DE CPTE TRIMESTRE", -150.00))
    ops_b1.append((date(2026, 11, 20), "VIR INST CLIENT ALPHA FACTURE", 4200.00))
    bpgo = os.path.join(tmp, "operations.csv")
    with open(bpgo, "w", encoding="cp1252", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        for r in dd.lignes_csv_bpgo(ops_b1, date(2026, 1, 1), DEPART, 20000.00):
            w.writerow(r)

    ops_b2 = [(date(2026, 11, 2), "VIR JEAN DUPONT", 1500.50, None), (date(2026, 11, 2), "VIR MARIE CURIE", 2000.25, None),
              (date(2026, 11, 2), "VIR KONICA MINOLTA 1396", 9999.00, None),
              (date(2026, 10, 1), "VIR JEAN DUPONT", 1480.00, None),
              (date(2026, 10, 10), "PRLV SEPA ABEILLE IARD ET SANTE", 80.00, None),
              (date(2026, 11, 10), "PRLV SEPA ABEILLE IARD ET SANTE", 80.00, None),
              (date(2026, 10, 26), "VIR SEPA REMUNERATION RCH", 900.00, None),
              (date(2026, 11, 25), "VIR SEPA REMUNERATION RCH", 900.00, None)]
    cic = os.path.join(tmp, "comptes_1.xlsx")
    dd.ecrire_xlsx_cic(cic, dd.lignes_xlsx_cic(ops_b2, REF, 30000.00))

    ancien = os.path.join(tmp, "ancien.xlsx")
    wb = xlsxwriter.Workbook(ancien)
    ws = wb.add_worksheet("DEC26")
    ws.write_row(2, 0, [None, "Date", "Tiers", "Libellé", "Type", "Périodicité", "BANQUE", "Recettes Prévues",
                        "Recettes Réalisées", "Dépenses Prévues", "Dep récurrente", "Debité", "Dépenses Réalisés", "Solde"])
    lignes = [(1, "SALAIRES GLOBAL", "", "SAL", "MENSUEL", "CIC", 3500.75),
              (5, "LOYER", "TILLEULS", "FG", "MENSUEL", "BPGO", 1000.00),
              (7, "FRAIS ARRETE COMPTE", "", "BQ", "TRIMESTRE", "BPGO", 150.00),
              (25, "BPGO - prel cotis rythmeo", "", "BQ", "MENSUEL", "BPGO", 14.20),
              (24, "COTISEO", "", "BQ", "MENSUEL", "BPGO", 15.09),
              (10, "ABEILLE SANTE", "", "SAL", "TRIMESTRE", "CIC", 80.00),
              (25, "RCH", "", "DIV", "MENSUEL", "CIC", 900.00)]
    for i, (j, tiers, lib, typ, per, bq, mt) in enumerate(lignes, start=3):
        ws.write_row(i, 0, [None, j, tiers, lib, typ, per, bq, None, None, None, mt])
    ws.write_row(len(lignes) + 3, 0, [None, None, "ECHEANCE DEC", "", "DIV", "", "CIC", None, None, 7500])
    wb.close()
    return bpgo, cic, ancien


def test_preparation_complete(tmp_path):
    bpgo, cic, ancien = _fichiers(str(tmp_path))
    sortie = str(tmp_path / "client")
    r = preparer(bpgo, cic, ancien, "DEC26", 2026, 12, DEPART, sortie, REF, societe="SAS Exemple")

    # Soldes au soir du 24/11 : BPGO = fin d'export ; CIC = fin d'export + opération du 25/11 (débit 900)
    attendu_b1 = 20000.00 - 11 * 1015.09 - 600.00 + 4200.00
    assert abs(r["soldes"][0] - attendu_b1) < 0.005
    assert abs(r["soldes"][1] - 30900.00) < 0.005

    ch = {c.libelle: c for c in r["charges"]}
    # Salaires : un par personne (le virement fournisseur du même jour est exclu : chiffres)
    assert "SALAIRES GLOBAL" not in ch
    assert ch["Salaire Jean Dupont"].montant == 1500.50 and ch["Salaire Marie Curie"].montant == 2000.25
    assert ch["Salaire Jean Dupont"].premiere == date(2026, 12, 1)
    # Mensuelle déjà payée en novembre -> décembre ; mensuelle du 25 non encore passée -> novembre
    assert ch["LOYER - TILLEULS"].premiere == date(2026, 12, 5)
    assert ch["RCH"].premiere == date(2026, 11, 25)
    # Trimestrielle calée sur le dernier paiement (06/10) -> janvier
    assert (ch["FRAIS ARRETE COMPTE"].frequence, ch["FRAIS ARRETE COMPTE"].premiere) == ("Trimestrielle", date(2027, 1, 7))
    # « Trimestrielle » prélevée en réalité chaque mois -> mensuelle
    assert (ch["ABEILLE SANTE"].frequence, ch["ABEILLE SANTE"].premiere) == ("Mensuelle", date(2026, 12, 10))
    # Doublon probable (même prélèvement, montant voisin, jour voisin) désactivé
    assert not ch["BPGO - prel cotis rythmeo"].actif and ch["COTISEO"].actif
    assert [x[1] for x in r["ignorees"]] == ["ECHEANCE DEC"]

    # Règles : départ + apprises, sans mot trop général
    cles = [x.mot_cle for x in r["regles"]]
    assert "URSSAF" in cles and "VIR" not in cles and "LOYER" not in cles
    assert any("TILLEULS" in k for k in cles)

    # Première actualisation : seule l'opération du 25/11 est nouvelle, rapprochée avec certitude
    c = r["controle"]
    assert sum(v["ajoutees"] for v in c["import_b1"].values()) == 0
    assert sum(v["ajoutees"] for v in c["import_b2"].values()) == 1
    # correspondance certaine validée automatiquement (un seul bouton : Actualiser)
    assert [(p[2], p[3]) for p in c["pointees"]] == [(-900.00, "RCH")] and c["propositions"] == []
    assert abs(c["solde_reel"][1] - 30000.00) < 0.005

    assert os.path.isfile(os.path.join(sortie, "Tresorerie.xlsm"))
    assert os.listdir(os.path.join(sortie, "Imports", "Banque1")) == ["operations.csv"]
    assert os.listdir(os.path.join(sortie, "Imports", "Banque2")) == ["comptes_1.xlsx"]
    assert "Soldes de départ" in open(os.path.join(sortie, "rapport_reprise.txt"), encoding="utf-8-sig").read()
