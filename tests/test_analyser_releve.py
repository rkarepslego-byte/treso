"""Analyse automatique des exports bancaires (outils/analyser_releve.py)."""

import os

import xlsxwriter

from build.generer_classeur import construire
from outils.analyser_releve import analyser
from reference.formats import FORMAT_B1, FORMAT_B2

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def ecrire(tmp_path, nom, texte, encodage="utf-8"):
    chemin = tmp_path / nom
    chemin.write_bytes(texte.encode(encodage))
    return str(chemin)


def test_formats_de_demonstration_retrouves(tmp_path):
    construire(str(tmp_path))
    r1 = analyser(str(tmp_path / "Imports" / "Banque1" / "releve_alpha_partie1.csv"))
    r2 = analyser(str(tmp_path / "Imports" / "Banque2" / "releve_beta.xlsx"))
    for cle in ("Separateur", "Encodage", "LignesIgnorees", "ColDate", "ColDateValeur", "FormatDate", "ColsLibelle",
                "ModeMontant", "ColMontant", "SeparateurDecimal"):
        assert str(r1["params"][cle]) == str(FORMAT_B1[cle]), cle
    for cle in ("LignesIgnorees", "ColDate", "ColsLibelle", "ModeMontant", "ColDebit", "ColCredit", "ColReference"):
        assert str(r2["params"][cle]) == str(FORMAT_B2[cle]), cle
    assert r1["controle"]["nb_lignes"] >= 5 and not r1["controle"]["erreur"]
    assert r2["controle"]["nb_lignes"] == r2["controle"]["nb_lignes_donnees"]


def test_csv_utf8_virgule_point_decimal(tmp_path):
    f = ecrire(tmp_path, "export.csv",
               "﻿Date,Description,Amount,Balance\n2026-11-03,\"PRLV EDF, contrat 12\",-1234.56,1000.00\n"
               "2026-11-04,VIR CLIENT,2500.00,3500.00\n")
    r = analyser(f)
    p = r["params"]
    assert (p["Separateur"], p["Encodage"], p["FormatDate"], p["SeparateurDecimal"]) == (",", 65001, "yyyy-MM-dd", ".")
    assert (p["ColDate"], p["ColsLibelle"], p["ColMontant"]) == ("Date", "Description", "Amount")
    c = r["controle"]
    assert c["nb_lignes"] == 2 and c["total_debits"] == -1234.56 and c["total_credits"] == 2500.0
    assert c["apercu"][0][1] == "PRLV EDF, contrat 12"


def test_montant_et_sens_tabulation_annee_courte(tmp_path):
    f = ecrire(tmp_path, "export.txt",
               "Relevé de compte\nDate\tLibellé\tMontant\tSens\n03/11/26\tCB GARAGE\t120,50\tD\n"
               "04/11/26\tVIR RECU\t1 000,00\tC\n", "cp1252")
    r = analyser(f)
    p = r["params"]
    assert (p["Separateur"], p["Encodage"], p["LignesIgnorees"], p["FormatDate"]) == ("TAB", 1252, 1, "dd/MM/yy")
    assert (p["ModeMontant"], p["ColSens"], p["ValeursDebit"]) == ("Montant + Sens", "Sens", "D")
    c = r["controle"]
    assert c["total_debits"] == -120.5 and c["total_credits"] == 1000.0


def test_montants_tous_positifs_signales(tmp_path):
    f = ecrire(tmp_path, "export.csv", "Date;Libellé;Montant\n03/11/2026;CB A;10,00\n04/11/2026;CB B;20,00\n")
    r = analyser(f)
    assert any("Inverser le signe" in a for a in r["alertes"])


def test_numero_de_ligne_non_retenu_comme_reference(tmp_path):
    chemin = str(tmp_path / "export.xlsx")
    wb = xlsxwriter.Workbook(chemin)
    ws = wb.add_worksheet()
    for i, row in enumerate([["N° opération", "Date", "Libellé", "Débit", "Crédit"],
                             [1, "03/11/2026", "CB A", 10, None], [2, "04/11/2026", "VIR B", None, 50]]):
        for j, v in enumerate(row):
            if v is not None:
                ws.write(i, j, v)
    wb.close()
    r = analyser(chemin)
    assert r["params"]["ColReference"] == ""
    assert r["controle"]["nb_lignes"] == 2
