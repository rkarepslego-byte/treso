"""Vérifie les FORMULES du classeur généré : recalcul par LibreOffice puis comparaison
avec le moteur de référence Python (soldes réels, prévisionnels, jour par jour, KPI,
suivi chronologique, statuts...).

Nécessite LibreOffice et l'interpréteur système /usr/bin/python3 (module uno).
"""

import os
import shutil
import subprocess
from datetime import date, datetime, timedelta

import openpyxl
import pytest

from build import donnees_demo as dd
from build.generer_classeur import Generateur
from reference import moteur as m
from reference.formats import FORMAT_B1, FORMAT_B2, vers_parametres_moteur

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = date(2026, 11, 25)
LO_DISPONIBLE = shutil.which("soffice") is not None and os.path.exists("/usr/bin/python3")


def etat_apres_actualisation(ref: date) -> m.Etat:
    """Démo complète : import des relevés, recherche, validations et traitements."""
    etat = dd.etat_initial(ref)
    f1, f2 = dd.releves_demo(etat)
    p1 = vers_parametres_moteur(FORMAT_B1, dd.NOM_B1)
    p2 = vers_parametres_moteur(FORMAT_B2, dd.NOM_B2)
    m.importer(etat, m.preparer_import(f1, p1), datetime(2026, 11, 25, 9, 30))
    m.importer(etat, m.preparer_import(f2, p2), datetime(2026, 11, 25, 9, 30))
    m.generer_echeances(etat)
    m.rechercher(etat)
    # L'utilisateur valide toutes les correspondances certaines sauf une (laissée en attente)
    certaines = [pr for pr in etat.propositions if pr.niveau == m.NIVEAU_CERTAINE]
    m.valider(etat, [(pr.op.id, pr.prev.id, "Oui") for pr in certaines[:-1]])
    # Traite deux opérations non prévues, en laisse les autres « à traiter »
    a_traiter = m.a_traiter(etat)
    m.traiter(etat, a_traiter[0].id, "Conserver comme non prévue", "Fournitures")
    m.traiter(etat, a_traiter[1].id, "Créer charge mensuelle", "Frais bancaires")
    m.rechercher(etat)
    etat.date_calcul_fixe = True
    return etat


def generer_et_recalculer(etat: m.Etat, dossier: str) -> openpyxl.Workbook:
    chemin = os.path.join(dossier, "test.xlsm")
    Generateur(chemin, etat).generer()
    sortie = os.path.join(dossier, "recalcule.xlsx")
    subprocess.run(["/usr/bin/python3", os.path.join(RACINE, "tests", "outils", "recalcul_lo.py"), chemin, sortie],
                   check=True, timeout=180)
    return openpyxl.load_workbook(sortie, data_only=True)


@pytest.fixture(scope="module")
def classeur(tmp_path_factory):
    if not LO_DISPONIBLE:
        pytest.skip("LibreOffice indisponible")
    etat = etat_apres_actualisation(REF)
    wb = generer_et_recalculer(etat, str(tmp_path_factory.mktemp("lo")))
    return etat, wb


def valeurs_nommees(wb):
    calc = wb["Calculs"]
    return {calc.cell(r, 1).value: calc.cell(r, 2).value for r in range(1, 40) if calc.cell(r, 1).value}


def approx(a, b):
    return abs((a or 0) - (b or 0)) < 0.005


def test_aucune_erreur_de_formule(classeur):
    _, wb = classeur
    erreurs = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")):
                    erreurs.append(f"{ws.title}!{c.coordinate}={c.value}")
    assert erreurs == []


def test_soldes_reels_et_previsionnels(classeur):
    etat, wb = classeur
    v = valeurs_nommees(wb)
    b1, b2 = dd.NOM_B1, dd.NOM_B2
    assert approx(v["SoldeReel_B1"], m.solde_reel(etat, b1))
    assert approx(v["SoldeReel_B2"], m.solde_reel(etat, b2))
    assert approx(v["SoldeReel_Total"], m.solde_reel(etat, b1) + m.solde_reel(etat, b2))
    for h in (30, 60, 90):
        assert approx(v[f"Prev_B1_{h}"], m.solde_previsionnel(etat, b1, h)), h
        assert approx(v[f"Prev_B2_{h}"], m.solde_previsionnel(etat, b2, h)), h
        assert approx(v[f"Prev_Total_{h}"], m.solde_previsionnel(etat, b1, h) + m.solde_previsionnel(etat, b2, h)), h
    assert approx(v["SoldePrev_B1"], m.solde_previsionnel(etat, b1, 30))  # horizon par défaut 30 j


def test_soldes_jour_par_jour(classeur):
    etat, wb = classeur
    calc = wb["Calculs"]
    s1 = m.serie_journaliere(etat, dd.NOM_B1, 90)
    s2 = m.serie_journaliere(etat, dd.NOM_B2, 90)
    for k in range(91):
        assert calc.cell(2 + k, 5).value.date() == REF + timedelta(days=k)
        assert approx(calc.cell(2 + k, 7).value, s1[k]), k
        assert approx(calc.cell(2 + k, 9).value, s2[k]), k
        assert approx(calc.cell(2 + k, 10).value, s1[k] + s2[k]), k
    globale = [a + b for a, b in zip(s1, s2)]
    v = valeurs_nommees(wb)
    assert approx(v["PointBas"], min(globale[:31]))
    assert v["PointBas_Date"].date() == REF + timedelta(days=globale[:31].index(min(globale[:31])))


def test_indicateurs(classeur):
    etat, wb = classeur
    v = valeurs_nommees(wb)
    assert approx(v["Depenses_A_Venir"], m.total_a_venir(etat, m.TYPE_DEPENSE, 30))
    assert approx(v["Recettes_Attendues"], m.total_a_venir(etat, m.TYPE_RECETTE, 30))
    assert v["NbATraiter"] == len(m.a_traiter(etat))
    assert v["NbOpsProposees"] == len({pr.op.id for pr in etat.propositions})
    assert v["NbPropositions"] == len(etat.propositions)
    assert v["NbEchues"] == sum(1 for p in etat.previsions if m.statut_suivi(etat.params, p) == m.PREV_A_CONTROLER)
    non_prevues = [o for o in etat.operations if o.statut in (m.OB_NON_PREVUE, m.OB_A_TRAITER)
                   and o.montant < 0 and m.op_incluse(etat.params, o)]
    assert approx(v["NonPrevues_Montant"], -sum(o.montant for o in non_prevues))
    assert v["NonPrevues_Nb"] == len(non_prevues)
    assert v["DernierReleve_B1"].date() == max(o.date for o in etat.operations if o.banque == dd.NOM_B1)


def test_colonnes_calculees_des_previsions(classeur):
    etat, wb = classeur
    ws = wb["Saisie"]
    entetes = {ws.cell(10, c).value: c for c in range(2, 40) if ws.cell(10, c).value}
    par_id = {p.id: p for p in etat.previsions}
    vus = 0
    for r in range(11, 11 + len(etat.previsions)):
        p = par_id[ws.cell(r, entetes["ID"]).value]
        assert approx(ws.cell(r, entetes["Montant signé"]).value, p.montant_signe)
        assert ws.cell(r, entetes["Suivi"]).value == m.statut_suivi(etat.params, p)
        assert approx(ws.cell(r, entetes["Montant à venir"]).value, m.montant_a_venir(etat.params, p))
        vus += 1
    assert vus == len(etat.previsions)
    # lignes libres : aucune formule ne doit afficher de valeur parasite
    r = 11 + len(etat.previsions)
    assert ws.cell(r, entetes["Suivi"]).value in (None, "")
    assert ws.cell(r, entetes["Contrôle"]).value in (None, "")
    assert ws.cell(r, entetes["Montant à venir"]).value == 0


def test_suivi_chronologique(classeur):
    etat, wb = classeur
    ws = wb["Suivi prévisionnel"]
    # Lignes attendues : opérations incluses puis prévisions restantes, triées par date effective
    attendu = []
    for i, o in enumerate(etat.operations):
        if m.op_incluse(etat.params, o):
            attendu.append(((o.date, 0, i), o.montant, o.id))
    for i, p in enumerate(etat.previsions):
        mav = m.montant_a_venir(etat.params, p)
        if mav:
            attendu.append(((max(p.date, REF), 1, i), mav, p.id))
    attendu.sort()
    ouverture = sum(b.solde_initial for b in etat.params.banques)
    du, au = REF - timedelta(days=30), REF + timedelta(days=90)
    solde = ouverture
    lignes = []
    for (d, _t, _i), mt, id_ in attendu:
        solde += mt
        if du <= d <= au:
            lignes.append((id_, round(mt, 2), round(solde, 2)))
    lus = []
    r = 9
    while ws.cell(r, 11).value:
        lus.append((ws.cell(r, 11).value, round(ws.cell(r, 7).value, 2), round(ws.cell(r, 9).value, 2)))
        r += 1
    assert lus == lignes
    # Le dernier solde de la vue = solde prévisionnel global à la fin de la période
    jours = (au - REF).days
    assert approx(lus[-1][2], m.solde_previsionnel(etat, dd.NOM_B1, jours) + m.solde_previsionnel(etat, dd.NOM_B2, jours))


def test_prochaine_echeance_des_charges(classeur):
    etat, wb = classeur
    ws = wb["Charges fixes"]
    entetes = {ws.cell(7, c).value: c for c in range(2, 30) if ws.cell(7, c).value}
    for r in range(8, 8 + len(etat.charges)):
        cid = ws.cell(r, entetes["ID"]).value
        c = next(x for x in etat.charges if x.id == cid)
        pas = m.FREQUENCES[c.frequence]
        k = 0
        while m.edate(c.premiere, k * pas) < REF:
            k += 1
        attendu = m.edate(c.premiere, k * pas)
        assert ws.cell(r, entetes["Prochaine échéance"]).value.date() == attendu, cid


def test_classeur_demo_livre_sans_erreur(tmp_path):
    if not LO_DISPONIBLE:
        pytest.skip("LibreOffice indisponible")
    etat = dd.etat_initial(REF)
    etat.date_calcul_fixe = True
    wb = generer_et_recalculer(etat, str(tmp_path))
    v = valeurs_nommees(wb)
    assert approx(v["Prev_Total_90"], m.solde_previsionnel(etat, dd.NOM_B1, 90) + m.solde_previsionnel(etat, dd.NOM_B2, 90))
    assert v["NbATraiter"] == 0


def test_classeur_vide_mise_en_service(tmp_path):
    """État après remise à zéro (mise en service) : aucune opération, prévision ni charge."""
    if not LO_DISPONIBLE:
        pytest.skip("LibreOffice indisponible")
    etat = m.Etat(dd.parametres_demo(REF))
    etat.date_calcul_fixe = True
    wb = generer_et_recalculer(etat, str(tmp_path))
    erreurs = [f"{ws.title}!{c.coordinate}" for ws in wb.worksheets for row in ws.iter_rows() for c in row
               if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    assert erreurs == []
    v = valeurs_nommees(wb)
    total = sum(b.solde_initial for b in etat.params.banques)
    assert v["SoldeReel_Total"] == total and v["Prev_Total_90"] == total and v["PointBas"] == total
    assert v["NbATraiter"] == 0 and v["Suivi_NbLignes"] == 0
