"""Automatisations : règles apprises, virements internes non prévus, prévision proche, alertes."""

from datetime import date, timedelta

from reference import moteur as m
from tests.test_scenarios import B1, B2, csv_b1, importer_b1, importer_b2, nouvel_etat, prev, xlsx_b2


def actualiser(etat, fichiers_b1=None, fichiers_b2=None):
    if fichiers_b1:
        importer_b1(etat, fichiers_b1)
    if fichiers_b2:
        importer_b2(etat, fichiers_b2)
    m.generer_echeances(etat)
    m.rechercher(etat)
    return m.traitements_automatiques(etat)


def test_regle_apprise_puis_appliquee_le_mois_suivant():
    etat = nouvel_etat()
    actualiser(etat, {"oct.csv": csv_b1([("05/11/2026", "FRAIS TENUE DE COMPTE 11/2026", "-18,50")])})
    o = m.a_traiter(etat)[0]
    m.traiter(etat, o.id, m.ACTION_CONSERVER, "Frais bancaires")
    assert [(r.mot_cle, r.categorie, r.action) for r in etat.regles] == [("FRAIS TENUE COMPTE", "Frais bancaires", m.ACTION_CONSERVER)]
    r = actualiser(etat, {"nov.csv": csv_b1([("05/12/2026", "FRAIS TENUE DE COMPTE 12/2026", "-18,50")])} | {})
    # (la date du 05/12 est postérieure à la date de calcul : sans importance pour le traitement)
    o2 = etat.operations[1]
    assert r["regles"] == 1 and (o2.statut, o2.categorie, o2.mode) == (m.OB_NON_PREVUE, "Frais bancaires", "Automatique")
    assert m.a_traiter(etat) == []


def test_regle_non_appliquee_si_prevision_proche_de_montant_different():
    etat = nouvel_etat()
    etat.regles.append(m.Regle("EDF", "Énergie", m.ACTION_CONSERVER))
    p = prev(etat, date(2026, 11, 10), B1, "Facture EDF", 120.00)
    r = actualiser(etat, {"nov.csv": csv_b1([("11/11/2026", "PRLV SEPA EDF CLIENTS PARTICULIERS", "-123,45")])})
    o = etat.operations[0]
    assert r["regles"] == 0 and o.statut == m.OB_A_TRAITER
    assert m.prevision_proche(etat, o) == p.id
    # sans prévision proche, la règle s'applique
    etat2 = nouvel_etat()
    etat2.regles.append(m.Regle("EDF", "Énergie", m.ACTION_CONSERVER))
    assert actualiser(etat2, {"nov.csv": csv_b1([("11/11/2026", "PRLV SEPA EDF", "-123,45")])})["regles"] == 1


def test_regle_la_plus_precise_et_regle_sans_action():
    etat = nouvel_etat()
    etat.regles += [m.Regle("FRAIS", "Divers"), m.Regle("FRAIS TENUE", "Frais bancaires", m.ACTION_CONSERVER),
                    m.Regle("COTIS", "Frais bancaires", m.ACTION_CONSERVER)]
    assert m.trouver_regle(etat, "FRAIS TENUE DE COMPTE").categorie == "Frais bancaires"
    assert m.trouver_regle(etat, "FRAIS DE PORT").action == ""
    assert m.trouver_regle(etat, "COMMISSION COTISATION CARTE").mot_cle == "COTIS"
    r = actualiser(etat, {"nov.csv": csv_b1([("06/11/2026", "FRAIS DE PORT", "-9,90")])})
    assert r["regles"] == 0 and etat.operations[0].statut == m.OB_A_TRAITER   # règle sans action : catégorie seulement


def test_virement_interne_non_prevu_detecte():
    etat = nouvel_etat()
    r = actualiser(etat, {"b1.csv": csv_b1([("18/11/2026", "VIR SEPA VERS BANQUE BETA", "-2 000,00"),
                                            ("19/11/2026", "CB FOURNISSEUR", "-500,00")])},
                   {"b2.xlsx": xlsx_b2([(date(2026, 11, 19), "VIR RECU", "BANQUE ALPHA", None, 2000, "R1"),
                                        (date(2026, 11, 19), "REMISE CHEQUE", "CLIENT", None, 500, "R2")])})
    assert r["virements"] == 1
    statuts = {o.libelle: o.statut for o in etat.operations}
    assert statuts["VIR SEPA VERS BANQUE BETA"] == statuts["VIR RECU BANQUE ALPHA"] == m.OB_VIREMENT
    # montants opposés mais sans libellé de virement : pas de traitement automatique
    assert statuts["CB FOURNISSEUR"] == statuts["REMISE CHEQUE CLIENT"] == m.OB_A_TRAITER
    assert m.solde_reel(etat, B1) + m.solde_reel(etat, B2) == 15000.00


def test_virement_trop_eloigne_non_automatique():
    etat = nouvel_etat()
    r = actualiser(etat, {"b1.csv": csv_b1([("10/11/2026", "VIR SEPA VERS BETA", "-700,00")])},
                   {"b2.xlsx": xlsx_b2([(date(2026, 11, 16), "VIR RECU", "ALPHA", None, 700, "R1")])})
    assert r["virements"] == 0 and len(m.a_traiter(etat)) == 2


def test_automatisations_desactivables():
    etat = nouvel_etat(apprendre_regles=False, traitement_auto=False, virements_auto=False)
    etat.regles.append(m.Regle("EDF", "Énergie", m.ACTION_CONSERVER))
    r = actualiser(etat, {"nov.csv": csv_b1([("11/11/2026", "PRLV EDF", "-80,00"),
                                             ("12/11/2026", "VIR VERS BETA", "-100,00")])},
                   {"b2.xlsx": xlsx_b2([(date(2026, 11, 12), "VIR RECU", "ALPHA", None, 100, "R1")])})
    assert r == {"virements": 0, "regles": 0}
    m.traiter(etat, etat.operations[0].id, m.ACTION_CONSERVER, "Énergie")
    assert len(etat.regles) == 1


def test_alertes_de_tresorerie():
    etat = nouvel_etat()
    prev(etat, date(2026, 12, 10), B1, "Grosse dépense", 12000.00)
    assert m.premiere_alerte(etat, B1, 0) == date(2026, 12, 10)
    assert m.premiere_alerte(etat, B2, 0) is None
    assert m.premiere_alerte(etat, None, 2000) is None          # global au plus bas : 15 000 - 12 000 = 3 000
    assert m.premiere_alerte(etat, None, 4000) == date(2026, 12, 10)
