"""Import des prévisions d'un ancien classeur (reprise de l'existant)."""

from datetime import date, datetime

from reference import moteur as m
from tests.test_scenarios import B1, B2, nouvel_etat


def test_colonnes_reconnues_et_import():
    lignes = [
        ["Prévisionnel de trésorerie 2026"],
        [],
        ["Échéance", "Intitulé", "Rubrique", "Compte", "Dépenses", "Recettes", "Remarques"],
        [datetime(2026, 12, 5), "Loyer", "Loyer", "Banque Alpha", 2400, None, ""],
        ["15/12/2026", "Facture client X", "Ventes", "beta", None, "5 000,00", "relancer"],
        [date(2026, 10, 15), "Ancien", "Divers", "Banque Alpha", 50, None, ""],   # antérieure au solde initial
        ["TOTAL", None, None, None, 2450, 5000, None],                              # ligne ignorée
    ]
    roles = m.detecter_colonnes_previsions(lignes[2])
    assert roles == {"date": 0, "debit": 4, "credit": 5, "banque": 3, "categorie": 2, "commentaire": 6, "libelle": 1}
    etat = nouvel_etat()
    r = m.importer_previsions(etat, lignes, B1)
    assert r == {"importees": 2, "ignorees": 1, "doublons": 0, "anterieures": 1}
    p1, p2 = etat.previsions
    assert (p1.banque, p1.type, p1.montant, p1.montant_signe) == (B1, m.TYPE_DEPENSE, 2400, -2400)
    assert (p2.banque, p2.type, p2.montant, p2.commentaire, p2.origine) == (B2, m.TYPE_RECETTE, 5000, "relancer", "Import")
    assert m.solde_previsionnel(etat, B1, 30) == 10000 - 2400
    # Relancer l'import ne crée pas de doublon
    r = m.importer_previsions(etat, lignes, B1)
    assert (r["importees"], r["doublons"]) == (0, 2)


def test_montant_signe_et_colonne_type():
    lignes = [["Date", "Libellé", "Montant", "Type"],
              ["03/12/2026", "Achat", -120.5, ""],
              ["04/12/2026", "Subvention", 3000, "Recette"],
              ["05/12/2026", "Assurance", 900, "Dépense"]]
    etat = nouvel_etat()
    m.importer_previsions(etat, lignes, B2)
    assert [(p.libelle, p.montant_signe, p.banque) for p in etat.previsions] == [
        ("Achat", -120.5, B2), ("Subvention", 3000.0, B2), ("Assurance", -900.0, B2)]


def test_entete_introuvable():
    import pytest
    with pytest.raises(ValueError):
        m.importer_previsions(nouvel_etat(), [["a", "b"], [1, 2]], B1)
