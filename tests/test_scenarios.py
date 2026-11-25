"""Scénarios métier (phase 7 du cahier des charges) exécutés sur le moteur de référence.

Les montants attendus sont écrits « en dur » (calculés à la main) : ils ne dépendent
pas du moteur testé.
"""

from datetime import date, datetime

import pytest

from reference import moteur as m
from reference.formats import FORMAT_B1, FORMAT_B2, vers_parametres_moteur

B1, B2 = "BPGO", "CIC"
REF = date(2026, 11, 25)
P_B1 = vers_parametres_moteur(FORMAT_B1, B1)
P_B2 = vers_parametres_moteur(FORMAT_B2, B2)


def nouvel_etat(**kw) -> m.Etat:
    params = m.Parametres(
        banques=[
            m.Banque("B1", B1, 10000.00, date(2026, 10, 31)),
            m.Banque("B2", B2, 5000.00, date(2026, 10, 31)),
        ],
        date_ref=REF, **kw,
    )
    return m.Etat(params)


def prev(etat, d, banque, libelle, montant, type_=m.TYPE_DEPENSE, categorie="Divers"):
    p = m.Prevision(m.prochain_id("P-", [x.id for x in etat.previsions], 6), d, banque,
                    libelle, categorie, type_, montant)
    etat.previsions.append(p)
    return p


def csv_b1(lignes):
    """Construit un relevé au format réel BPGO (lignes brutes CSV) : (date, libellé, montant[, infos[, référence]])."""
    rows = [["Code de la banque : 13807", "Code de l'agence : 00000", "Date de début de téléchargement : 01/11/2026",
             "Date de fin de téléchargement : 30/11/2026"],
            ["Numéro de compte : 00000000000", "Nom du compte : Compte Courant", "Devise : EUR"],
            [],
            ["Solde en fin de période", "", "", "12345,67"],
            ["Date comptable", "Libelle simplifie", "Reference", "Informations complementaires", "Type operation",
             "Debit", "Credit", "Date operation", "Date de valeur", "Pointage"]]
    for l in lignes:
        d, lib, mt = l[0], l[1], l[2]
        info = l[3] if len(l) > 3 else ""
        ref = l[4] if len(l) > 4 else ""
        brut = mt.replace(" ", "")
        debit = brut if brut.startswith("-") else ""
        credit = "" if debit else ("+" + brut.lstrip("+"))
        rows.append([d, lib, ref, info, "Virement", debit, credit, d, d, "Non"])
    rows.append(["Solde en début de période", "", "", "1000,00"])  # ligne de pied à ignorer
    return rows


def xlsx_b2(lignes):
    """Construit un relevé au format réel CIC (onglet du compte) : (date, libellé, détail, débit, crédit, réf.).

    Le CIC n'exporte ni référence ni détail séparé : le détail est accolé au libellé, la référence ignorée.
    """
    rows = [["Situation de votre compte C/C Contrat Pro Global EXEMPLE (EUR) au 30/11/2026"],
            ["R.I.B. : 00000 00000 00000000000"],
            [None, None, None, "Solde initial : ", None, None, "EUR"],
            ["Liste de vos comptes"],
            ["Date", "Valeur", "Libellé", "Débit", "Crédit", "Solde", "Dev"]]
    for d, lib, det, deb, cre, _ref in lignes:
        libelle = " ".join(x for x in (lib, det) if x)
        if isinstance(deb, (int, float)):
            deb = -abs(deb)
        elif isinstance(deb, str) and deb and not deb.startswith("-"):
            deb = "-" + deb
        rows.append([d, d, libelle, deb, cre, None, "EUR"])
    rows += [[None] * 7, [None, None, None, "Solde au 30/11/2026 : ", None, 1234.5, "EUR"], ["Liste de vos comptes"]]
    return rows


def importer_b1(etat, fichiers):
    return m.importer(etat, m.preparer_import(fichiers, P_B1), datetime(2026, 11, 25, 9, 0))


def importer_b2(etat, fichiers):
    return m.importer(etat, m.preparer_import(fichiers, P_B2), datetime(2026, 11, 25, 9, 0))


# ---------------------------------------------------------------------------
# Normalisation des relevés
# ---------------------------------------------------------------------------
def test_normalisation_formats_demo():
    t1 = m.preparer_import({"releve.csv": csv_b1([
        ("05/11/2026", "PRLV SEPA EDF", "-1 234,56"),
        ("06/11/2026", "VIR RECU CLIENT DUPONT", "2 500,00"),
    ])}, P_B1)
    assert [(l.date, l.montant, l.libelle_norm) for l in t1] == [
        (date(2026, 11, 5), -1234.56, "PRLV SEPA EDF"),
        (date(2026, 11, 6), 2500.00, "VIR RECU CLIENT DUPONT"),
    ]
    assert t1[0].cle == "B1|2026-11-05|-123456|PRLV SEPA EDF||1"

    t2 = m.preparer_import({"releve.xlsx": xlsx_b2([
        (date(2026, 11, 3), "PRLV", "Orange Business", 89.9, None, "A001"),
        (datetime(2026, 11, 4), "VIR", "Client Martin", None, "1500,00", "A002"),
    ])}, P_B2)
    assert [(l.date, l.montant, l.libelle, l.reference) for l in t2] == [
        (date(2026, 11, 3), -89.9, "PRLV Orange Business", ""),
        (date(2026, 11, 4), 1500.0, "VIR Client Martin", ""),
    ]


def test_fichier_au_mauvais_format_ne_bloque_pas_les_autres():
    transit = m.preparer_import({
        "bon.csv": csv_b1([("05/11/2026", "PRLV SEPA EDF", "-10,00")]),
        "mauvais.csv": [["a"], ["b"], ["c"], ["Colonne1", "Colonne2"], ["x", "y"]],
    }, P_B1)
    erreurs = [l for l in transit if l.erreur]
    assert len(erreurs) == 1 and erreurs[0].fichier == "mauvais.csv"
    assert FORMAT_B1["ColDate"] in erreurs[0].erreur
    etat = nouvel_etat()
    rapport = m.importer(etat, transit, datetime(2026, 11, 25))
    assert rapport["bon.csv"]["ajoutees"] == 1
    assert rapport["mauvais.csv"]["erreur"]
    assert len(etat.operations) == 1


# ---------------------------------------------------------------------------
# Scénario 1 : paiement prévu puis débité
# ---------------------------------------------------------------------------
def test_s1_paiement_prevu_puis_debite():
    etat = nouvel_etat()
    p = prev(etat, date(2026, 11, 5), B1, "Loyer bureau", 1500.00, categorie="Loyer")
    assert m.solde_previsionnel(etat, B1, 30) == 8500.00  # 10 000 - 1 500 prévu

    importer_b1(etat, {"nov.csv": csv_b1([("06/11/2026", "PRLV SEPA SCI DU PARC LOYER NOV", "-1 500,00")])})
    assert m.solde_reel(etat, B1) == 8500.00
    # Avant recherche : l'opération est dans le réel ET la prévision est encore due
    assert m.solde_previsionnel(etat, B1, 30) == 7000.00

    props = m.rechercher(etat)
    assert len(props) == 1 and props[0].niveau == m.NIVEAU_CERTAINE
    # Proposition en attente : la prévision n'est plus comptée -> pas de double comptage
    assert m.solde_previsionnel(etat, B1, 30) == 8500.00
    assert m.statut_suivi(etat.params, p) == m.PREV_A_CONTROLER

    m.valider_certaines(etat)
    assert p.statut == m.PREV_POINTE and p.id_operation == "OB-000001"
    assert etat.op("OB-000001").statut == m.OB_POINTE
    assert etat.op("OB-000001").categorie == "Loyer"
    assert m.solde_reel(etat, B1) == 8500.00
    assert m.solde_previsionnel(etat, B1, 30) == 8500.00
    assert etat.propositions == []

    # Ré-actualisation : rien n'est reproposé ni compté deux fois
    importer_b1(etat, {"nov.csv": csv_b1([("06/11/2026", "PRLV SEPA SCI DU PARC LOYER NOV", "-1 500,00")])})
    m.rechercher(etat)
    assert etat.propositions == [] and len(etat.operations) == 1
    assert m.solde_previsionnel(etat, B1, 30) == 8500.00

    # Annulation du rapprochement puis nouvelle proposition
    assert m.delier(etat, "OB-000001")
    assert p.statut == m.PREV_PREVU and len(etat.propositions) == 1


# ---------------------------------------------------------------------------
# Scénario 2 : dépense non prévue
# ---------------------------------------------------------------------------
def test_s2_depense_non_prevue():
    etat = nouvel_etat()
    importer_b1(etat, {"nov.csv": csv_b1([("10/11/2026", "CB AMAZON PAYMENTS", "-89,90")])})
    m.rechercher(etat)
    assert etat.propositions == []
    liste = m.a_traiter(etat)
    assert [o.id for o in liste] == ["OB-000001"]
    assert m.traiter(etat, "OB-000001", "Conserver comme non prévue", "Fournitures", "Câbles") == "ok"
    o = etat.op("OB-000001")
    assert (o.statut, o.categorie, o.commentaire) == (m.OB_NON_PREVUE, "Fournitures", "Câbles")
    assert m.a_traiter(etat) == []
    assert m.solde_reel(etat, B1) == 9910.10
    assert m.solde_previsionnel(etat, B1, 90) == 9910.10


def test_s2bis_depense_non_prevue_devient_charge_recurrente():
    etat = nouvel_etat()
    importer_b1(etat, {"nov.csv": csv_b1([("12/11/2026", "PRLV SEPA OVH CLOUD", "-59,99")])})
    m.rechercher(etat)
    m.traiter(etat, "OB-000001", "Créer charge mensuelle", "Informatique")
    c = etat.charges[0]
    assert (c.id, c.montant, c.premiere, c.frequence, c.type) == ("CF-001", 59.99, date(2026, 12, 12), "Mensuelle", m.TYPE_DEPENSE)
    echeances = sorted(p.date for p in etat.previsions if p.ref_charge == "CF-001")
    assert echeances[0] == date(2026, 12, 12) and len(echeances) == 12
    assert m.solde_previsionnel(etat, B1, 30) == round(10000 - 59.99 - 59.99, 2)


# ---------------------------------------------------------------------------
# Scénario 3 : paiement identique à deux dates différentes
# ---------------------------------------------------------------------------
def test_s3_paiement_identique_deux_dates():
    etat = nouvel_etat()
    p_nov = prev(etat, date(2026, 11, 3), B1, "Abonnement logiciel", 200.00)
    p_dec = prev(etat, date(2026, 12, 3), B1, "Abonnement logiciel", 200.00)
    importer_b1(etat, {"nov.csv": csv_b1([("04/11/2026", "PRLV SEPA ABONNEMENT LOGICIEL", "-200,00")])})
    props = m.rechercher(etat)
    assert len(props) == 1 and props[0].prev is p_nov and props[0].niveau == m.NIVEAU_CERTAINE
    m.valider_certaines(etat)
    assert p_nov.statut == m.PREV_POINTE and p_dec.statut == m.PREV_PREVU
    assert m.solde_previsionnel(etat, B1, 30) == 9600.00  # 10 000 - 200 réel - 200 (déc.)


def test_s3bis_deux_candidats_dans_la_tolerance_reste_incertain():
    etat = nouvel_etat()
    p_a = prev(etat, date(2026, 11, 3), B1, "Abonnement logiciel", 200.00)
    p_b = prev(etat, date(2026, 11, 9), B1, "Abonnement logiciel", 200.00)
    importer_b1(etat, {"nov.csv": csv_b1([("06/11/2026", "PRLV SEPA ABONNEMENT LOGICIEL", "-200,00")])})
    props = m.rechercher(etat)
    assert len(props) == 2 and {pr.niveau for pr in props} == {m.NIVEAU_INCERTAINE}
    assert [pr.prev for pr in props if pr.suggere] == [p_a]  # égalité de score -> date la plus ancienne
    # Jamais de validation automatique d'une correspondance ambiguë
    assert m.valider_certaines(etat)["liees"] == 0
    # Une seule des deux prévisions est provisoirement neutralisée : pas de double comptage
    assert m.solde_previsionnel(etat, B1, 30) == 9600.00
    m.valider(etat, [(props[1].op.id, p_b.id, "Oui")])
    assert p_b.statut == m.PREV_POINTE and p_a.statut == m.PREV_PREVU
    assert m.solde_previsionnel(etat, B1, 30) == 9600.00


# ---------------------------------------------------------------------------
# Scénario 4 : deux paiements identiques légitimes
# ---------------------------------------------------------------------------
def test_s4_deux_paiements_identiques_legitimes():
    etat = nouvel_etat()
    releve = csv_b1([("15/11/2026", "CARTE TOTAL STATION", "-45,00"),
                     ("15/11/2026", "CARTE TOTAL STATION", "-45,00")])
    r = importer_b1(etat, {"nov.csv": releve})
    assert r["nov.csv"]["ajoutees"] == 2
    assert [o.cle.split("|")[-1] for o in etat.operations] == ["1", "2"]
    assert m.solde_reel(etat, B1) == 9910.00
    r = importer_b1(etat, {"nov.csv": releve})
    assert (r["nov.csv"]["ajoutees"], r["nov.csv"]["doublons"]) == (0, 2)

    p = prev(etat, date(2026, 11, 15), B1, "Carburant", 45.00)
    props = m.rechercher(etat)
    assert len(props) == 2 and props[0].niveau == m.NIVEAU_INCERTAINE
    m.valider(etat, [(props[0].op.id, p.id, "Oui")])
    assert [o.id for o in m.a_traiter(etat)] == ["OB-000002"]
    assert m.solde_previsionnel(etat, B1, 30) == 9910.00


# ---------------------------------------------------------------------------
# Scénario 5 : virement entre les deux banques
# ---------------------------------------------------------------------------
def test_s5_virement_interne_prevu():
    etat = nouvel_etat()
    sortant, entrant = m.ajouter_virement(etat, date(2026, 11, 20), B1, B2, 5000.00, "Virement trésorerie")
    assert (sortant.montant_signe, entrant.montant_signe) == (-5000.00, 5000.00)
    assert sortant.lien_virement == entrant.id
    # Le virement n'est ni une dépense ni une recette
    assert m.total_a_venir(etat, m.TYPE_DEPENSE, 30) == 0
    assert m.total_a_venir(etat, m.TYPE_RECETTE, 30) == 0
    total_prev = m.solde_previsionnel(etat, B1, 30) + m.solde_previsionnel(etat, B2, 30)
    assert (m.solde_previsionnel(etat, B1, 30), m.solde_previsionnel(etat, B2, 30), total_prev) == (5000.00, 10000.00, 15000.00)

    importer_b1(etat, {"b1.csv": csv_b1([("20/11/2026", "VIR SEPA VERS BANQUE BETA", "-5 000,00")])})
    importer_b2(etat, {"b2.xlsx": xlsx_b2([(date(2026, 11, 21), "VIR BQ POP", "BPGO", None, 5000, "X1")])})
    props = m.rechercher(etat)
    assert len(props) == 2 and all(pr.niveau == m.NIVEAU_CERTAINE for pr in props)
    m.valider_certaines(etat)
    assert (m.solde_reel(etat, B1), m.solde_reel(etat, B2)) == (5000.00, 10000.00)
    assert m.solde_reel(etat, B1) + m.solde_reel(etat, B2) == 15000.00
    assert m.solde_previsionnel(etat, B1, 30) + m.solde_previsionnel(etat, B2, 30) == 15000.00


def test_s5bis_virement_interne_non_prevu():
    etat = nouvel_etat()
    importer_b1(etat, {"b1.csv": csv_b1([("18/11/2026", "VIR SEPA VERS BANQUE BETA", "-1 000,00")])})
    importer_b2(etat, {"b2.xlsx": xlsx_b2([(date(2026, 11, 19), "VIR BQ POP", "BPGO", None, 1000, "X2")])})
    m.rechercher(etat)
    o1, o2 = m.a_traiter(etat)
    assert m.suggestion_virement(etat, o1) == o2.id
    m.traiter(etat, o1.id, "Virement interne")
    m.traiter(etat, o2.id, "Virement interne")
    assert (o1.statut, o2.statut) == (m.OB_VIREMENT, m.OB_VIREMENT)
    assert m.solde_reel(etat, B1) + m.solde_reel(etat, B2) == 15000.00


# ---------------------------------------------------------------------------
# Scénario 6 : charge fixe mensuelle
# ---------------------------------------------------------------------------
def test_s6_charge_fixe_mensuelle_fin_de_mois():
    etat = nouvel_etat()
    etat.charges.append(m.Charge("CF-001", "Loyer entrepôt", B1, "Loyer", m.TYPE_DEPENSE, 1200.00,
                                 date(2026, 1, 31), "Mensuelle"))
    r = m.generer_echeances(etat)
    dates = sorted(p.date for p in etat.previsions)
    assert r["creees"] == 12
    assert dates[:4] == [date(2026, 11, 30), date(2026, 12, 31), date(2027, 1, 31), date(2027, 2, 28)]
    assert dates[-1] == date(2027, 10, 31)
    ids = sorted(p.id for p in etat.previsions)
    # Idempotence : une nouvelle génération ne change rien (mêmes ID)
    assert m.generer_echeances(etat) == {"creees": 0, "supprimees": 0}
    assert sorted(p.id for p in etat.previsions) == ids
    assert m.solde_previsionnel(etat, B1, 30) == 10000 - 1200 * 1  # 30/11 dans J+30, 31/12 non
    assert m.solde_previsionnel(etat, B1, 90) == 10000 - 1200 * 3


def test_s6bis_frequences_et_date_de_fin():
    etat = nouvel_etat()
    etat.charges += [
        m.Charge("CF-001", "TVA", B1, "Impôts", m.TYPE_DEPENSE, 3000, date(2026, 1, 20), "Trimestrielle"),
        m.Charge("CF-002", "Assurance", B2, "Assurance", m.TYPE_DEPENSE, 900, date(2026, 3, 1), "Semestrielle"),
        m.Charge("CF-003", "CFE", B1, "Impôts", m.TYPE_DEPENSE, 1100, date(2025, 12, 15), "Annuelle"),
        m.Charge("CF-004", "Crédit véhicule", B1, "Emprunt", m.TYPE_DEPENSE, 450, date(2026, 6, 5), "Mensuelle",
                 fin=date(2027, 2, 5)),
        m.Charge("CF-005", "Ancien contrat", B1, "Divers", m.TYPE_DEPENSE, 99, date(2026, 1, 1), "Mensuelle",
                 actif=False),
    ]
    m.generer_echeances(etat)
    par = {}
    for p in etat.previsions:
        par.setdefault(p.ref_charge, []).append(p.date)
    assert sorted(par["CF-001"]) == [date(2027, 1, 20), date(2027, 4, 20), date(2027, 7, 20), date(2027, 10, 20)]
    assert sorted(par["CF-002"]) == [date(2027, 3, 1), date(2027, 9, 1)]
    assert sorted(par["CF-003"]) == [date(2026, 12, 15)]
    assert sorted(par["CF-004"]) == [date(2026, 12, 5), date(2027, 1, 5), date(2027, 2, 5)]
    assert "CF-005" not in par


# ---------------------------------------------------------------------------
# Scénario 7 : modification d'une charge fixe (pas d'effet rétroactif)
# ---------------------------------------------------------------------------
def test_s7_modification_charge_fixe():
    etat = nouvel_etat()
    c = m.Charge("CF-001", "Loyer entrepôt", B1, "Loyer", m.TYPE_DEPENSE, 1200.00, date(2026, 11, 20), "Mensuelle")
    etat.charges.append(c)
    m.generer_echeances(etat)

    def ech(d):
        return next(p for p in etat.previsions if p.ref_charge == "CF-001" and p.date == d)

    # L'échéance de novembre est débitée et pointée
    importer_b1(etat, {"nov.csv": csv_b1([("20/11/2026", "PRLV SEPA SCI ENTREPOT", "-1 200,00")])})
    m.rechercher(etat)
    m.valider_certaines(etat)
    assert ech(date(2026, 11, 20)).statut == m.PREV_POINTE
    # L'utilisateur modifie à la main l'échéance de janvier et annule celle de mars
    ech(date(2027, 1, 20)).montant = 1250.00
    ech(date(2027, 3, 20)).statut = m.PREV_ANNULE

    c.montant = 1300.00  # hausse de loyer
    m.generer_echeances(etat)
    assert ech(date(2026, 11, 20)).montant == 1200.00  # réalisée : inchangée
    assert ech(date(2026, 12, 20)).montant == 1300.00  # future : nouveau montant
    assert ech(date(2027, 1, 20)).montant == 1250.00   # modifiée à la main : conservée
    assert ech(date(2027, 3, 20)).statut == m.PREV_ANNULE  # annulée : pas ressuscitée
    assert len([p for p in etat.previsions if p.date == date(2027, 3, 20)]) == 1

    # Changement de jour d'échéance : pas de doublon sur le mois déjà réalisé
    c.premiere = date(2026, 11, 25)
    m.generer_echeances(etat)
    dates = sorted(p.date for p in etat.previsions if p.ref_charge == "CF-001")
    assert date(2026, 11, 25) not in dates and date(2026, 12, 20) not in dates
    assert date(2026, 12, 25) in dates and date(2027, 1, 25) not in dates  # janvier couvert par l'échéance modifiée
    assert ech(date(2026, 11, 20)).statut == m.PREV_POINTE
    assert m.solde_reel(etat, B1) == 8800.00

    # Désactivation : les échéances futures non figées disparaissent, les figées restent
    c.actif = False
    m.generer_echeances(etat)
    restantes = sorted(p.date for p in etat.previsions if p.ref_charge == "CF-001")
    assert restantes == [date(2026, 11, 20), date(2027, 1, 20), date(2027, 3, 20)]


# ---------------------------------------------------------------------------
# Scénario 8 : import répété du même relevé
# ---------------------------------------------------------------------------
def test_s8_import_repete():
    etat = nouvel_etat()
    releve = {"oct-nov.csv": csv_b1([("02/11/2026", "VIR RECU CLIENT A", "1 000,00"),
                                     ("03/11/2026", "PRLV URSSAF", "-2 345,00"),
                                     ("04/11/2026", "CB RESTAURANT", "-32,50")])}
    assert importer_b1(etat, releve)["oct-nov.csv"]["ajoutees"] == 3
    for _ in range(3):
        r = importer_b1(etat, releve)["oct-nov.csv"]
        assert (r["ajoutees"], r["doublons"]) == (0, 3)
    assert len(etat.operations) == 3
    assert m.solde_reel(etat, B1) == 8622.50
    assert len(etat.historique) == 4  # historique conservé à chaque import


# ---------------------------------------------------------------------------
# Scénario 9 : relevés qui se chevauchent
# ---------------------------------------------------------------------------
def test_s9_releves_qui_se_chevauchent():
    a = csv_b1([("02/11/2026", "VIR RECU CLIENT A", "1 000,00"),
                ("12/11/2026", "CB PEAGE", "-12,40"),
                ("12/11/2026", "CB PEAGE", "-12,40"),
                ("13/11/2026", "PRLV EDF", "-150,00")])
    b = csv_b1([("12/11/2026", "CB PEAGE", "-12,40"),
                ("12/11/2026", "CB PEAGE", "-12,40"),
                ("13/11/2026", "PRLV EDF", "-150,00"),
                ("20/11/2026", "VIR RECU CLIENT B", "500,00")])
    etat = nouvel_etat()
    importer_b1(etat, {"a.csv": a})
    r = importer_b1(etat, {"a.csv": a, "b.csv": b})
    assert (r["b.csv"]["ajoutees"], r["b.csv"]["doublons"]) == (1, 3)
    assert len(etat.operations) == 5
    assert m.solde_reel(etat, B1) == round(10000 + 1000 - 12.40 * 2 - 150 + 500, 2)

    # Ordre inverse, tout dans le même import : même résultat
    etat2 = nouvel_etat()
    importer_b1(etat2, {"b.csv": b, "a.csv": a})
    assert len(etat2.operations) == 5 and m.solde_reel(etat2, B1) == m.solde_reel(etat, B1)


def test_s9bis_ligne_identique_supplementaire_dans_le_second_releve():
    """Le 2e relevé contient 3 péages identiques là où le 1er n'en avait que 2 : 1 seul ajout."""
    a = csv_b1([("12/11/2026", "CB PEAGE", "-12,40"), ("12/11/2026", "CB PEAGE", "-12,40")])
    b = csv_b1([("12/11/2026", "CB PEAGE", "-12,40")] * 3)
    etat = nouvel_etat()
    importer_b1(etat, {"a.csv": a})
    r = importer_b1(etat, {"b.csv": b})
    assert (r["b.csv"]["ajoutees"], r["b.csv"]["doublons"]) == (1, 2)


# ---------------------------------------------------------------------------
# Scénario 10 : libellé bancaire différent du libellé prévisionnel
# ---------------------------------------------------------------------------
def test_s10_libelle_different():
    etat = nouvel_etat()
    p = prev(etat, date(2026, 11, 12), B1, "Assurance locaux", 320.00)
    importer_b1(etat, {"nov.csv": csv_b1([("14/11/2026", "PRLV SEPA AXA FRANCE IARD CONTRAT 1234", "-320,00")])})
    props = m.rechercher(etat)
    assert len(props) == 1 and props[0].sim < 0.4
    assert props[0].niveau == m.NIVEAU_CERTAINE  # unique et écart de 2 jours <= 3
    m.valider_certaines(etat)
    assert p.statut == m.PREV_POINTE


def test_s10bis_similarite_departage_les_candidats():
    etat = nouvel_etat()
    p_orange = prev(etat, date(2026, 11, 8), B1, "Abonnement Orange", 100.00)
    prev(etat, date(2026, 11, 9), B1, "Fournitures Bureau Vallée", 100.00)
    importer_b1(etat, {"nov.csv": csv_b1([("09/11/2026", "PRLV SEPA ORANGE SA", "-100,00")])})
    props = m.rechercher(etat)
    assert len(props) == 2 and all(pr.niveau == m.NIVEAU_INCERTAINE for pr in props)
    sugg = [pr for pr in props if pr.suggere]
    assert len(sugg) == 1 and sugg[0].prev is p_orange


def test_s10ter_unique_mais_faible_reste_incertain():
    etat = nouvel_etat()
    prev(etat, date(2026, 11, 2), B1, "Frais divers", 75.00)
    importer_b1(etat, {"nov.csv": csv_b1([("08/11/2026", "CB GARAGE MOREAU", "-75,00")])})
    props = m.rechercher(etat)
    assert len(props) == 1 and props[0].niveau == m.NIVEAU_INCERTAINE  # 6 j d'écart, libellés différents


# ---------------------------------------------------------------------------
# Règles transverses
# ---------------------------------------------------------------------------
def test_operations_anterieures_au_solde_initial_ignorees():
    etat = nouvel_etat()
    r = importer_b1(etat, {"oct.csv": csv_b1([("31/10/2026", "PRLV DEJA DANS LE SOLDE", "-500,00"),
                                              ("01/11/2026", "PRLV NOUVEAU", "-100,00")])})
    assert (r["oct.csv"]["anterieures"], r["oct.csv"]["ajoutees"]) == (1, 1)
    assert m.solde_reel(etat, B1) == 9900.00
    prev(etat, date(2026, 10, 30), B1, "Ancienne prévision", 999.00)
    assert m.solde_previsionnel(etat, B1, 30) == 9900.00


def test_prevision_echue_reste_due_et_a_controler():
    etat = nouvel_etat()
    p = prev(etat, date(2026, 11, 10), B1, "Facture client X", 2000.00, type_=m.TYPE_RECETTE)
    assert m.statut_suivi(etat.params, p) == m.PREV_A_CONTROLER
    assert m.solde_previsionnel(etat, B1, 0) == 12000.00


def test_rejet_d_une_proposition():
    etat = nouvel_etat()
    p = prev(etat, date(2026, 11, 5), B1, "Loyer bureau", 1500.00)
    importer_b1(etat, {"nov.csv": csv_b1([("06/11/2026", "PRLV SEPA SCI DU PARC", "-1 500,00")])})
    props = m.rechercher(etat)
    m.valider(etat, [(props[0].op.id, p.id, "Non")])
    assert etat.propositions == []
    assert [o.id for o in m.a_traiter(etat)] == ["OB-000001"]
    assert p.id_propose == "" and m.solde_previsionnel(etat, B1, 30) == 7000.00


def test_validation_conflictuelle_refusee():
    etat = nouvel_etat()
    p1 = prev(etat, date(2026, 11, 5), B1, "Carburant", 45.00)
    p2 = prev(etat, date(2026, 11, 6), B1, "Carburant", 45.00)
    importer_b1(etat, {"nov.csv": csv_b1([("05/11/2026", "CARTE TOTAL", "-45,00")])})
    props = m.rechercher(etat)
    with pytest.raises(m.ConflitValidation):
        m.valider(etat, [(props[0].op.id, p1.id, "Oui"), (props[0].op.id, p2.id, "Oui")])
    assert p1.statut == p2.statut == m.PREV_PREVU


def test_validation_automatique_des_certaines():
    etat = nouvel_etat(auto_valider=True)
    p = prev(etat, date(2026, 11, 5), B1, "Loyer bureau", 1500.00)
    p_amb1 = prev(etat, date(2026, 11, 10), B1, "Carburant", 45.00)
    p_amb2 = prev(etat, date(2026, 11, 11), B1, "Carburant", 45.00)
    importer_b1(etat, {"nov.csv": csv_b1([("06/11/2026", "PRLV LOYER BUREAU", "-1 500,00"),
                                          ("10/11/2026", "CARTE TOTAL", "-45,00")])})
    m.rechercher(etat)
    assert p.statut == m.PREV_POINTE and etat.op("OB-000001").mode == m.MODE_AUTO
    assert p_amb1.statut == p_amb2.statut == m.PREV_PREVU
    assert len(etat.propositions) == 2


def test_rapprochement_manuel_montant_different():
    etat = nouvel_etat()
    p = prev(etat, date(2026, 11, 5), B1, "Facture plombier", 1000.00)
    importer_b1(etat, {"nov.csv": csv_b1([("07/11/2026", "VIR SEPA PLOMBERIE DURAND", "-1 002,50")])})
    m.rechercher(etat)
    assert etat.propositions == []  # montant différent : pas de proposition automatique
    assert m.traiter(etat, "OB-000001", "Rapprocher avec la prévision indiquée", id_prevision=p.id) == "ok"
    assert p.statut == m.PREV_POINTE and p.montant_reel == -1002.50
    assert m.solde_previsionnel(etat, B1, 30) == 8997.50


def test_similarite():
    assert m.similarite("Loyer bureau", "PRLV SEPA SCI DU PARC LOYER BUREAU NOV") == 1.0
    assert m.similarite("Abonnement Orange", "PRLV SEPA ORANGE SA") == 1.0
    assert m.similarite("Électricité EDF", "PRLV EDF ELECTRICITE") == 1.0
    assert m.similarite("Assurance locaux", "PRLV AXA FRANCE IARD") < 0.4
    assert m.similarite("", "PRLV") == 0.0
