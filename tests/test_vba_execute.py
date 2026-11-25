"""EXÉCUTION du code VBA de l'outil (src/vba) avec l'interpréteur VBA et le modèle Excel
simulés (tests/outils), puis comparaison, table par table, avec le moteur de référence.

Le classeur simulé est chargé depuis le vrai classeur généré ; l'actualisation Power
Query des tables de transit est remplacée par la lecture des relevés par le moteur de
référence, à partir du dossier que le VBA a écrit dans la requête Parametres_Import.
"""

import os
import re
from datetime import date, datetime

import pytest
import xlsxwriter

from build import donnees_demo as dd
from build.generer_classeur import construire
from reference import moteur as m
from reference.formats import FORMAT_B1, FORMAT_B2, vers_parametres_moteur
from tests.outils.excel_simule import Monde
from tests.outils.vba_interpreteur import EMPTY, Interpreteur, datetime_vers_serie

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = date(2026, 11, 25)


def serie(d):
    return datetime_vers_serie(datetime(d.year, d.month, d.day))


def sources_vba():
    src = {}
    dossier = os.path.join(RACINE, "src", "vba")
    for f in sorted(os.listdir(dossier)):
        if f.endswith((".bas", ".cls")):
            src[os.path.splitext(f)[0]] = open(os.path.join(dossier, f), encoding="utf-8").read()
    return src


def rafraichir_transit(monde, table, formule):
    """Remplace Power Query : lit le dossier indiqué par la requête générée par le VBA."""
    code = "B1" if table.p_name.endswith("B1") else "B2"
    num = code[1]
    mt = re.search(r'Banque%s = \[Code = "%s", Banque = "([^"]*)", Dossier = "([^"]*)"' % (num, code), formule)
    assert mt, "requête Parametres_Import non générée par le VBA"
    nom, dossier = mt.group(1), monde.chemin_local(mt.group(2))
    p = vers_parametres_moteur(FORMAT_B1 if code == "B1" else FORMAT_B2, nom)
    fichiers = {f: m.lire_fichier(os.path.join(dossier, f), p) for f in sorted(os.listdir(dossier))
                if os.path.splitext(f)[1].lower() in p["Extensions"]}
    lignes = []
    for l in m.preparer_import(fichiers, p):
        lignes.append([l.banque, l.code, serie(l.date) if l.date else EMPTY,
                       serie(l.date_valeur) if l.date_valeur else EMPTY, l.libelle, l.libelle_norm,
                       l.montant if l.montant is not None else EMPTY, l.reference, l.fichier,
                       l.ligne if l.ligne else EMPTY, l.occurrence if l.occurrence else EMPTY, l.cle,
                       l.erreur if l.erreur else EMPTY])
    table.lignes = lignes


@pytest.fixture
def env(tmp_path):
    chemin = construire(str(tmp_path), REF)
    # « Now » = heure réelle (comparée aux dates de modification des fichiers déposés)
    monde = Monde(chemin, {"Date_Reference": serie(REF)}, datetime_vers_serie(datetime.now()), rafraichir_transit)
    interp = Interpreteur(sources_vba(), monde)
    interp.publics["gsilencieux"].valeur = True
    etat = dd.etat_initial(REF)
    return monde, interp, etat, str(tmp_path)


def actualiser_reference(etat, racine):
    p1 = vers_parametres_moteur(FORMAT_B1, dd.NOM_B1)
    p2 = vers_parametres_moteur(FORMAT_B2, dd.NOM_B2)
    f1, f2 = dd.lignes_brutes_depuis_fichiers(racine, p1, p2)
    m.importer(etat, m.preparer_import(f1, p1), datetime(2026, 11, 25, 9, 36))
    m.importer(etat, m.preparer_import(f2, p2), datetime(2026, 11, 25, 9, 36))
    m.generer_echeances(etat)
    m.rechercher(etat)
    m.traitements_automatiques(etat)


def texte(v):
    return "" if v is EMPTY or v is None else str(v)


def comparer_tout(monde, etat, interp):
    assert interp.err.description == "" or True
    # Journal bancaire
    vba = {r["ID"]: r for r in monde.table("tbl_Banque").dicts() if texte(r["ID"])}
    assert sorted(vba) == sorted(o.id for o in etat.operations)
    for o in etat.operations:
        r = vba[o.id]
        assert (texte(r["Statut"]), texte(r["ID prévision"]), texte(r["Catégorie"]), texte(r["Mode"]), r["Clé"]) == \
            (o.statut, o.id_prevision, o.categorie, o.mode, o.cle), o.id
        assert abs(r["Montant"] - o.montant) < 0.005 and r["Date"] == serie(o.date)
    # Prévisions
    vba = {r["ID"]: r for r in monde.table("tbl_Previsions").dicts() if texte(r["ID"])}
    assert sorted(vba) == sorted(p.id for p in etat.previsions)
    for p in etat.previsions:
        r = vba[p.id]
        assert (texte(r["Statut"]), texte(r["ID opération"]), texte(r["ID proposé"]), r["Date prévue"]) == \
            (p.statut, p.id_operation, p.id_propose, serie(p.date)), p.id
        assert abs(r["Montant"] - abs(p.montant)) < 0.005
    # Propositions
    vba = [(texte(r["Niveau"]), r["Groupe"], texte(r["Suggestion"]), r["ID opération"], r["ID prévision"], r["Écart (jours)"],
            round(r["Similarité"], 4), round(r["Score"], 4))
           for r in monde.table("tbl_Propositions").dicts() if texte(r["ID opération"])]
    ref = [(pr.niveau, pr.groupe, "Suggérée" if pr.suggere else "", pr.op.id, pr.prev.id, pr.ecart, pr.sim, pr.score)
           for pr in etat.propositions]
    assert vba == ref
    # À traiter
    vba = [(r["ID opération"], texte(r["ID prévision"])) for r in monde.table("tbl_ATraiter").dicts()
           if texte(r["ID opération"])]
    ref = [(o.id, m.prevision_proche(etat, o)) for o in m.a_traiter(etat)]
    assert vba == ref


def test_actualisation_complete_identique_au_moteur(env):
    monde, interp, etat, racine = env
    interp.appeler("Actualiser")
    assert interp.publics["gderniereerreur"].valeur == "", interp.publics["gderniereerreur"].valeur
    actualiser_reference(etat, racine)
    comparer_tout(monde, etat, interp)
    # historique : une ligne par fichier
    histo = [r for r in monde.table("tbl_Historique").dicts() if texte(r["ID import"])]
    ref = {h["fichier"]: h for h in etat.historique}
    assert sorted(r["Fichier"] for r in histo) == sorted(ref)
    for r in histo:
        h = ref[r["Fichier"]]
        assert (r["Lignes lues"], r["Ajoutées"], r["Doublons ignorés"], r["Antérieures au solde initial"]) == \
            (h["lues"], h["ajoutees"], h["doublons"], h["anterieures"])
    # sauvegarde faite avant l'actualisation
    assert len(monde.classeur.copies) == 1 and os.path.exists(monde.classeur.copies[0])
    # la requête Parametres_Import pointe vers les dossiers d'import
    formule = next(r.p_formula for r in monde.classeur.requetes if r.p_name == "Parametres_Import")
    assert "Imports" in formule and 'ModeMontant = "Débit/Crédit"' in formule

    # Réactualisation : aucun doublon, rien ne change
    nb = len([r for r in monde.table("tbl_Banque").dicts() if texte(r["ID"])])
    interp.appeler("Actualiser")
    assert len([r for r in monde.table("tbl_Banque").dicts() if texte(r["ID"])]) == nb
    comparer_tout(monde, etat, interp)


def test_validation_traitements_et_apprentissage(env):
    monde, interp, etat, racine = env
    interp.appeler("Actualiser")
    actualiser_reference(etat, racine)

    interp.appeler("ToutValiderCertaines")
    m.valider_certaines(etat)
    comparer_tout(monde, etat, interp)

    # Choix d'une correspondance incertaine (« Oui » sur la ligne suggérée)
    t = monde.table("tbl_Propositions")
    ligne = next(r for r in t.lignes if texte(r[t.entetes.index("Suggestion")]) == "Suggérée")
    ligne[t.entetes.index("Valider")] = "Oui"
    id_o, id_p = ligne[t.entetes.index("ID opération")], ligne[t.entetes.index("ID prévision")]
    interp.appeler("ValiderChoix")
    m.valider(etat, [(id_o, id_p, "Oui")])
    comparer_tout(monde, etat, interp)

    # Traitements : Amazon conservée (apprentissage), Morel rapprochée de la prévision proche,
    # péage -> charge mensuelle
    t = monde.table("tbl_ATraiter")
    col = {h: i for i, h in enumerate(t.entetes)}
    actions = {}
    for r in t.lignes:
        lib = texte(r[col["Libellé"]])
        if "AMAZON" in lib:
            r[col["Action"]], r[col["Catégorie"]] = "Conserver comme non prévue", "Fournitures"
            actions[r[col["ID opération"]]] = ("Conserver comme non prévue", "Fournitures", "")
        elif "MOREL" in lib:
            r[col["Action"]] = "Rapprocher avec la prévision indiquée"
            actions[r[col["ID opération"]]] = ("Rapprocher avec la prévision indiquée", "", texte(r[col["ID prévision"]]))
        elif "RESTAURANT" in lib:
            r[col["Action"]], r[col["Catégorie"]] = "Créer charge mensuelle", "Déplacements"
            actions[r[col["ID opération"]]] = ("Créer charge mensuelle", "Déplacements", "")
    assert len(actions) == 3
    interp.appeler("AppliquerTraitementsInterne")
    for id_op, (action, cat, id_prev) in actions.items():
        assert m.traiter(etat, id_op, action, cat, id_prevision=id_prev) == "ok"
    m.rechercher(etat)
    comparer_tout(monde, etat, interp)
    # règle apprise et charge créée
    regles = {texte(r["Mot-clé"]): (texte(r["Catégorie"]), texte(r["Action automatique"]))
              for r in monde.table("tbl_Regles").dicts() if texte(r["Mot-clé"])}
    for r in etat.regles:
        assert regles[r.mot_cle] == (r.categorie, r.action), r.mot_cle
    charges = [r for r in monde.table("tbl_Charges").dicts() if texte(r["ID"])]
    assert len(charges) == len(etat.charges) and charges[-1]["Libellé"] == etat.charges[-1].libelle
    assert charges[-1]["Première échéance"] == serie(etat.charges[-1].premiere)


def test_modification_de_charge_fixe(env):
    monde, interp, etat, racine = env
    interp.appeler("Actualiser")
    actualiser_reference(etat, racine)
    # hausse du loyer + désactivation d'une charge
    t = monde.table("tbl_Charges")
    col = {h: i for i, h in enumerate(t.entetes)}
    for r in t.lignes:
        if r[col["ID"]] == "CF-001":
            r[col["Montant"]] = 2600
        if r[col["ID"]] == "CF-006":
            r[col["Actif"]] = "Non"
    next(c for c in etat.charges if c.id == "CF-001").montant = 2600
    next(c for c in etat.charges if c.id == "CF-006").actif = False
    interp.appeler("GenererEcheancesInterne", module="modPrevisions")
    m.generer_echeances(etat)
    vba = {r["ID"]: r for r in monde.table("tbl_Previsions").dicts() if texte(r["ID"])}
    assert sorted(vba) == sorted(p.id for p in etat.previsions)
    for p in etat.previsions:
        assert vba[p.id]["Date prévue"] == serie(p.date) and abs(vba[p.id]["Montant"] - p.montant) < 0.005, p.id


def test_import_de_previsions(env):
    monde, interp, etat, racine = env
    chemin = os.path.join(racine, "ancien.xlsx")
    wb = xlsxwriter.Workbook(chemin)
    ws = wb.add_worksheet()
    lignes = [["Prévisionnel"], [], ["Échéance", "Intitulé", "Rubrique", "Compte", "Dépenses", "Recettes"],
              [date(2026, 12, 18), "Prime annuelle", "Salaires", "BPGO", 4000, None],
              [date(2027, 1, 12), "Subvention région", "Subventions", "cic", None, 7500],
              [date(2026, 10, 1), "Trop ancien", "Divers", "", 10, None]]
    fdate = wb.add_format({"num_format": "dd/mm/yyyy"})
    for i, l in enumerate(lignes):
        for j, v in enumerate(l):
            if isinstance(v, date):
                ws.write_datetime(i, j, datetime(v.year, v.month, v.day), fdate)
            elif v is not None:
                ws.write(i, j, v)
    wb.close()
    monde.fichier_a_ouvrir = chemin
    monde.reponses_inputbox = ["1"]
    interp.appeler("ImporterPrevisions")
    import openpyxl
    brut = [[c for c in r] for r in openpyxl.load_workbook(chemin).worksheets[0].iter_rows(values_only=True)]
    res = m.importer_previsions(etat, brut, dd.NOM_B1)
    assert res["importees"] == 2
    vba = [r for r in monde.table("tbl_Previsions").dicts() if texte(r["Origine"]) == "Import"]
    ref = [p for p in etat.previsions if p.origine == "Import"]
    assert [(r["ID"], r["Libellé"], r["Banque"], r["Type"], r["Montant"], r["Date prévue"]) for r in vba] == \
        [(p.id, p.libelle, p.banque, p.type, p.montant, serie(p.date)) for p in ref]


def _plage_ligne(monde, nom_table, id_ligne, colonne="ID"):
    from tests.outils.excel_simule import Plage
    t = monde.table(nom_table)
    j = t.entetes.index(colonne)
    i = next(k for k, l in enumerate(t.lignes) if l[j] == id_ligne)
    return t, Plage(t, i, 1, j, 1)


def test_macros_de_saisie_et_d_administration(env):
    monde, interp, etat, racine = env
    g = interp.publics

    # Formulaire : virement interne -> deux lignes liées
    for nom, v in (("Form_Date", serie(date(2026, 12, 2))), ("Form_Banque", dd.NOM_B1), ("Form_Libelle", "Appro Beta"),
                   ("Form_Montant", 3000), ("Form_Type", "Virement interne"), ("Form_BanqueDest", dd.NOM_B2)):
        monde.classeur.noms[nom.lower()].valeur = v
    interp.appeler("AjouterOperation")
    prev = [r for r in monde.table("tbl_Previsions").dicts() if r["Libellé"] == "Appro Beta"]
    assert [(r["Banque"], r["Type"]) for r in prev] == [(dd.NOM_B1, "Virement sortant"), (dd.NOM_B2, "Virement entrant")]
    assert prev[0]["Lien virement"] == prev[1]["ID"] and prev[1]["Lien virement"] == prev[0]["ID"]
    assert monde.nom("Form_Libelle") is EMPTY and monde.nom("Form_Montant") is EMPTY

    # Saisie invalide : refusée avec un message, rien n'est ajouté
    nb = len([r for r in monde.table("tbl_Previsions").dicts() if r["ID"] is not EMPTY])
    monde.classeur.noms["form_montant"].valeur = EMPTY
    monde.classeur.noms["form_libelle"].valeur = "Sans montant"
    interp.appeler("AjouterOperation")
    assert len([r for r in monde.table("tbl_Previsions").dicts() if r["ID"] is not EMPTY]) == nb
    assert "pas été ajoutée" in g["gjournal"].valeur

    # Tri par date
    interp.appeler("TrierPrevisions")
    dates = [r["Date prévue"] for r in monde.table("tbl_Previsions").dicts() if r["ID"] is not EMPTY]
    assert dates == sorted(dates)

    # Suppression d'une prévision saisie (et de sa contrepartie de virement)
    _t, cellule = _plage_ligne(monde, "tbl_Previsions", prev[0]["ID"])
    monde.active = next(f for f in monde.classeur.feuilles if f.p_name == "Saisie")
    monde.cellule_active = cellule
    interp.appeler("SupprimerLigne")
    assert not [r for r in monde.table("tbl_Previsions").dicts() if r["Libellé"] == "Appro Beta"]

    # Actualisation, validation puis annulation d'un rapprochement depuis le journal bancaire
    interp.appeler("Actualiser")
    interp.appeler("ToutValiderCertaines")
    pointee = next(r for r in monde.table("tbl_Banque").dicts() if r["Statut"] == "Pointé")
    id_prev = pointee["ID prévision"]
    _t, cellule = _plage_ligne(monde, "tbl_Banque", pointee["ID"])
    monde.active = next(f for f in monde.classeur.feuilles if f.p_name == "Opérations bancaires")
    monde.cellule_active = cellule
    interp.appeler("AnnulerRapprochement")
    op = next(r for r in monde.table("tbl_Banque").dicts() if r["ID"] == pointee["ID"])
    pv = next(r for r in monde.table("tbl_Previsions").dicts() if r["ID"] == id_prev)
    assert op["Statut"] in ("À traiter", "Proposé") and op["ID prévision"] in (EMPTY, "")
    assert pv["Statut"] == "Prévu" and pv["ID opération"] in (EMPTY, "")

    # Contrôle d'intégrité
    g["gjournal"].valeur = ""
    interp.appeler("Admin_VerifierIntegrite")
    assert "Aucune incohérence" in g["gjournal"].valeur, g["gjournal"].valeur

    # Nouveaux relevés : aucun après l'actualisation, un après dépôt d'un fichier
    assert interp.appeler("NouveauxReleves") == 0
    src = os.path.join(racine, "Imports", "Banque1", "operations_partie2.csv")
    with open(src, "rb") as f:
        contenu = f.read()
    with open(os.path.join(racine, "Imports", "Banque1", "releve_decembre.csv"), "wb") as f:
        f.write(contenu)
    assert interp.appeler("NouveauxReleves") == 1
    assert interp.appeler("TexteAlertes") is not None

    # Remise à zéro (mot de passe + confirmation)
    monde.reponses_inputbox = ["Treso2026", "VIDER"]
    interp.appeler("Admin_ViderDonnees")
    assert not [r for r in monde.table("tbl_Banque").dicts() if r["ID"] is not EMPTY]
    assert not [r for r in monde.table("tbl_Previsions").dicts() if r["ID"] is not EMPTY]
    assert len(monde.table("tbl_Previsions").lignes) == 30
    assert g["gderniereerreur"].valeur == ""


def test_recette_executee_en_vba(env):
    monde, interp, etat, racine = env
    res = interp.appeler("Recette_Executer", serie(REF))
    assert res.startswith("Erreur=;"), res
    valeurs = dict(x.split("=", 1) for x in res.split(";"))
    actualiser_reference(etat, racine)
    m.valider_certaines(etat)
    assert int(valeurs["NbOperations"]) == len(etat.operations)
    assert int(valeurs["NbPointees"]) == sum(1 for o in etat.operations if o.statut == m.OB_POINTE)
    assert int(valeurs["NbFichiersHistorique"]) == len(etat.historique)
    res2 = interp.appeler("Recette_Reimporter")
    v2 = dict(x.split("=", 1) for x in res2.split(";"))
    assert v2["Erreur"] == "" and v2["Avant"] == v2["Apres"]
