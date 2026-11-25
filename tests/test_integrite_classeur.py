"""Contrôles d'intégrité du classeur livré (structure OpenXML, projet VBA, Power Query)."""

import io
import os
import re
import shutil
import subprocess
import sys
import zipfile

import olefile
import openpyxl
import pytest

from build import datamashup as dm
from build import schema as S
from build.generer_classeur import construire
from tests.outils import lint_vba

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARSER_M = os.environ.get("PQ_PARSER", "/tmp/claude-0/-home-user-treso/a0800f07-5658-5a20-8df8-ee1197784284/"
                          "scratchpad/pqparser/node_modules/@microsoft/powerquery-parser")


@pytest.fixture(scope="module")
def classeur(tmp_path_factory):
    dossier = tmp_path_factory.mktemp("livraison")
    return construire(str(dossier))


def test_code_vba_sans_anomalie(classeur):
    noms = set(openpyxl.load_workbook(classeur).defined_names.keys())
    erreurs, procedures = lint_vba.verifier(lint_vba.lire_modules(), noms)
    erreurs = [e for e in erreurs if "« Dossier_ »" not in e]  # nom construit dynamiquement (Dossier_B1 / B2)
    for macro in sorted(lint_vba.macros_boutons()):
        info = procedures.get(macro.lower())
        if info is None or not info[1] or info[2] != 0:
            erreurs.append(f"macro de bouton {macro} introuvable, privée ou avec paramètres")
    assert erreurs == []


def test_projet_vba_relu_par_olefile(classeur):
    with zipfile.ZipFile(classeur) as z:
        binaire = z.read("xl/vbaProject.bin")
    ole = olefile.OleFileIO(io.BytesIO(binaire))
    flux = {"/".join(e) for e in ole.listdir()}
    attendus = {"PROJECT", "PROJECTwm", "VBA/dir", "VBA/_VBA_PROJECT", "VBA/ThisWorkbook"}
    attendus |= {"VBA/" + os.path.splitext(f)[0] for f in os.listdir(os.path.join(RACINE, "src", "vba")) if f.endswith(".bas")}
    attendus |= {"VBA/" + code for _n, code, _v in S.ONGLETS}
    assert attendus <= flux
    from oletools.olevba import VBA_Parser
    vba = VBA_Parser("x.xlsm", data=open(classeur, "rb").read())
    modules = {os.path.splitext(nom)[0]: code for _f, _s, nom, code in vba.extract_macros()}
    for f in os.listdir(os.path.join(RACINE, "src", "vba")):
        if f.endswith(".bas"):
            nom = os.path.splitext(f)[0]
            source = open(os.path.join(RACINE, "src", "vba", f), encoding="utf-8").read()
            corps = [l for l in source.splitlines() if not l.startswith("Attribute ")]
            relu = [l for l in modules[nom].replace("\r\n", "\n").split("\n") if not l.startswith("Attribute ")]
            assert [l.rstrip() for l in corps if l.strip()] == [l.rstrip() for l in relu if l.strip()], nom


def test_parties_openxml_coherentes(classeur):
    with zipfile.ZipFile(classeur) as z:
        noms = set(z.namelist())
        ct = z.read("[Content_Types].xml").decode()
        for partie in re.findall(r'PartName="/([^"]+)"', ct):
            assert partie in noms, partie
        for rels in [n for n in noms if n.endswith(".rels")]:
            base = os.path.dirname(os.path.dirname(rels))
            for cible, mode in re.findall(r'Target="([^"]+)"(?: TargetMode="(\w+)")?', z.read(rels).decode()):
                if mode == "External":
                    continue
                chemin = os.path.normpath(os.path.join(base, cible)).replace("\\", "/") if not cible.startswith("/") else cible[1:]
                assert chemin in noms, (rels, cible)
        wb = z.read("xl/workbook.xml").decode()
        assert 'name="ExternalData_1"' in wb and 'name="ExternalData_2"' in wb
        for i in (1, 2):
            qt = z.read(f"xl/queryTables/queryTable{i}.xml").decode()
            assert f'connectionId="{i}"' in qt and qt.count("<queryTableField ") == len(S.COLONNES_TRANSIT)
        assert z.read("xl/connections.xml").decode().count("Microsoft.Mashup.OleDb.1") == 2
        dessins = "".join(z.read(n).decode() for n in noms if n.startswith("xl/drawings/drawing"))
        assert 'macro="[0]!Actualiser"' in dessins and 'name="btn_Actualiser"' in dessins


def test_requetes_power_query_integrees(classeur):
    with zipfile.ZipFile(classeur) as z:
        section = dm.lire_section_m(z.read("customXml/item1.xml"))
    for nom in ("Parametres_Import", "fnNormaliserReleve", "fnImporterBanque", "Import_Banque1", "Import_Banque2"):
        assert f"shared {nom} = " in section
    if not (shutil.which("node") and os.path.isdir(PARSER_M)):
        pytest.skip("analyseur M indisponible")
    chemin = os.path.join(os.path.dirname(classeur), "Section1.m")
    with open(chemin, "w", encoding="utf-8") as f:
        f.write(section)
    sources = [os.path.join(RACINE, "src", "powerquery", f) for f in os.listdir(os.path.join(RACINE, "src", "powerquery"))]
    r = subprocess.run(["node", os.path.join(RACINE, "tests", "outils", "verifier_m.js"), PARSER_M, chemin, *sources],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr


def test_releves_de_demonstration_et_import_python(classeur):
    """Les relevés livrés se relisent avec les paramètres par défaut (miroir de Power Query)."""
    from build import donnees_demo as dd
    from reference import moteur as m
    from reference.formats import FORMAT_B1, FORMAT_B2, vers_parametres_moteur
    racine = os.path.dirname(classeur)
    p1 = vers_parametres_moteur(FORMAT_B1, dd.NOM_B1)
    p2 = vers_parametres_moteur(FORMAT_B2, dd.NOM_B2)
    f1, f2 = dd.lignes_brutes_depuis_fichiers(racine, p1, p2)
    t1 = m.preparer_import(f1, p1)
    t2 = m.preparer_import(f2, p2)
    assert not [l for l in t1 + t2 if l.erreur]
    assert len(t1) >= 10 and len(t2) >= 8
    # les deux péages identiques du même jour sont deux lignes distinctes
    peages = [l for l in t2 if "PEAGE" in l.libelle_norm]
    assert len(peages) == 2 and len({l.cle for l in peages}) == 2


def test_relecture_vba_par_libreoffice(classeur):
    if not (shutil.which("soffice") and os.path.exists("/usr/bin/python3")):
        pytest.skip("LibreOffice indisponible")
    r = subprocess.run(["/usr/bin/python3", os.path.join(RACINE, "tests", "outils", "relire_vba_lo.py"), classeur,
                        os.path.join(RACINE, "src", "vba")], capture_output=True, text=True, timeout=180)
    assert r.returncode == 0 and "ERREUR" not in r.stdout, r.stdout + r.stderr
    nb_sources = len([f for f in os.listdir(os.path.join(RACINE, "src", "vba")) if f.endswith((".bas", ".cls"))])
    assert r.stdout.count("identiques à la source") == nb_sources
