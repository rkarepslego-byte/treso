"""Charge le classeur .xlsm dans LibreOffice et relit le projet VBA importé.

Contrôle du fichier vbaProject.bin généré par une implémentation indépendante
(celle de LibreOffice) : chaque module doit être présent et son code identique aux
sources de src/vba. (LibreOffice ne signale pas de façon fiable les erreurs de
compilation VBA : la syntaxe est contrôlée par tests/outils/lint_vba.py.)

Usage : /usr/bin/python3 compiler_vba_lo.py <classeur.xlsm> <dossier src/vba>
Sortie : une ligne par module « OK » ou « ERREUR <message> ».
"""
import os
import re
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(nom, valeur):
    p = PropertyValue()
    p.Name, p.Value = nom, valeur
    return p


def normaliser(code):
    lignes = []
    for ligne in code.replace("\r\n", "\n").split("\n"):
        l = ligne.rstrip()
        if not l or l.startswith("Attribute ") or l.startswith("Rem Attribute") or l.startswith("Option VBASupport") \
                or l.startswith("Option Compatible"):
            continue
        lignes.append(l)
    return "\n".join(lignes)


def lire_sources(dossier):
    sources = {}
    for f in os.listdir(dossier):
        if f.endswith((".bas", ".cls")):
            with open(os.path.join(dossier, f), encoding="utf-8") as fh:
                sources[os.path.splitext(f)[0]] = fh.read()
    return sources


def main(chemin, dossier_sources):
    sources = lire_sources(dossier_sources)
    tuyau = "treso_vba_%d" % os.getpid()
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             "-env:UserInstallation=file:///tmp/lo_profil_treso_vba",
                             "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % tuyau],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    code_retour = 0
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(60):
            try:
                ctx = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % tuyau)
                break
            except Exception:
                time.sleep(0.5)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        # MacroExecutionMode 4 = toujours exécuter sans avertissement (nécessaire pour compiler)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(chemin)), "_blank", 0,
                                           (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        libs = doc.BasicLibraries
        vus = set()
        print("Bibliothèques :", list(libs.getElementNames()))
        for nom_lib in libs.getElementNames():
            libs.loadLibrary(nom_lib)
            lib = libs.getByName(nom_lib)
            for nom_mod in lib.getElementNames():
                source = lib.getByName(nom_mod)
                nb = len(source.splitlines())
                vus.add(nom_mod)
                attendu = sources.get(nom_mod)
                if attendu is None:
                    print("OK      %s.%s (module de feuille vide)" % (nom_lib, nom_mod) if nb <= 4
                          else "ERREUR  %s.%s : module inattendu" % (nom_lib, nom_mod))
                    continue
                if normaliser(source) == normaliser(attendu):
                    print("OK      %s.%s (%d lignes identiques à la source)" % (nom_lib, nom_mod, nb))
                else:
                    print("ERREUR  %s.%s : code différent de la source" % (nom_lib, nom_mod))
                    code_retour = 1
        manquants = set(sources) - vus
        for nom_mod in sorted(manquants):
            print("ERREUR  module absent : %s" % nom_mod)
            code_retour = 1
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
    sys.exit(code_retour)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
