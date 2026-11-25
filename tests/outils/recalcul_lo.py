"""Recalcule un classeur avec LibreOffice (UNO) et l'enregistre au format .xlsx.

Usage : /usr/bin/python3 recalcul_lo.py <entree.xlsx|xlsm> <sortie.xlsx>
(à lancer avec l'interpréteur système qui fournit le module « uno »)
"""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(nom, valeur):
    p = PropertyValue()
    p.Name, p.Value = nom, valeur
    return p


def main(entree, sortie):
    profil = "file:///tmp/lo_profil_treso"
    tuyau = "treso_recalc_%d" % os.getpid()
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             "-env:UserInstallation=" + profil,
                             "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % tuyau],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        url = uno.systemPathToFileUrl(os.path.abspath(entree))
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True), prop("MacroExecutionMode", 0)))
        doc.calculateAll()
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(sortie)), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
