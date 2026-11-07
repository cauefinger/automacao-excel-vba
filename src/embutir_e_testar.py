r"""Embute PrazosModulo.bas no .ods e roda um teste headless.
Execute com o python do LibreOffice:  program\python.exe src\embutir_e_testar.py
"""
import os, sys, time, subprocess, uno
from com.sun.star.beans import PropertyValue
base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
soffice = r"C:\Program Files\LibreOffice\program\soffice.exe"
p = subprocess.Popen([soffice, "-env:UserInstallation=file:///" + os.environ["TEMP"].replace("\\","/") + "/lo_prof_claude", "--headless", "--norestore",
    '--accept=socket,host=localhost,port=2021;urp;'])
ctx = None
for _ in range(40):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = res.resolve("uno:socket,host=localhost,port=2021;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
def pv(n, v):
    x = PropertyValue(); x.Name = n; x.Value = v; return x
src = uno.systemPathToFileUrl(os.path.join(base, "controle_prazos.ods"))
doc = desktop.loadComponentFromURL(src, "_blank", 0, (pv("Hidden", True),))
code = open(os.path.join(base, "src", "PrazosModulo.bas"), encoding="utf-8").read()
libs = doc.BasicLibraries
if not libs.hasByName("Standard"): libs.createLibrary("Standard")
std = libs.getByName("Standard")
if std.hasByName("Module1"): std.replaceByName("Module1", code)
else: std.insertByName("Module1", code)
if "--salvar" in sys.argv:
    doc.storeToURL(uno.systemPathToFileUrl(os.path.join(base, "controle_prazos.ods")), (pv("FilterName","calc8"),))
    doc.close(True)
    try: desktop.terminate()
    except Exception: pass
    p.wait(timeout=20); sys.exit(0)
# teste: seleciona linhas e roda a macro
sheets = doc.Sheets
ctrl = doc.CurrentController
pessoas = sheets.getByName("Pessoas")
sp = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider(doc)
def run(name):
    sp.getScript("vnd.sun.star.script:Standard.Module1.%s?language=Basic&location=document" % name).invoke((), (), ())
ctrl.setActiveSheet(pessoas)
for r in (1, 2, 4, 1):          # 2 e (repetida) 2
    ctrl.select(pessoas.getCellByPosition(0, r))
    ctrl.setActiveSheet(pessoas)
    run("EnviarParaPrazo")
pr = sheets.getByName("Prazos")
for r in range(0, 6):
    print([pr.getCellByPosition(c, r).String for c in range(6)])
if "--salvar" in sys.argv:
    doc.storeToURL(uno.systemPathToFileUrl(os.path.join(base, "controle_prazos.ods")), (pv("FilterName","calc8"),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
p.wait(timeout=20)
