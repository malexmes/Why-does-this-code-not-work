"""Drive every game action through the VBA inside LibreOffice (Excel-compatible mode).

Needs a headless soffice listening on port 2002 and app_v3.xlsx in the working directory.
"""
import uno, os
from com.sun.star.beans import PropertyValue


def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
comps = desktop.getComponents().createEnumeration()
while comps.hasMoreElements():
    c = comps.nextElement()
    try:
        c.close(True)
    except Exception:
        pass
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath('app_v3.xlsx')), "_blank", 0, (pv("Hidden", True), pv("MacroExecutionMode", 4)))
libs = doc.BasicLibraries; libs.VBACompatibilityMode = True; libs.ProjectName = 'VBAProject'
if not libs.hasByName('VBAProject'):
    libs.createLibrary('VBAProject')
lib = libs.getByName('VBAProject')
if lib.hasByName('Game'):
    lib.removeByName('Game')
lib.insertByName('Game', 'Option VBASupport 1\n' + open('Game.bas').read())
mi = uno.createUnoStruct('com.sun.star.script.ModuleInfo'); mi.ModuleType = 1
lib.insertModuleInfo('Game', mi)
sp = doc.getScriptProvider()


def call(name, *args):
    return sp.getScript(f"vnd.sun.star.script:VBAProject.Game.{name}?language=Basic&location=document").invoke(tuple(args), (), ())


st = doc.Sheets.getByName('State')
B = lambda r: st.getCellByPosition(1, r - 1).Value
click = lambda sheet, addr: call('Dispatch', sheet, addr)
active = lambda: doc.CurrentController.ActiveSheet.Name
fails = []


def check(label, got, exp):
    ok = (got == exp) if isinstance(exp, str) else abs(got - exp) < 1e-6
    print(('PASS' if ok else 'FAIL'), label, got, 'expected', exp)
    if not ok:
        fails.append(label)


RFL = 'Race for Life'
call('ResetGame'); check('reset step', B(1), 1); check('reset stage', B(30), 1); check('reset sheet', active(), RFL)
for _ in range(3):
    click(RFL, 'T29')
check('3x next', B(1), 4)
click(RFL, 'Y16'); click(RFL, 'Y16'); click(RFL, 'X16'); click(RFL, 'V16'); check('base after ++ ++ + -', B(2), 1000)
for _ in range(5):
    click(RFL, 'U16')
check('base floors at 0', B(2), 0)
click(RFL, 'B29'); check('back', B(1), 3); click(RFL, 'B29')
click(RFL, 'T11'); check('quiz1 A', B(17), 1); click(RFL, 'T13'); check('quiz1 B', B(17), 2)
click(RFL, 'Y16'); check('nudge ignored on quiz step', B(2), 0)
for _ in range(9):
    click(RFL, 'T29')
check('at step 11', B(1), 11)
for _ in range(12):
    click(RFL, 'X16')
check('memory capped 90', B(6), 90)
for _ in range(5):
    click(RFL, 'T29')
check('at step 16', B(1), 16)
click(RFL, 'V11'); click(RFL, 'V11'); click(RFL, 'V12'); click(RFL, 'U12'); click(RFL, 'U12'); check('TV', B(8), 2); check('video floor', B(9), 0)
click(RFL, 'N38'); check('budget nudge ignored before step 22', B(51), 985)
click(RFL, 'G29'); click(RFL, 'L29'); check('hint on', B(15), 1); check('reveal on', B(14), 1)
click(RFL, 'T29'); check('hint off after next', B(15), 0); check('reveal off after next', B(14), 0)
for _ in range(5):
    click(RFL, 'T29')
check('at step 22 (budget)', B(1), 22); check('still on RFL tab', active(), RFL)
click(RFL, 'N38'); click(RFL, 'N38'); click(RFL, 'M43'); check('DRTV +2', B(51), 985 + 2 * 98); check('posters -1', B(56), 521 - 52)
for _ in range(20):
    click(RFL, 'N38')
check('DRTV cap', B(51), 1478)
for _ in range(20):
    click(RFL, 'M50')
check('door drops floor', B(63), 11)
click(RFL, 'G55'); check('AP mode', B(64), 1); click(RFL, 'N39'); check('mode back to mine on nudge', B(64), 0)
click(RFL, 'L55'); check('reset DRTV', B(51), 985); check('reset door drops', B(63), 22)
click(RFL, 'T48'); check('quiz8 B', B(24), 2); click(RFL, 'T47'); check('quiz8 A', B(24), 1)
click(RFL, 'B55'); check('budget back to step 21', B(1), 21); click(RFL, 'T29'); check('forward to 22 again', B(1), 22)
click(RFL, 'T55'); check('to Brand', B(30), 2); check('Brand sheet', active(), 'Brand')
click('Brand', 'T22'); check('quiz9 C', B(25), 3); click('Brand', 'T20'); check('quiz9 A', B(25), 1)
click('Brand', 'B29'); check('back to RFL', B(30), 1); check('RFL sheet', active(), RFL); check('RFL still at 22', B(1), 22)
click(RFL, 'T29'); check('RFL next at 22 goes to Brand', B(30), 2)
click('Brand', 'T29'); check('to Giving', B(30), 3); check('Giving sheet', active(), 'Giving')
click('Giving', 'T20'); check('quiz10 A', B(26), 1)
click('Giving', 'B29'); check('Giving back to Brand', B(30), 2); click('Brand', 'T29')
click('Giving', 'T29'); check('to Finish', B(30), 4); check('Finish sheet', active(), 'Finish')
click('Finish', 'B29'); check('finish back to Giving', B(30), 3); click('Giving', 'T29')
click('Finish', 'T29'); check('start again step', B(1), 1); check('start again stage', B(30), 1); check('start again quiz', B(24), 0); check('start again memory', B(6), 0)
call('OpenGame'); check('open game stage', B(30), 1); check('open game sheet', active(), RFL)
doc.close(True)
print('RESULT:', 'ALL PASS' if not fails else f'FAILED {fails}')
