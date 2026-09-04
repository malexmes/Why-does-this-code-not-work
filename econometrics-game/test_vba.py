import uno, os, sys
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
    try: c.close(True)
    except Exception: pass
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath('app_v3.xlsx')), "_blank", 0, (pv("Hidden", True), pv("MacroExecutionMode", 4)))
libs = doc.BasicLibraries; libs.VBACompatibilityMode = True; libs.ProjectName = 'VBAProject'
if not libs.hasByName('VBAProject'): libs.createLibrary('VBAProject')
lib = libs.getByName('VBAProject')
def add(name, code, mtype=1, obj=None):
    if lib.hasByName(name): lib.removeByName(name)
    lib.insertByName(name, code)
    mi = uno.createUnoStruct('com.sun.star.script.ModuleInfo'); mi.ModuleType = mtype
    if obj is not None: mi.ModuleObject = obj
    lib.insertModuleInfo(name, mi)
add('Game', 'Option VBASupport 1\n' + open('Game.bas').read())
sp = doc.getScriptProvider()
def call(name, *args):
    return sp.getScript(f"vnd.sun.star.script:VBAProject.Game.{name}?language=Basic&location=document").invoke(tuple(args), (), ())
st = doc.Sheets.getByName('State')
B = lambda r: st.getCellByPosition(1, r-1).Value
click = lambda sheet, addr: call('Dispatch', sheet, addr)
fails = []
def check(label, got, exp):
    ok = abs(got - exp) < 1e-6
    print(('PASS' if ok else 'FAIL'), label, got, 'expected', exp)
    if not ok: fails.append(label)
call('ResetGame'); check('reset step', B(1), 1); check('reset stage', B(30), 1)
for _ in range(3): click('Play', 'T29')
check('3x next', B(1), 4)
click('Play', 'Y16'); click('Play', 'Y16'); click('Play', 'X16'); click('Play', 'V16'); check('base after ++ ++ + -', B(2), 1000)
click('Play', 'V16'); click('Play', 'V16'); click('Play', 'U16'); click('Play', 'U16'); click('Play', 'U16'); check('base floors at 0', B(2), 0)
click('Play', 'B29'); check('back', B(1), 3); click('Play', 'B29'); click('Play', 'T11'); check('quiz1 A', B(17), 1); click('Play', 'T13'); check('quiz1 B', B(17), 2)
click('Play', 'Y16'); check('nudge ignored on quiz step', B(2), 0)
for _ in range(9): click('Play', 'T29')
check('at step 11', B(1), 11)
for _ in range(12): click('Play', 'X16')
check('memory capped 90', B(6), 90)
for _ in range(5): click('Play', 'T29')
check('at step 16', B(1), 16)
click('Play', 'V11'); click('Play', 'V11'); click('Play', 'V12'); click('Play', 'U12'); click('Play', 'U12'); check('TV', B(8), 2); check('video floor', B(9), 0)
click('Play', 'G29'); click('Play', 'L29'); check('hint on', B(15), 1); check('reveal on', B(14), 1)
click('Play', 'T29'); check('hint off after next', B(15), 0); check('reveal off after next', B(14), 0)
for _ in range(10): click('Play', 'T29')
check('step capped 21', B(1), 21)
click('Play', 'T29'); check('to budget stage', B(30), 2)
print('active sheet:', doc.CurrentController.ActiveSheet.Name)
click('Budget', 'N11'); click('Budget', 'N11'); click('Budget', 'M16'); check('DRTV +2', B(51), 985 + 2 * 98); check('posters -1', B(56), 521 - 52)
for _ in range(20): click('Budget', 'N11')
check('DRTV cap', B(51), 1478)
for _ in range(20): click('Budget', 'M23')
check('door drops floor', B(63), 11)
click('Budget', 'G29'); check('AP mode', B(64), 1); click('Budget', 'N12'); check('mode back to mine on nudge', B(64), 0)
click('Budget', 'L29'); check('reset DRTV', B(51), 985); check('reset door drops', B(63), 22)
click('Budget', 'T21'); check('quiz8 B', B(24), 2); click('Budget', 'T20'); check('quiz8 A', B(24), 1)
click('Budget', 'T29'); check('to long game', B(30), 3); print('active sheet:', doc.CurrentController.ActiveSheet.Name)
click('Long game', 'T22'); check('quiz9 C', B(25), 3); click('Long game', 'T20'); check('quiz9 A', B(25), 1)
click('Long game', 'B29'); check('back to budget', B(30), 2); click('Budget', 'T29'); click('Long game', 'T29'); check('to giving', B(30), 4)
click('Giving', 'T20'); check('quiz10 A', B(26), 1); click('Giving', 'T29'); check('to finish', B(30), 5); print('active sheet:', doc.CurrentController.ActiveSheet.Name)
click('Finish', 'B29'); check('finish back', B(30), 4); click('Giving', 'T29')
click('Finish', 'T29'); check('start again step', B(1), 1); check('start again stage', B(30), 1); check('start again quiz', B(24), 0); check('start again memory', B(6), 0)
call('OpenGame'); check('open game stage', B(30), 1)
doc.close(True)
print('RESULT:', 'ALL PASS' if not fails else f'FAILED {fails}')
