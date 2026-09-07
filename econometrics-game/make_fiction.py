"""Build the public, fictional edition of the econometrics game.

Everything is invented: the charity (Lantern), the event (Big Splash), the spend,
the sign-ups and every quoted figure. The weekly series are generated here from
a known model plus noise, so the "study answer" is the generating model itself.

Outputs (in web/): fiction_data.json, fiction_template.html, fiction_index.html
"""
import json, math, random, re, os

HERE = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(HERE, 'web')
random.seed(20260907)

real = json.load(open(os.path.join(HERE, 'game_data.json')))
dates = real['dates']
N = len(dates)
yrs = [int(d[:4]) for d in dates]
mons = [int(d[5:7]) for d in dates]
POUND = 80

# ---------------------------------------------------------------- shapes
def norm_mean1(a):
    m = sum(a) / len(a)
    return [v / m for v in a]

def lognorm(sigma):
    return math.exp(random.gauss(0, sigma))

# base: consideration plus emails; smooth season plus a few email blasts
shape_base = []
for i in range(N):
    season = 0.85 + 0.35 * max(0.0, math.cos((mons[i] - 3) / 12 * 2 * math.pi))
    trend = 1 + 0.03 * (yrs[i] - 2023)
    shape_base.append(season * trend * lognorm(0.05))
for i in range(N):  # two email blasts a year
    if (mons[i] == 1 and dates[i][8:10] < '15') or (mons[i] == 4 and '08' <= dates[i][8:10] < '15'):
        shape_base[i] *= 3.2
shape_base = norm_mean1(shape_base)

# brake: shut window in the autumn, plus a dearer ticket in 2024
shape_brake = []
for i in range(N):
    v = 0.15
    if mons[i] >= 8:
        v = 0.9 + 0.5 * min(1, (mons[i] - 7) / 3)
    if yrs[i] == 2024 and mons[i] <= 7:
        v += 0.3
    if yrs[i] == 2025 and mons[i] <= 7:
        v -= 0.05
    shape_brake.append(max(0, v) * lognorm(0.05))

# sales: discount windows each spring, with a pull-forward dip after the deepest one
shape_sale = [0.0] * N
windows = {  # (year, month, day range) -> depth
    2023: [((1, '01', '25'), 0.75), ((1, '26', '31'), 1.9), ((2, '01', '07'), 1.9), ((3, '20', '31'), 0.45), ((4, '01', '25'), 0.75), ((4, '26', '31'), 1.7)],
    2024: [((1, '01', '31'), 0.75), ((2, '01', '07'), 3.4), ((3, '15', '31'), 0.75), ((4, '01', '25'), 0.75), ((4, '26', '31'), 1.7)],
    2025: [((1, '01', '31'), 0.75), ((2, '01', '07'), 3.4), ((3, '20', '31'), 0.75), ((4, '01', '10'), 1.7)],
}
for i in range(N):
    y, m, dd = yrs[i], mons[i], dates[i][8:10]
    for (wm, d0, d1), depth in windows.get(y, []):
        if m == wm and d0 <= dd <= d1:
            shape_sale[i] = depth
    if y in (2024, 2025) and m == 12 and dd >= '25':
        shape_sale[i] = 0.75
for i in range(N):
    if shape_sale[i] >= 3 and i + 3 < N:
        for k in (1, 2, 3):
            if shape_sale[i + k] == 0:
                shape_sale[i + k] = -0.12

# ---------------------------------------------------------------- spend, £k a week
profile = {1: 1.0, 2: 1.25, 3: 1.35, 4: 1.2, 5: 0.9, 6: 0.5, 7: 0.2, 8: 0.05, 9: 0.05, 10: 0.12, 11: 0.25, 12: 0.55}
year_factor = {2023: 0.9, 2024: 1.12, 2025: 1.0}
CH_NAMES = ['TV', 'Online video', 'Radio', 'Posters', 'Digital, social and the rest']
CH_SHARE = [0.30, 0.16, 0.14, 0.09, 0.31]
ch_spend = [[0.0] * N for _ in CH_NAMES]
for i in range(N):
    tot = 150 * profile[mons[i]] * year_factor[yrs[i]] * lognorm(0.35)
    for k in range(5):
        on = 1.0
        if k == 2 and random.random() < 0.25: on = 0.0      # radio has weeks off
        if k == 3 and random.random() < 0.35: on = 0.0      # posters run in fewer weeks
        ch_spend[k][i] = round(tot * CH_SHARE[k] * on * lognorm(0.25), 2)
rfl_spend = [round(sum(ch_spend[k][i] for k in range(5)), 2) for i in range(N)]

cog_profile = {1: 1.2, 2: 1.1, 3: 1.0, 4: 1.0, 5: 0.9, 6: 0.8, 7: 0.7, 8: 0.7, 9: 0.9, 10: 1.0, 11: 0.5, 12: 0.3}
cog_year = {2023: 1.0, 2024: 1.05, 2025: 0.62}
cog_spend = [round(200 * cog_profile[mons[i]] * cog_year[yrs[i]] * lognorm(0.3), 2) for i in range(N)]
leg_spend = [round((120 if mons[i] in (9, 10, 11) else 30) * lognorm(0.3), 2) for i in range(N)]
brand_spend = [round((260 if mons[i] in (3, 9) else 60 if mons[i] in (4, 10) else 15) * (0.8 if yrs[i] == 2025 else 1) * lognorm(0.3), 2) for i in range(N)]
halo_spend = [round(cog_spend[i] + leg_spend[i] + brand_spend[i], 2) for i in range(N)]
total_media = [round(rfl_spend[i] + halo_spend[i] + 35 * lognorm(0.2), 2) for i in range(N)]

# ---------------------------------------------------------------- the generating model
ANS = {'base': 640, 'brake': 700, 'sale': 6200, 'mem': 75, 'ch': [18.6, 19.5, 19.0, 6.4, 19.9], 'chMem': [75] * 5}
COG_COEF, OTHER_COEF = 3.4, 1.5

def adstock(x, m):
    o, a = [], 0.0
    for v in x:
        a = v + m * a
        o.append(a)
    return o

def model_parts():
    ap_base = [ANS['base'] * v for v in shape_base]
    ap_brake = [-ANS['brake'] * v for v in shape_brake]
    ap_sale = [ANS['sale'] * v for v in shape_sale]
    ap_ch = []
    for k in range(5):
        m = ANS['chMem'][k] / 100
        ap_ch.append([ANS['ch'][k] * (1 - m) * v for v in adstock(ch_spend[k], m)])
    ap_rfl = [sum(ap_ch[k][i] for k in range(5)) for i in range(N)]
    ap_halo = [COG_COEF * cog_spend[i] + OTHER_COEF * (leg_spend[i] + brand_spend[i]) for i in range(N)]
    return ap_base, ap_brake, ap_sale, ap_ch, ap_rfl, ap_halo

# keep the true model above the floor by easing the brake in the deepest weeks, then renormalise
for _ in range(4):
    shape_brake = norm_mean1(shape_brake)
    ap_base, ap_brake, ap_sale, ap_ch, ap_rfl, ap_halo = model_parts()
    for i in range(N):
        tot = ap_base[i] + ap_brake[i] + ap_sale[i] + ap_rfl[i] + ap_halo[i]
        if tot < 220:
            shape_brake[i] = max(0.05, shape_brake[i] - (220 - tot) / ANS['brake'])
shape_brake = norm_mean1(shape_brake)
ap_base, ap_brake, ap_sale, ap_ch, ap_rfl, ap_halo = model_parts()
truth = [ap_base[i] + ap_brake[i] + ap_sale[i] + ap_rfl[i] + ap_halo[i] for i in range(N)]
assert min(truth) > 150, min(truth)

actual = []
for i in range(N):
    noise = random.gauss(0, 0.15 * abs(truth[i]) + 220)
    actual.append(float(max(90, round(truth[i] + noise))))

# single-dial answers fitted to the generating model (least squares through the origin)
def fit_through_origin(x, y):
    return sum(a * b for a, b in zip(x, y)) / sum(a * a for a in x)

m = ANS['mem'] / 100
rfl_x = [(1 - m) * v for v in adstock(rfl_spend, m)]
ANS['rfl'] = round(fit_through_origin(rfl_x, ap_rfl), 1)
ANS['halo'] = round(fit_through_origin(halo_spend, ap_halo), 2)

# consideration and monthly-giver sign-ups
consideration, c = [], 0.503
for i in range(N):
    c += -0.002 / 52 + 0.0000012 * total_media[i] / 100 - 0.00001 * (i % 52) / 52
    consideration.append(round(c + random.gauss(0, 0.003), 4))
# smooth to a slow line
consideration = [round(sum(consideration[max(0, i - 3):i + 4]) / len(consideration[max(0, i - 3):i + 4]), 4) for i in range(N)]
cog_ad = adstock(cog_spend, 0.5)
cog_signups = [float(max(30, round(0.5 * 0.9 * cog_ad[i] * (0.9 if yrs[i] == 2025 else 1) + 60 + random.gauss(0, 25)))) for i in range(N)]

# ---------------------------------------------------------------- budget game channels (2025 plan, £k)
channels = [
    ('TV', 840, 1.72), ('Online video', 760, 1.61), ('Radio', 690, 1.55), ('Paid social', 910, 1.38), ('Regional press', 610, 1.47),
    ('Posters', 470, 0.61), ('Digital audio', 390, 1.49), ('Display', 380, 1.52), ('Paid search', 270, 1.50), ('YouTube', 110, 1.63),
    ('Telemarketing', 90, 1.05), ('Direct mail', 35, 1.66), ('Door drops', 20, 1.75)]
# the modellers' recommended plan: money out of the weak channels, into the strong ones, same total
moves = {'Posters': -0.4, 'Telemarketing': -0.4, 'Paid social': -0.2, 'Regional press': -0.15, 'TV': 0.25, 'Online video': 0.2, 'Radio': 0.12, 'YouTube': 0.5, 'Digital audio': 0.15, 'Display': 0.15}
plan = {n: s * (1 + moves.get(n, 0)) for n, s, _ in channels}
diff = sum(s for _, s, _ in channels) - sum(plan.values())
plan['Display'] += diff  # balance the total exactly
CHANNELS = [{'name': n, 'spend': s, 'roi': r, 'ap': int(round(plan[n]))} for n, s, r in channels]
fix = sum(c['spend'] for c in CHANNELS) - sum(c['ap'] for c in CHANNELS)
CHANNELS[7]['ap'] += fix

def signups(c, sp):
    return 0 if sp <= 0 else (c['roi'] * c['spend'] * 1000 / POUND) / math.sqrt(c['spend']) * math.sqrt(sp)

plan_gain = sum(signups(c, c['ap']) for c in CHANNELS) - sum(signups(c, c['spend']) for c in CHANNELS)
# best possible under the game's rules (each channel within 50% to 150%, same total): greedy marginal search
bud = [c['spend'] for c in CHANNELS]
for _ in range(4000):
    best = None
    for a in range(13):
        for b in range(13):
            if a == b: continue
            step = max(1, round(CHANNELS[a]['spend'] * 0.02))
            if bud[a] - step < CHANNELS[a]['spend'] * 0.5 or bud[b] + step > CHANNELS[b]['spend'] * 1.5: continue
            gain = signups(CHANNELS[b], bud[b] + step) - signups(CHANNELS[b], bud[b]) - (signups(CHANNELS[a], bud[a]) - signups(CHANNELS[a], bud[a] - step))
            if gain > 0 and (best is None or gain > best[0]): best = (gain, a, b, step)
    if not best: break
    _, a, b, step = best
    bud[a] -= step; bud[b] += step
max_gain = sum(signups(c, bud[i]) for i, c in enumerate(CHANNELS)) - sum(signups(c, c['spend']) for c in CHANNELS)
TARGET = int(round(max_gain / 2, -2))

# ---------------------------------------------------------------- headline figures, all derived from the invented data
def year_return(y):
    s = sum(ap_rfl[i] for i in range(N) if yrs[i] == y) * POUND / 1000
    return s / sum(rfl_spend[i] for i in range(N) if yrs[i] == y)

R = {y: round(year_return(y), 2) for y in (2023, 2024, 2025)}
R25 = R[2025]
MARGINAL = round(R25 * 0.5, 2)
MULT = 1.18
LT_EXTRA = round(R25 * (MULT - 1), 2)
R25_LT = round(R25 + LT_EXTRA, 2)
sum_actual = sum(actual)
share = {'adverts': sum(ap_rfl) / sum_actual, 'sales': sum(ap_sale[i] + ap_brake[i] for i in range(N)) / sum_actual, 'base': sum(ap_base) / sum_actual, 'halo': sum(ap_halo) / sum_actual}
share_sales_2025 = sum(ap_sale[i] + ap_brake[i] for i in range(N) if yrs[i] == 2025) / sum(a for a, y in zip(actual, yrs) if y == 2025)

def r2(model, mask=None):
    idx = [i for i in range(N) if mask is None or mask[i]]
    mean = sum(actual[i] for i in idx) / len(idx)
    sse = sum((actual[i] - model[i]) ** 2 for i in idx); sst = sum((actual[i] - mean) ** 2 for i in idx)
    return max(0, 1 - sse / sst)

def mape(model):
    v = [abs(actual[i] - model[i]) / actual[i] for i in range(N) if actual[i] > 2000]
    return sum(v) / len(v)

def unexplained(model):
    a24 = sum(actual[i] for i in range(N) if yrs[i] == 2024 and mons[i] <= 10); m24 = sum(model[i] for i in range(N) if yrs[i] == 2024 and mons[i] <= 10)
    a25 = sum(actual[i] for i in range(N) if yrs[i] == 2025); m25 = sum(model[i] for i in range(N) if yrs[i] == 2025)
    return ((a25 - a24) - (m25 - m24)) / a24

R2 = r2(truth); MAPE = mape(truth); UNEXP = unexplained(truth)
cog_jo = lambda y: sum(cog_spend[i] for i in range(N) if yrs[i] == y and mons[i] <= 10)
CUT_PCT = round((1 - cog_jo(2025) / cog_jo(2024)) * 100)
STUDY_LOST = int(round(COG_COEF * (cog_jo(2024) - cog_jo(2025)), -1))
halo_share_year = lambda y: sum(COG_COEF * cog_spend[i] for i in range(N) if yrs[i] == y) / sum(a for a, yy in zip(actual, yrs) if yy == y)
HS23, HS25 = round(halo_share_year(2023) * 100), round(halo_share_year(2025) * 100)
quiet = [actual[i] for i in range(N) if mons[i] >= 8]
QUIET_MEAN = sum(quiet) / len(quiet)
BASE_Q = int(round(QUIET_MEAN, -2))

# giving chapter (invented)
GIVE = {'fall': 19, 'inflation': 640, 'cut': 230, 'cut_pct': 45, 'per_giver': 250, 'floor_p': 14, 'fy_return': 0.18, 'share3': 68, 'share_fy': 91}
WATERFALL = [['Inflation', -640], ['Monthly Giving adverts cut', -230], ['Brand adverts cut', -80], ['Big Splash adverts cut', -70], ['Consideration slipped', -25], ['Gifts in wills adverts up', 30], ['Less interest in a rival appeal', 35], ['Health awareness campaign halo', 60], ['Challenge events adverts up', 210]]
BRAND = {'pct': '16.5%', 'drift': 0.2, 'lift': 0.5, 'mult_cog': 1.03}
CH_RETURN = [[c['name'], c['roi']] for c in CHANNELS if c['name'] in ('TV', 'Online video', 'Radio', 'Digital audio', 'Regional press', 'Display', 'Paid social', 'YouTube', 'Posters')]
LT_ROWS = [['All Big Splash adverts', R25, LT_EXTRA]] + [[n, r, round(r * (MULT - 1) * (0.6 + (hash(n) % 5) / 6), 2)] for n, r in CH_RETURN] + [['Monthly Giving adverts', GIVE['fy_return'], 0.01]]
TOTAL = [['Big Splash adverts', R25, 0.02], ['Monthly Giving adverts', 0.14, 0.31], ['Brand adverts', 0.21, 0.0], ['All marketing', 0.79, 0.11]]
ALL_RET = round(TOTAL[3][1] + TOTAL[3][2], 2)
BS_TOT = round(TOTAL[0][1] + TOTAL[0][2], 2); MG_TOT = round(TOTAL[1][1] + TOTAL[1][2], 2)

L = lambda a, nd=2: [round(float(v), nd) for v in a]
data = {
    'dates': dates, 'actual': L(actual, 0), 'shapeBase': L(shape_base, 4), 'shapeBrake': L(shape_brake, 4), 'shapeSale': L(shape_sale, 4),
    'rflSpend': rfl_spend, 'haloSpend': halo_spend, 'cogSpend': cog_spend, 'chNames': CH_NAMES, 'chSpend': ch_spend,
    'apBase': L(ap_base, 1), 'apBrake': L(ap_brake, 1), 'apSale': L(ap_sale, 1), 'apRfl': L(ap_rfl, 1), 'apHalo': L(ap_halo, 1), 'apCh': [L(a, 1) for a in ap_ch],
    'cogSignups': cog_signups, 'consideration': consideration, 'totalMedia': total_media,
    'answers': ANS, 'channels': CHANNELS, 'waterfall': WATERFALL,
}
json.dump(data, open(os.path.join(WEB, 'fiction_data.json'), 'w'))

# ---------------------------------------------------------------- rewrite the page
html = open(os.path.join(WEB, 'game_template.html')).read()
count = 0

def rep(old, new, n=None):
    global html, count
    k = html.count(old)
    assert k > 0, 'not found: ' + old[:80]
    if n is not None:
        assert k == n, f'expected {n} found {k}: {old[:80]}'
    html = html.replace(old, new); count += k

def fm(v, d=0):
    return f'{v:,.{d}f}'

def p(v):  # pounds
    return f'£{v:.2f}'

def pence(v):
    return f'{round(v * 100)}p'

# --- names
rep("Every weekly figure is Analytic Partners' own, from the charts in the Wave 1 results deck (July 2026) and the methodology document (17 July 2026): real Race for Life sign-ups, spend by channel and product, their weekly decomposition by driver, brand consideration and new Committed Givers, January 2023 to October 2025. The numbers you guess are fitted to that decomposition. Sponsorship income is counted at about £90 per sign-up, as AP do. CRUK confidential, internal use only.",
    f"Lantern is an invented charity and Big Splash is its invented sponsored swim. Every weekly figure on this site is made up: sign-ups, spend by channel and product, the weekly split by driver, brand consideration and new monthly givers, January 2023 to October 2025, were all generated from a known model plus noise so that the game has an answer to check you against. Sponsorship income is counted at £{POUND} per sign-up. No real charity, agency or dataset is shown here.")
rep('<title>The Econometrics Game</title>', '<title>The Econometrics Game</title>')
rep("Race for Life", "Big Splash")
rep("Committed Givers", "monthly givers"); rep("Committed Giving", "Monthly Giving"); rep("Legacy", "Gifts in wills")
rep("All CRUK media", "All Lantern media", 2)
rep("Analytic Partners built it in 2026. Every number in this game is theirs.", "It is the model the invented data was built from, so it is the answer key.")
rep("Not a Wave 1 finding.", "Not a finding of the study.")
html = re.sub(r"\bAP(’s)?\b(?!\[| =)", lambda m_: "the modellers’" if m_.group(1) else "the modellers", html)
assert "the modellers’ plan" in html
rep("Wave 1", "the study")
for phrase in ("the study", "the modellers"):
    html = re.sub(r"(^|['`.:>] ?)" + phrase, lambda m_: m_.group(1) + phrase[0].upper() + phrase[1:], html, flags=re.M)
rep("Wave 1", "the study") if "Wave 1" in html else None

# --- pounds per sign-up and the response dial
rep("POUND = 90", f"POUND = {POUND}")
rep("* 1000 / 90)", f"* 1000 / {POUND})", 2)
rep("About 16 sign-ups per £1,000.", f"About {round(ANS['rfl'])} sign-ups per £1,000.")
rep("hint:'Somewhere between 5 and 30.'", "hint:'Somewhere between 5 and 30.'")

# --- base, brake, sale
rep("tick:'About 900 a week: people who already consider us", f"tick:'About {fm(ANS['base'])} a week: people who already consider us")
rep("opts:['About 900','About 9,000','Nobody at all']", f"opts:['About {fm(BASE_Q)}','About {fm(BASE_Q * 10)}','Nobody at all']")
rep("const v = [900, 9000, 0][q - 1]", f"const v = [{BASE_Q}, {BASE_Q * 10}, 0][q - 1]")
rep("Nine thousand a week would be a spike every week.", f"{fm(BASE_Q * 10)} a week would be a spike every week.")
rep("Race for Life has a small base of a few hundred sign-ups a week.", "Big Splash has a small base of a few hundred sign-ups a week.") if "Race for Life has" in html else None
rep("tick:'About 950 a week are lost to price and the shut window.", f"tick:'About {fm(ANS['brake'])} a week are lost to price and the shut window.")
fifth = 'a fifth' if abs(share_sales_2025 - 0.2) < 0.04 else 'a quarter' if abs(share_sales_2025 - 0.25) < 0.04 else f'{round(share_sales_2025 * 100)}%'
rep("tick:'A sale week adds about 7,600 sign-ups. Sales explained a fifth of 2025 sign-ups.'", f"tick:'A sale week adds about {fm(ANS['sale'])} sign-ups. Sales explained about {fifth} of 2025 sign-ups.'")
rep("Sale weeks drove a fifth of sign-ups.", f"Sale weeks drove about {fifth} of sign-ups.")

# --- returns by year and the ROI quiz
rep("const YEAR_RETURN = [['2023', 1.50], ['2024', 1.57], ['2025', 1.66]];", f"const YEAR_RETURN = [['2023', {R[2023]}], ['2024', {R[2024]}], ['2025', {R25}]];")
rep("opts:['£1.66','16p','£16'], tick:'Yes. £1.66 on average, up from £1.50 in 2023.", f"opts:['{p(R25)}','{pence(R25 / 10)}','£{round(R25 * 10)}'], tick:'Yes. {p(R25)} on average, against {p(R[2023])} in 2023.")
rep("const v = q ? [1.66, 0.16, 16][q - 1] : null;", f"const v = q ? [{R25}, {round(R25 / 10, 2)}, {round(R25 * 10)}][q - 1] : null;")
rep("Your line lands on the 2025 bar. £1.66 back for every £1, and rising each year.", f"Your line lands on the 2025 bar. {p(R25)} back for every £1.")
rep("No channel returns £16 for £1.", f"No channel returns £{round(R25 * 10)} for £1.")

# --- halo
rep("tick:'About 3 per £1,000. Small per pound, but the spend is large. It adds up to 15% of sign-ups.'", f"tick:'About {ANS['halo']:.1f} per £1,000. Small per pound, but the spend is large. It adds up to {round(share['halo'] * 100)}% of sign-ups.'")
rep("In 2025 we spent 42% less on Monthly Giving adverts.", f"In 2025 we spent {CUT_PCT}% less on Monthly Giving adverts.")
rep("The study said 9,139.", f"The study said {fm(STUDY_LOST)}.", 2)
rep("The modellers found Monthly Giving adverts stronger than Gifts in wills or Brand.", "The modellers found Monthly Giving adverts stronger than Gifts in wills or Brand.")
rep("Monthly Giving adverts were 13% of Big Splash sign-ups in 2023 and 5% in 2025.", f"Monthly Giving adverts were {HS23}% of Big Splash sign-ups in 2023 and {HS25}% in 2025.")
rep("The Monthly Giving halo is 2 to 4 sign-ups per £1,000.", f"The Monthly Giving halo is {COG_COEF - 0.8:.1f} to {COG_COEF + 0.8:.1f} sign-ups per £1,000.")

# --- channels
rep("const CH_RETURN = [['DRTV', 1.95], ['Online video', 1.86], ['Radio', 1.75], ['Digital audio', 1.71], ['Regional', 1.71], ['Display', 1.74], ['Paid social', 1.59], ['YouTube', 1.85], ['Posters', 0.54]];",
    "const CH_RETURN = " + json.dumps(CH_RETURN) + ";")
lt_old = html[html.index("const LT = ["):html.index("\n", html.index("const LT = ["))]
html = html.replace(lt_old, "const LT = " + json.dumps(LT_ROWS) + ";"); count += 1
rep("const TOTAL_RETURNS = [['Big Splash adverts', 1.66, 0.01], ['Monthly Giving adverts', 0.17, 0.38], ['Brand adverts', 0.24, 0], ['All marketing', 0.88, 0.13]];", "const TOTAL_RETURNS = " + json.dumps(TOTAL) + ";")
posters = [c for c in CHANNELS if c['name'] == 'Posters'][0]['roi']
others_min = min(c['roi'] for c in CH_RETURN_ROWS) if (CH_RETURN_ROWS := [dict(name=n, roi=r) for n, r in CH_RETURN if n != 'Posters']) else 0
rep("tick:'Posters: about 54p back per £1, far below the rest, with a wide range either side.'", f"tick:'Posters: about {pence(posters)} back per £1, far below the rest, with a wide range either side.'")
rep("const pick = q ? ['Posters', 'DRTV', 'Paid social'][q - 1] : null;", "const pick = q ? ['Posters', 'TV', 'Paid social'][q - 1] : null;")
rep("Posters, the lit bar, return about 54p per £1. Every other channel returns £1.59 or more.", f"Posters, the lit bar, return about {pence(posters)} per £1. Every other channel returns {p(others_min)} or more.")
rep("pick === 'DRTV' ? 'TV, the strongest bar'", "pick === 'TV' ? 'TV, the strongest bar'")
rep("TV, video, radio, posters, and digital, social and the rest.", "TV, video, radio, posters, and digital, social and the rest.")

# --- tests
rep("['Pattern explained', r2(F), 0.94, pct, v => v >= 0.85]", f"['Pattern explained', r2(F), {R2:.2f}, pct, v => v >= 0.85]")
rep("['Typical weekly miss', mape(F), 0.28, pct, v => v <= 0.35]", f"['Typical weekly miss', mape(F), {MAPE:.2f}, pct, v => v <= 0.35]")
rep("['Unexplained 2025 change', unexplained(F), 0.014,", f"['Unexplained 2025 change', unexplained(F), {UNEXP:.3f},")
rep("The study explains 94% of the weekly pattern, predicts weeks it never saw and left 1.4% of the 2025 fall unexplained.", f"The study explains {round(R2 * 100)}% of the weekly pattern, predicts weeks it never saw and left {abs(UNEXP) * 100:.1f}% of the 2025 change unexplained.")
rep("The study explains 94%. The modellers’ guide is above 85%.", f"The study explains {round(R2 * 100)}%. The modellers’ guide is above 85%.")
rep("The model explains 94% of the weekly pattern in sign-ups.", f"The model explains {round(R2 * 100)}% of the weekly pattern in sign-ups.")
rep("28% on weeks with more than 2,000 sign-ups.", f"{round(MAPE * 100)}% on weeks with more than 2,000 sign-ups.")
rep("On an average season week the model is within about a quarter of the real number.", f"On an average season week the model is within about {round(MAPE * 100)}% of the real number.")
rep("1.4% of sign-ups in the 2024 to 2025 comparison, well inside the modellers’ tolerance.", f"{abs(UNEXP) * 100:.1f}% of sign-ups in the 2024 to 2025 comparison, well inside the modellers’ tolerance.")
rep("The model left 1.4% unexplained, which is small.", f"The model left {abs(UNEXP) * 100:.1f}% unexplained, which is small.")
rep("The modellers tested 18 marketing drivers and seven others", "The modellers tested 20 marketing drivers and six others")
rep("A 12% lower fee in 2025 added about 2,500 sign-ups.", "A cheaper entry fee in 2025 added about 1,800 sign-ups.")
rep("The modellers read 148 weeks of sign-ups", "The modellers read 148 weeks of sign-ups")

# --- decomposition sentence
def word_share(v):
    return 'about half' if abs(v - 0.5) < 0.06 else 'about a third' if abs(v - 0.33) < 0.05 else 'about a quarter' if abs(v - 0.25) < 0.04 else 'a fifth' if abs(v - 0.2) < 0.03 else 'a tenth' if abs(v - 0.1) < 0.03 else f'{round(v * 100)}%'
rep("Say it in the room: about half of Big Splash sign-ups come from Big Splash adverts, a fifth from sales and 15% from other causes’ adverts.",
    f"Say it in the room: {word_share(share['adverts'])} of Big Splash sign-ups come from Big Splash adverts, {word_share(share['sales'])} from sales and {round(share['halo'] * 100)}% from other causes’ adverts.")
rep("The study says about half, a fifth, a tenth and 15 to 20%.", f"The study says {word_share(share['adverts'])}, {word_share(share['sales'])}, {word_share(share['base'])} and {round(share['halo'] * 100)}%.")

# --- marginal return and budget
rep("The average £1 returned £1.66. The last £1 returned 82p.", f"The average £1 returned {p(R25)}. The last £1 returned {pence(MARGINAL)}.")
rep("opts:['Not yet: move money first','Yes, every £1 returns £1.66','No, cut everything']", f"opts:['Not yet: move money first','Yes, every £1 returns {p(R25)}','No, cut everything']")
rep("Moving money between channels was worth about £1.1m before adding a single pound.", "Moving money between channels was worth about £0.8m before adding a single pound.")
rep("The average pound returned £1.66 but the last pound returned 82p. Moving money is worth £1.1m before adding a single pound.", f"The average pound returned {p(R25)} but the last pound returned {pence(MARGINAL)}. Moving money is worth £0.8m before adding a single pound.")
rep("In 2025 the last pound of Big Splash adverts returned 82p while the average pound returned £1.66.", f"In 2025 the last pound of Big Splash adverts returned {pence(MARGINAL)} while the average pound returned {p(R25)}.")
rep("The average pound returned £1.66 but the last pound returned 82p, so adding budget needs care.", f"The average pound returned {p(R25)} but the last pound returned {pence(MARGINAL)}, so adding budget needs care.")
rep("The modellers’ real curves differ by channel, so their plan gained 12,300 sign-ups with them and about 600 here.", f"The modellers’ real curves differ by channel, so their plan gained 9,800 sign-ups with them and about {fm(round(plan_gain, -2))} here.")
rep("const won = okTotal && extra >= 1000;", f"const won = okTotal && extra >= {TARGET};")
rep("The most this table allows is about 2,000.", f"The most this table allows is about {fm(round(max_gain, -2))}.")
rep("Move money from the weakest channel to the strongest. Find 1,000 more.", f"Move money from the weakest channel to the strongest. Find {fm(TARGET)} more.")

# --- brand
rep("opts:['£2.01','£1.66','82p'], tick:'Counting the long-term effect takes Big Splash adverts from £1.66 to £2.01 per pound.", f"opts:['{p(R25_LT)}','{p(R25)}','{pence(MARGINAL)}'], tick:'Counting the long-term effect takes Big Splash adverts from {p(R25)} to {p(R25_LT)} per pound.")
rep("[2.01, 1.66, 0.82][q - 1]", f"[{R25_LT}, {R25}, {MARGINAL}][q - 1]", 2)
rep("Your line touches the top of the first bar: £1.66 short term plus £0.35 long term is <b>£2.01</b>.", f"Your line touches the top of the first bar: {p(R25)} short term plus {p(LT_EXTRA)} long term is <b>{p(R25_LT)}</b>.")
rep("Your line is below every bar. 82p is the return on the last pound, not the average.", f"Your line is below every bar. {pence(MARGINAL)} is the return on the last pound, not the average.")
rep("Counting the long-term effect takes Big Splash adverts from £1.66 to £2.01 per pound.", f"Counting the long-term effect takes Big Splash adverts from {p(R25)} to {p(R25_LT)} per pound.")
rep("1.21 for Big Splash adverts (£1.66 becomes £2.01) and 1.04 for Monthly Giving.", f"{MULT} for Big Splash adverts ({p(R25)} becomes {p(R25_LT)}) and {BRAND['mult_cog']} for Monthly Giving.")
rep("'18.8%'", f"'{BRAND['pct']}'")
rep("18.8% of today’s consideration traces back to past media. It slips 0.25 points a year. £1m lifts it 0.4 points.", f"{BRAND['pct']} of today’s consideration traces back to past media. It slips {BRAND['drift']} points a year. £1m lifts it {BRAND['lift']} points.")
rep("each extra £1m lifts consideration by about 0.4 points over one to two years, against a drift of 0.25 points a year with nothing extra.", f"each extra £1m lifts consideration by about {BRAND['lift']} points over one to two years, against a drift of {BRAND['drift']} points a year with nothing extra.")
rep("const last = D.consideration[N - 1], drift = -0.0025, lift = 0.004 * S.brandExtra;", f"const last = D.consideration[N - 1], drift = -{BRAND['drift'] / 100}, lift = {BRAND['lift'] / 100} * S.brandExtra;")
rep("Nearly a fifth of the people who consider us", "About a sixth of the people who consider us", 2)
rep("Nearly a fifth of it was built by past adverts.", "About a sixth of it was built by past adverts.")

# --- giving
g = GIVE
rep("Monthly Giving looks like 9p per £1 because the model only counts new givers inside the window", f"Monthly Giving looks like {g['floor_p']}p per £1 because the model only counts new givers inside the window")
rep("The bars show why new monthly givers fell 23% in 2025. Inflation and the advertising cut each took about a quarter.", f"The bars show why new monthly givers fell {g['fall']}% in 2025. Inflation was the biggest cause and the advertising cut came next.")
rep("Why does Monthly Giving show only 9p per £1?", f"Why does Monthly Giving show only {g['floor_p']}p per £1?")
rep("Monthly Giving shows 9p per pound because the model only counts new givers", f"Monthly Giving shows {g['floor_p']}p per pound because the model only counts new givers")
rep("'73%'", f"'{g['share3']}%'")
rep("In FY25/26 it was 94%, because inflation ate the base.", f"In FY25/26 it was {g['share_fy']}%, because inflation ate the base.")
rep("The study saw a 50% cut cost 307 new givers in seven months.", f"The study saw a {g['cut_pct']}% cut cost {g['cut']} new givers in seven months.")
rep("const lost = Math.round(S.giveCut / 50 * 307);", f"const lost = Math.round(S.giveCut / {g['cut_pct']} * {g['cut']});")
rep("About £${fmt(lost * 300 / 1000)}k of income in the modelled window.", f"About £${{fmt(lost * {g['per_giver']} / 1000)}}k of income in the modelled window.")
rep("The study counts new givers inside the window only, about £300 each. This shows why the modellers call their 9p to 20p a floor.", f"The study counts new givers inside the window only, about £{g['per_giver']} each. This shows why the modellers call their {g['floor_p']}p to {round(g['fy_return'] * 100)}p a floor.")
rep("Inflation took <b>904</b> and the advertising cut <b>307</b>.", f"Inflation took <b>{g['inflation']}</b> and the advertising cut <b>{g['cut']}</b>.")
rep("but it is not why the return reads 9p.", f"but it is not why the return reads {g['floor_p']}p.")
rep("data:years.map(y => 0.2 * y)", f"data:years.map(y => {g['fy_return']} * y)")
rep("hlines:[{ at:0.2, color:css('--navy'), label:'The study counts this: £0.20 in FY25/26' }]", f"hlines:[{{ at:{g['fy_return']}, color:css('--navy'), label:'The study counts this: {p(g['fy_return'])} in FY25/26' }}]")
rep("The study counts the first bar only: <b>£0.20</b> per £1 in FY25/26 (9p across all Monthly Giving media in 2025).", f"The study counts the first bar only: <b>{p(g['fy_return'])}</b> per £1 in FY25/26 ({g['floor_p']}p across all Monthly Giving media in 2025).")
rep("returns <b>${gbp(0.2 * S.giveYears)}</b>", f"returns <b>${{gbp({g['fy_return']} * S.giveYears)}}</b>")
rep("Inflation cost Monthly Giving about 900 new givers in FY25/26, as much as the advertising cut.", f"Inflation cost Monthly Giving about {g['inflation']} new givers in FY25/26, more than the advertising cut.")
rep("Inflation and the advertising cut each cost about a quarter of new givers. One we chose, one we did not.", "Inflation and the advertising cut were the two biggest reasons new givers fell. One we chose, one we did not.")

# --- finish
rep("Every £1 of marketing returned £1.01 of donation income in the short term (2025).", f"Every £1 of marketing returned {p(ALL_RET)} of donation income in the short term (2025).")
rep("Big Splash adverts return <b>£1.67</b>, Monthly Giving <b>£0.54</b> once its halo onto Big Splash is counted, Brand <b>£0.24</b>. All marketing together: <b>£1.01</b>. Gifts in wills income and the existing givers are not in these numbers, so they are floors.",
    f"Big Splash adverts return <b>{p(BS_TOT)}</b>, Monthly Giving <b>{p(MG_TOT)}</b> once its halo onto Big Splash is counted, Brand <b>{p(TOTAL[2][1])}</b>. All marketing together: <b>{p(ALL_RET)}</b>. Gifts in wills income and the existing givers are not in these numbers, so they are floors.")
rep("Across all marketing the short-term return was £1.01.", f"Across all marketing the short-term return was {p(ALL_RET)}.")
rep("£1.66 back for every £1 in.", f"{p(R25)} back for every £1 in.")
rep("Sponsorship income only, about £90 per sign-up.", f"Sponsorship income only, £{POUND} per sign-up.")
rep("Every £1 of Big Splash adverts returned £1.66 of sponsorship income in 2025.", f"Every £1 of Big Splash adverts returned {p(R25)} of sponsorship income in 2025.")
rep("'econ-game-v4'", "'econ-game-fiction-v1'")
rep("Memory of 80%: the effect fades slowly over months.", f"Memory of {ANS['mem']}%: the effect fades slowly over months.")
rep("Source: The study.", "Source: the study.")
rep("Ten sentences you can now say in the room, each one true to the study.", "Ten sentences you can now say in the room, each one true to the study behind this invented charity.")
rep("The everyday picture first, then what it meant in the study, then the sentence to say out loud.", "The everyday picture first, then what it meant in the study, then the sentence to say out loud.")

# anything left over?
left = re.findall(r"CRUK|Cancer|Analytic|Wave 1|Race for Life|Committed|Legacy|Macmillan|DRTV|1\.66|2\.01|9,139|18\.8|£90\b|\b82p", html)
assert not left, left
open(os.path.join(WEB, 'fiction_template.html'), 'w').write(html)
open(os.path.join(WEB, 'fiction_index.html'), 'w').write(html.replace('/*DATA*/{}', json.dumps(data, separators=(',', ':'))))
print('replacements', count)
print('answers', ANS, 'halo coef', COG_COEF)
print('returns', R, 'marginal', MARGINAL, 'LT', LT_EXTRA, R25_LT)
print('R2 %.3f MAPE %.3f UNEXP %.4f' % (R2, MAPE, UNEXP))
print('shares', {k: round(v, 3) for k, v in share.items()}, 'sales 2025 share %.3f' % share_sales_2025)
print('cut pct', CUT_PCT, 'study lost', STUDY_LOST, 'halo shares', HS23, HS25, 'quiet mean', round(QUIET_MEAN), 'base q', BASE_Q)
print('plan gain', round(plan_gain), 'max gain', round(max_gain), 'target', TARGET)
print('actual min/max/mean', min(actual), max(actual), round(sum(actual) / N), 'truth min', round(min(truth)))
print('r2 2025 holdout %.3f' % r2(truth, [y == 2025 for y in yrs]))
