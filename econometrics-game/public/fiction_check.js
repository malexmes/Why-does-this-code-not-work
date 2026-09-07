// Walk every step and tab of the fictional edition headlessly and report problems.
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const p = await b.newPage({ viewport: { width: 1200, height: 900 } });
  const errs = [];
  p.on('pageerror', e => errs.push('pageerror: ' + e.message));
  p.on('console', m => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  await p.route(/^https?:\/\//, r => r.abort());
  await p.goto('file://' + path.join(__dirname, 'fiction_test_local.html'));
  await p.waitForTimeout(500);
  const bad = txt => /undefined|NaN|Infinity|Race for Life|CRUK|Analytic|Wave 1|Committed|Legacy|\bAP\b/.test(txt);
  const report = [];
  const snap = async label => {
    const t = await p.evaluate(() => ({ title: document.getElementById('title').textContent, lead: document.getElementById('lead').textContent,
      fb: [...document.querySelectorAll('.feedback')].map(e => e.textContent).join(' | '), seeing: [...document.querySelectorAll('.seeing')].map(e => e.textContent).join(' | '),
      canvases: document.querySelectorAll('canvas').length, body: document.body.innerText }));
    if (bad(t.body)) report.push(label + ' BAD TEXT: ' + (t.body.match(/.{0,40}(undefined|NaN|Infinity|Race for Life|CRUK|Analytic|Wave 1|Committed|Legacy|\bAP\b).{0,40}/) || [''])[0]);
    return t;
  };
  const steps = await p.evaluate(() => STEPS.map(s => ({ type: s.type, dial: s.dial, q: s.q, pic: s.pic })));
  for (let i = 0; i < steps.length; i++) {
    const s = steps[i];
    await p.evaluate(i => { S.step = i + 1; S.stage = 1; S.reveal = false; render(); }, i);
    let t = await snap('step ' + (i + 1));
    let line = `step ${i + 1} [${s.type}${s.dial ? ':' + s.dial : ''}] ${t.title} | canvases ${t.canvases}`;
    if (s.type === 'dial') {
      await p.evaluate(d => { S[d] = dialSpec[d].ans; render(); }, s.dial);
      t = await snap('step ' + (i + 1) + ' answer');
      line += ` | at answer: ${t.fb.slice(0, 90)}`;
      if (!t.fb.includes('✓')) report.push(`step ${i + 1} dial ${s.dial} not ticked at answer: ${t.fb}`);
      await p.evaluate(() => { S.reveal = true; render(); }); t = await snap('step ' + (i + 1) + ' reveal');
    } else if (s.type === 'quiz') {
      for (let k = 1; k <= 3; k++) { await p.evaluate(([q, k]) => { S.quiz[q - 1] = k; render(); }, [s.q, k]); t = await snap(`step ${i + 1} option ${k}`); line += `\n     opt ${k}: ${t.fb.slice(0, 70)} // ${t.seeing.slice(0, 90)}`; }
      await p.evaluate(q => { S.quiz[q - 1] = 1; render(); }, s.q);
    } else if (s.type === 'dials5') {
      await p.evaluate(() => { S.ch = A.ch.slice(); render(); }); t = await snap('step ' + (i + 1) + ' answer'); line += ` | ${t.fb.slice(0, 80)}`;
      if (!t.fb.includes('✓')) report.push('dials5 not ticked at answer: ' + t.fb);
    } else if (s.type === 'tests') {
      line += ` | ${t.fb.slice(0, 80)} | ` + await p.evaluate(() => [...document.querySelectorAll('table.tests td.n')].map(e => e.textContent).join(' '));
      if (!t.fb.includes('✓')) report.push('tests not all green at the answers: ' + t.fb);
    } else if (s.type === 'budget') {
      await p.evaluate(() => { S.budAp = true; render(); }); t = await snap('budget plan'); line += ` | modellers plan: ${t.fb.slice(0, 90)}`;
      if (!t.fb.includes('✓')) report.push('budget: the modellers plan does not win: ' + t.fb);
      await p.evaluate(() => { S.quiz[7] = 1; render(); });
    }
    line += `\n     seeing: ${t.seeing.slice(0, 160)}`;
    console.log(line);
  }
  for (const st of [2, 3, 4]) {
    await p.evaluate(st => { S.stage = st; render(); }, st);
    let t = await snap('stage ' + st);
    console.log(`stage ${st}: ${t.title} | canvases ${t.canvases}\n     seeing: ${t.seeing.slice(0, 220)}`);
    if (st === 2) { for (const k of [1, 2, 3]) { await p.evaluate(k => { S.quiz[8] = k; render(); }, k); t = await snap('brand q ' + k); console.log(`     brand q${k}: ${t.fb.slice(0, 80)}`); } await p.evaluate(() => { S.brandExtra = 3; render(); }); t = await snap('brand slider'); console.log('     brand slider: ' + t.seeing.slice(0, 160)); }
    if (st === 3) { for (const k of [1, 2, 3]) { await p.evaluate(k => { S.quiz[9] = k; render(); }, k); t = await snap('giving q ' + k); console.log(`     giving q${k}: ${t.fb.slice(0, 80)}`); } await p.evaluate(() => { S.giveCut = 60; S.giveYears = 5; render(); }); t = await snap('giving sliders'); console.log('     giving sliders: ' + t.seeing.slice(0, 220)); }
    await p.screenshot({ path: path.join(__dirname, `fiction_stage${st}.png`), fullPage: true });
  }
  await p.evaluate(() => { S.stage = 1; S.step = 14; render(); });
  await p.screenshot({ path: path.join(__dirname, 'fiction_halo.png'), fullPage: true });
  const words = await p.evaluate(() => { document.getElementById('wordsBtn').click(); return document.getElementById('wordsTable').innerText; });
  if (bad(words)) report.push('WORDS BAD TEXT: ' + (words.match(/.{0,40}(undefined|NaN|Race for Life|CRUK|Analytic|Wave 1|Committed|Legacy|\bAP\b).{0,40}/) || [''])[0]);
  console.log('\nERRORS:', errs.length ? errs : 'none');
  console.log('PROBLEMS:', report.length ? report : 'none');
  await b.close();
})();
