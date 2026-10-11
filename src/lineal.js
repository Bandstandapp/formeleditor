/* Lineal über der Vorschau, wie in Word: cm ab dem linken Seitenrand, graue Seitenränder links und rechts.
   Zwei Reiter legen fest, wo eine abgesetzte Formel beginnt (Einzug von links) und wo ihre Gleichungsnummer
   rechtsbündig endet. Seitenbreite und Ränder kommen aus dem Word-Dokument (Abschnitt an der Cursorstelle).
   Mit „Werte für Formel übernehmen“ gelten die Werte beim Einfügen und Ersetzen; der Knopf darunter richtet
   alle Formeln im Dokument danach aus (nummern.js). */
(function () {
  const q = s => document.querySelector(s);
  const vorschau = q('section[aria-label="Ergebnis"]');
  const paper = vorschau && vorschau.querySelector('.paper');
  if (!paper || !window.FeWord) return;
  const STEP = 0.25;                                   // Raster wie im Word-Lineal
  const css = document.createElement('style');
  css.textContent = `
.fe-ruler { display: grid; gap: 4px; }
.fe-ruler[hidden] { display: none !important; }
.fe-ruler-bar { position: relative; height: 46px; touch-action: none; user-select: none; -webkit-user-select: none; }
.fe-ruler-bar svg { position: absolute; left: 0; top: 0; display: block; color: var(--fg); }
.fe-ruler-bar .fe-mark { all: unset; position: absolute; top: 22px; width: 22px; height: 24px; margin-left: -11px; cursor: ew-resize; display: grid; place-items: start center; }
.fe-ruler-bar .fe-mark svg { position: static; }
.fe-ruler-bar .fe-mark:focus-visible { outline: 3px solid var(--focus); border-radius: 4px; }
.fe-ruler-bar .fe-mark.off { opacity: .35; }
.fe-ruler-vals { display: flex; flex-wrap: wrap; gap: 4px 14px; font-size: .95rem; align-items: center; }
.fe-ruler-vals label { display: inline-flex; align-items: center; gap: 6px; }
.fe-ruler-vals input { font: inherit; font-size: 1rem; width: 4.2em; text-align: right; padding: 4px 6px; border: 2px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--fg); }
.fe-ruler-vals input:focus-visible { outline: 3px solid var(--focus); }
.fe-ruler-vals .sw { display: inline-block; width: 0; height: 0; border-left: 6px solid transparent; border-right: 6px solid transparent; border-bottom: 10px solid currentColor; }
.fe-ruler-vals .f .sw { color: var(--accent); }
.fe-ruler-vals .n .sw { color: #c25e00; }
.fe-ruler-vals .n.off { opacity: .45; }
.fe-ruler .fe-check { font-size: .95rem; }
.fe-ruler #fe-align.btn { justify-self: start; text-align: left; width: auto; min-height: 0; padding: 8px 12px; font-size: .95rem; }
.fe-ruler .fe-pbar { height: 12px; border-radius: 6px; background: var(--line); overflow: hidden; }
.fe-ruler .fe-pbar > span { display: block; height: 100%; width: 0; background: #e8730c; border-radius: 6px; transition: width .2s; }
.fe-ruler .fe-pbar[hidden], .fe-ruler #fe-align-msg:empty { display: none; }
.fe-ruler #fe-align-msg { margin: 0; }`;
  document.head.appendChild(css);

  const box = document.createElement('div');
  box.className = 'fe-ruler'; box.hidden = true;
  box.innerHTML = '<div class="fe-ruler-bar"><svg class="scale" aria-hidden="true"></svg>' +
    '<button type="button" class="fe-mark" data-k="f" role="slider" aria-label="Formel beginnt bei (cm vom linken Rand)">' +
    '<svg width="16" height="14" viewBox="0 0 16 14"><path d="M8 0 L16 13 L0 13 Z" fill="var(--accent)"/></svg></button>' +
    '<button type="button" class="fe-mark" data-k="n" role="slider" aria-label="Gleichungsnummer endet bei (cm vom linken Rand)">' +
    '<svg width="16" height="14" viewBox="0 0 16 14"><path d="M8 0 L16 13 L0 13 Z" fill="#c25e00"/></svg></button></div>' +
    '<div class="fe-ruler-vals"><label class="f"><span class="sw"></span>Formel ab <input type="text" inputmode="decimal" data-k="f" aria-label="Formel ab, in cm"> cm</label>' +
    '<label class="n"><span class="sw"></span>Nummer endet bei <input type="text" inputmode="decimal" data-k="n" aria-label="Nummer endet bei, in cm"> cm</label></div>' +
    '<label class="fe-check"><input type="checkbox" id="fe-lineal-use"> Werte für Formel übernehmen</label>' +
    '<button type="button" class="btn primary" id="fe-align">Formeln und Nummern im Dokument wie im Lineal ausrichten</button>' +
    '<div class="fe-pbar" id="fe-align-bar" role="presentation" hidden><span></span></div><p class="hint" id="fe-align-msg" role="status"></p>';
  paper.before(box);
  const bar = box.querySelector('.fe-ruler-bar'), scale = box.querySelector('svg.scale');
  const marks = {f: box.querySelector('.fe-mark[data-k="f"]'), n: box.querySelector('.fe-mark[data-k="n"]')};
  const inputs = {f: box.querySelector('input[data-k="f"]'), n: box.querySelector('input[data-k="n"]')};
  // Beim Einfügen und Ersetzen die Werte aus dem Lineal nehmen (Standard: ja)
  const use = box.querySelector('#fe-lineal-use');
  try { use.checked = localStorage.getItem('fe-lineal-use') !== '0'; } catch (e) { use.checked = true; }
  use.addEventListener('change', () => { try { localStorage.setItem('fe-lineal-use', use.checked ? '1' : '0'); } catch (e) {} });

  // Seite in cm (A4, 2,5 cm Ränder, bis Word etwas anderes meldet)
  let page = {w: 21, l: 2.5, r: 2.5};
  const tw = () => Math.max(1, page.w - page.l - page.r);   // Breite des Satzspiegels
  let pos = {f: 0, n: null};                                 // cm ab linkem Seitenrand des Satzspiegels; n null = rechter Rand
  try { const s = JSON.parse(localStorage.getItem('fe-lineal') || '{}'); if (isFinite(s.f)) pos.f = s.f; if (isFinite(s.n)) pos.n = s.n; } catch (e) {}
  let numOn = true;
  const snap = v => Math.round(v / STEP) * STEP;
  const nVal = () => pos.n == null ? tw() : pos.n;
  function clamp() {
    pos.f = Math.min(Math.max(0, pos.f), Math.max(0, tw() - 2));
    if (pos.n != null) pos.n = Math.min(Math.max(pos.f + 1, pos.n), tw());
  }
  const fmt = v => (Math.round(v * 100) / 100).toLocaleString('de-DE', {maximumFractionDigits: 2});
  const save = () => { try { localStorage.setItem('fe-lineal', JSON.stringify({f: pos.f, n: pos.n})); } catch (e) {} };

  function draw() {
    const W = bar.clientWidth;
    if (!W) return;
    const s = W / page.w, x = cm => (page.l + cm) * s;
    const fg = 'currentColor', every = s >= 20 ? 4 : 8;      // Zahlen jeden cm, bei schmaler Spalte alle 2 cm
    let h = '<rect x="0" y="2" width="' + W + '" height="18" rx="3" fill="var(--bg)" stroke="var(--line)"/>' +
      '<rect x="' + x(0) + '" y="2" width="' + tw() * s + '" height="18" fill="var(--surface)" stroke="var(--line)"/>';
    for (let c = -Math.floor(page.l / STEP) * STEP; c <= page.w - page.l + 1e-9; c += STEP) {
      const r = Math.round(c / STEP), px = x(c).toFixed(1);
      if (r % every === 0) {
        if (r) h += '<text x="' + px + '" y="15.5" text-anchor="middle" font-size="11" fill="' + fg + '" font-family="system-ui, sans-serif">' + Math.abs(r / 4) + '</text>';
      } else if (r % 2 === 0 || s >= 24) h += '<line x1="' + px + '" x2="' + px + '" y1="' + (r % 2 ? 9 : 7) + '" y2="' + (r % 2 ? 13 : 15) + '" stroke="' + fg + '" stroke-opacity=".55"/>';
    }
    // Hilfslinien unter den Reitern
    h += '<line x1="' + x(pos.f) + '" x2="' + x(pos.f) + '" y1="2" y2="20" stroke="var(--accent)" stroke-width="2"/>';
    h += '<line x1="' + x(nVal()) + '" x2="' + x(nVal()) + '" y1="2" y2="20" stroke="#c25e00" stroke-width="2" stroke-opacity="' + (numOn ? 1 : .35) + '"/>';
    scale.setAttribute('width', W); scale.setAttribute('height', 22); scale.setAttribute('viewBox', '0 0 ' + W + ' 22');
    scale.innerHTML = h;
    marks.f.style.left = x(pos.f) + 'px';
    marks.n.style.left = x(nVal()) + 'px';
    marks.n.classList.toggle('off', !numOn);
    box.querySelector('.fe-ruler-vals .n').classList.toggle('off', !numOn);
    for (const k of ['f', 'n']) {
      const v = k === 'f' ? pos.f : nVal();
      marks[k].setAttribute('aria-valuenow', v); marks[k].setAttribute('aria-valuetext', fmt(v) + ' cm');
      if (document.activeElement !== inputs[k]) inputs[k].value = fmt(v);
    }
  }
  function set(k, v) {
    if (k === 'f') pos.f = v;
    else pos.n = Math.abs(v - tw()) < 1e-6 ? null : v;   // ganz am rechten Rand: bleibt am Rand, auch bei anderen Seiten
    clamp(); draw();
  }

  // Ziehen mit der Maus
  for (const k of ['f', 'n']) {
    const m = marks[k];
    m.addEventListener('pointerdown', e => {
      e.preventDefault(); m.focus(); m.setPointerCapture(e.pointerId);
      const move = ev => { const r = bar.getBoundingClientRect(); set(k, snap((ev.clientX - r.left) / (r.width / page.w) - page.l)); };
      const up = () => { m.removeEventListener('pointermove', move); m.removeEventListener('pointerup', up); m.removeEventListener('pointercancel', up); save(); };
      m.addEventListener('pointermove', move); m.addEventListener('pointerup', up); m.addEventListener('pointercancel', up);
    });
    m.addEventListener('keydown', e => {
      const d = e.key === 'ArrowLeft' || e.key === 'ArrowDown' ? -1 : e.key === 'ArrowRight' || e.key === 'ArrowUp' ? 1 : 0;
      if (!d) return;
      e.preventDefault();
      set(k, snap((k === 'f' ? pos.f : nVal()) + d * (e.shiftKey ? 1 : STEP))); save();
    });
    inputs[k].addEventListener('change', () => {
      const v = parseFloat(inputs[k].value.replace(',', '.'));
      if (isFinite(v)) set(k, v);
      save(); draw();
    });
  }
  if (window.ResizeObserver) new ResizeObserver(draw).observe(bar);

  // Seite aus Word lesen: Seitenbreite und Ränder des Abschnitts an der Cursorstelle (Twips → cm)
  async function readPage() {
    if (typeof Word === 'undefined') return;
    try {
      let xml = '';
      await Word.run(async ctx => {
        const o = ctx.document.getSelection().paragraphs.getFirst().getOoxml();
        await ctx.sync();
        xml = o.value || '';
      });
      const sz = xml.match(/<w:pgSz\b[^>]*>/), mar = xml.match(/<w:pgMar\b[^>]*>/);
      const att = (t, a) => { const m = t && t[0].match(new RegExp('w:' + a + '="(-?\\d+)"')); return m ? +m[1] / 567 : null; };
      const w = att(sz, 'w'), l = att(mar, 'left') ?? att(mar, 'start'), r = att(mar, 'right') ?? att(mar, 'end');
      if (w && l != null && r != null && w - l - r > 2) { page = {w, l, r}; clamp(); draw(); }
    } catch (e) {}
  }

  // Sichtbar in wissenschaftlicher Darstellung, bei neuen und bei geladenen Formeln
  let shown = false;
  setInterval(() => {
    const ed = FeWord.editing();
    const show = mode === 'sci';
    const on = ed || !!(q('#fe-num-on') && q('#fe-num-on').checked);   // geladene Formel: Nummer zählt, falls sie eine hat
    if (show !== shown) { shown = show; box.hidden = !show; if (show) { readPage(); requestAnimationFrame(draw); } }
    if (on !== numOn) { numOn = on; draw(); }
  }, 400);
  if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
    let t;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => { clearTimeout(t); t = setTimeout(() => { if (shown) readPage(); }, 900); });
  }
  draw();

  const pt = cm => Math.round(cm / 2.54 * 72 * 10) / 10;
  // Werte für Word in Punkt: Einzug links der Formel, Einzug rechts so, dass die Nummer bei „Nummer bis“ endet
  window.FeLineal = {
    visible: () => shown,
    use: () => use.checked,
    leftPt: () => pt(pos.f),
    tabPt: () => pt(nVal()),
    // Absatzeigenschaften für das OOXML einer neuen Formel (Twips): Einzug links, rechter Tabstopp für die Nummer
    pPr: () => '<w:pPr><w:tabs><w:tab w:val="right" w:pos="' + Math.round(pt(nVal()) * 20) + '"/></w:tabs>' +
      '<w:ind w:left="' + Math.round(pt(pos.f) * 20) + '" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>',
    text: () => 'Formel ab ' + fmt(pos.f) + ' cm, Nummer endet bei ' + fmt(nVal()) + ' cm',
    readPage,
  };
})();
