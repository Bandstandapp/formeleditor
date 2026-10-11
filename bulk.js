/* Kachel „Ganzes Dokument bearbeiten“:
   - Alle Formeln konvertieren: holt das Dokument, wandelt alte Formeln um (auf Wunsch auch Schrift und
     Größe aller Formeln auf den Standard), öffnet das Ergebnis als neues Dokument und zeigt eine Übersicht.
   - Analysieren: zählt Formeln nach Format, Schrift und Schriftgröße.
   - Suchen: markiert passende Formeln im geöffneten Dokument farbig, springt von Formel zu Formel und
     formatiert eine oder alle im Standard. Die Darstellung (wissenschaftlich/einzeilig) bleibt je Formel erhalten. */
(function () {
  const sec = document.createElement('section');
  sec.className = 'panel'; sec.setAttribute('aria-label', 'Dokument');
  sec.innerHTML = `
  <div class="fe-group">
    <p class="fe-sub">Standard für dieses Dokument</p>
    <div class="pickrow">
      <label class="pick">Schrift<select id="fe-std-font"></select></label>
      <label class="pick">Schriftgröße<select id="fe-std-size"></select></label>
    </div>
  </div>
  <div class="fe-group">
    <button class="btn primary" type="button" id="fe-all">Alle Formeln konvertieren</button>
    <label class="fe-check"><input type="checkbox" checked disabled> Neues Datenformat</label>
    <label class="fe-check"><input type="checkbox" id="fe-apply-font"> Schriftart übernehmen</label>
    <label class="fe-check"><input type="checkbox" id="fe-apply-size"> Schriftgröße übernehmen</label>
    <progress id="fe-prog" max="1" value="0" hidden></progress>
    <p class="hint" id="fe-allmsg" role="status"></p>
    <div id="fe-report" class="fe-box2" hidden></div>
  </div>
  <div class="fe-group">
    <button class="btn" type="button" id="fe-analyze">Formeln im Dokument analysieren</button>
    <p class="hint" id="fe-anamsg" role="status"></p>
    <div id="fe-analysis" class="fe-box2" hidden></div>
  </div>
  <div class="fe-group">
    <p class="fe-sub">Formeln suchen</p>
    <div class="fe-search">
      <label class="pick">Format<select id="fe-q-format"></select></label>
      <label class="pick">Schrift<select id="fe-q-font"></select></label>
      <label class="pick">Schriftgröße<select id="fe-q-size"></select></label>
    </div>
    <button class="btn" type="button" id="fe-find">Suchen und markieren</button>
    <p class="hint" id="fe-findmsg" role="status"></p>
    <div id="fe-nav" hidden>
      <div class="fe-navrow">
        <button class="btn" type="button" id="fe-prev">◀ Zurück</button>
        <span id="fe-pos"></span>
        <button class="btn" type="button" id="fe-next">Weiter ▶</button>
      </div>
      <button class="btn" type="button" id="fe-fmt-one">Diese Formel im Standard formatieren</button>
      <button class="btn" type="button" id="fe-fmt-all">Alle markierten im Standard formatieren</button>
      <button class="btn" type="button" id="fe-unmark">Markierungen entfernen</button>
    </div>
  </div>`;
  (document.querySelector('.wrap') || document.body).appendChild(sec);
  const tip = 'Konvertieren: wandelt alte Formeln ins neue Format um und öffnet das Ergebnis als neues Dokument; das Original bleibt unverändert. ' +
    'Mit den Häkchen bekommen alle Formeln die Standard-Schrift bzw. -Größe. Die Darstellung (wissenschaftlich oder einzeilig) bleibt bei jeder Formel, wie sie ist.\n\n' +
    'Analysieren: zählt die Formeln nach Format, Schrift und Größe. Ein Klick auf eine Zeile sucht genau diese Formeln.\n\n' +
    'Suchen: markiert passende Formeln im geöffneten Dokument türkis. Mit Zurück und Weiter springt man von Formel zu Formel und kann sie einzeln oder alle auf einmal im Standard formatieren.';
  if (window.FeInfo) window.FeInfo(sec.querySelector('.fe-sub'), tip);   // Überschrift steht schon auf dem Reiter

  const css = document.createElement('style');
  css.textContent = `
section[aria-label="Dokument"] .fe-group { display: grid; gap: 8px; padding-top: 10px; border-top: 1px solid var(--line); }
section[aria-label="Dokument"] .fe-group:first-of-type { border-top: 0; padding-top: 0; }
.fe-sub { margin: 0; font-weight: 700; font-size: .95rem; }
.fe-check { display: flex; align-items: center; gap: 8px; font-size: .95rem; }
.fe-check input { width: 18px; height: 18px; accent-color: var(--accent); }
.fe-check:has(input:disabled) { color: var(--muted); }
#fe-prog { width: 100%; }
.fe-box2 { border: 2px solid var(--accent); border-radius: var(--r); padding: 10px 12px; display: grid; gap: 8px; }
.fe-box2[hidden], #fe-nav[hidden] { display: none !important; }
.fe-box2 h3 { margin: 0; font-family: var(--f-head); font-size: var(--s-2); }
.fe-box2 table { border-collapse: collapse; width: 100%; }
.fe-box2 td { padding: 3px 0; vertical-align: top; }
.fe-box2 td.n { text-align: right; font-weight: 700; padding-left: 10px; white-space: nowrap; }
.fe-box2 tr.sum td { border-top: 1px solid var(--line); padding-top: 6px; }
.fe-box2 tr.grp td { padding-top: 8px; font-weight: 700; }
.fe-box2 td.sub { padding-left: 12px; }
.fe-box2 mark { background: #fff200; color: #111; padding: 0 3px; }
.fe-box2 ul { margin: 0; padding-left: 18px; }
.fe-box2 .linkbtn { text-align: left; }
.fe-search { display: grid; grid-template-columns: minmax(0, 1fr); gap: 6px; }
#fe-nav { display: grid; gap: 6px; }
.fe-navrow { display: grid; grid-template-columns: auto 1fr auto; gap: 6px; align-items: center; text-align: center; font-weight: 700; }
section[aria-label="Dokument"] .btn:not(.primary) { padding: 8px 10px; font-size: .95rem; }`;
  document.head.appendChild(css);

  const $ = id => document.getElementById(id);
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const ptText = p => String(p).replace('.', ',') + ' pt';
  const SIZES = [10, 11, 12, 14, 16, 20, 24];
  const inWord = () => typeof Office !== 'undefined' && Office.context && Office.context.document;
  const MARK = 'Turquoise';

  // Standard-Schrift und -Größe (gemerkt)
  const stdFont = $('fe-std-font'), stdSize = $('fe-std-size');
  FONTS.forEach(f => stdFont.add(new Option(f.name, f.id)));
  SIZES.forEach(p => stdSize.add(new Option(ptText(p), p)));
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem('fe-std') || '{}'); } catch (e) {}
  stdFont.value = FONTS.some(f => f.id === saved.font) ? saved.font : (typeof font !== 'undefined' ? font : 'stix2');
  stdSize.value = SIZES.includes(+saved.pt) ? saved.pt : (SIZES.includes(sizePt) ? sizePt : 12);
  const saveStd = () => { try { localStorage.setItem('fe-std', JSON.stringify({font: stdFont.value, pt: +stdSize.value})); } catch (e) {} };
  stdFont.addEventListener('change', saveStd); stdSize.addEventListener('change', saveStd);

  // Suchfelder
  const qFormat = $('fe-q-format'), qFont = $('fe-q-font'), qSize = $('fe-q-size');
  function fillSearch(an) {
    const cnt = (obj, k) => (an && obj[k] ? ' (' + obj[k] + ')' : '');
    const keepF = qFormat.value, keepFo = qFont.value, keepS = qSize.value;
    qFormat.innerHTML = ''; qFont.innerHTML = ''; qSize.innerHTML = '';
    qFormat.add(new Option('alle Formate', ''));
    const formats = an ? Object.keys(an.formats) : ['Neues Format (Formeleditor)'];
    formats.forEach(k => qFormat.add(new Option(k + cnt(an ? an.formats : {}, k), k)));
    qFont.add(new Option('alle Schriften', ''));
    FONTS.forEach(f => qFont.add(new Option(f.name + cnt(an ? an.fonts : {}, f.name), f.name)));
    if (an) Object.keys(an.fonts).filter(k => !FONTS.some(f => f.name === k)).forEach(k => qFont.add(new Option(k + cnt(an.fonts, k), k)));
    qSize.add(new Option('alle Größen', ''));
    const sizes = Array.from(new Set(SIZES.concat(an ? Object.keys(an.sizes).map(Number) : []))).sort((a, b) => a - b);
    sizes.forEach(p => qSize.add(new Option(ptText(p) + cnt(an ? an.sizes : {}, p), p)));
    qFormat.value = keepF; qFont.value = keepFo; qSize.value = keepS;
  }
  fillSearch(null);

  // ---------- Dokument lesen und Formeln setzen ----------
  let jszip = null;
  function loadJSZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    if (jszip) return jszip;
    return (jszip = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = new URL('vendor/jszip/jszip.min.js', location.href).href;
      s.onload = () => res(window.JSZip);
      s.onerror = () => { jszip = null; rej(new Error('Hilfsprogramm zum Entpacken lädt nicht. Bitte Internet prüfen.')); };
      document.head.appendChild(s);
    }));
  }
  function documentBytes() {
    return new Promise((res, rej) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, {sliceSize: 4194304}, r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return rej(new Error(r.error.message));
        const file = r.value, chunks = [];
        const next = i => file.getSliceAsync(i, s => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return rej(new Error(s.error.message)); }
          chunks.push(s.value.data);
          if (i + 1 < file.sliceCount) return next(i + 1);
          file.closeAsync();
          const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
          let pos = 0;
          chunks.forEach(c => { out.set(c, pos); pos += c.length; });
          res(out);
        });
        next(0);
      });
    });
  }
  async function documentZip() {
    const [JSZip, bytes] = await Promise.all([loadJSZip(), documentBytes()]);
    return JSZip.loadAsync(bytes);
  }

  // Je Schrift ein unsichtbarer Satz-Rahmen, damit die Vorschau oben unberührt bleibt
  const frames = {};
  function fontFrame(id) {
    if (!frames[id]) {
      frames[id] = new Promise(res => {
        const f = document.createElement('iframe');
        f.setAttribute('aria-hidden', 'true'); f.tabIndex = -1;
        f.style.cssText = 'position:absolute;left:-10000px;top:0;width:700px;height:400px;border:0;';
        f.onload = async () => { const w = f.contentWindow; for (let i = 0; i < 300 && !w.__mjReady; i++) await sleep(100); res(w); };
        f.srcdoc = frameDoc(id);
        document.body.appendChild(f);
      });
    }
    return frames[id];
  }
  async function renderFormula(tex, m, pt, pad, fontId, dpi) {
    const w = await fontFrame(fontId);
    const disp = render(tex, m, true);
    w.__show(disp, disp, m);
    return w.__png(pt, dpi || 300, pad);
  }

  // ---------- Konvertieren ----------
  const btn = $('fe-all'), prog = $('fe-prog'), msg = $('fe-allmsg'), box = $('fe-report');
  function showReport(rep) {
    const fmts = Object.entries(rep.formats);
    let h = '<h3>Ergebnis</h3><table>';
    h += '<tr><td>Gefundene Formeln</td><td class="n">' + rep.total + '</td></tr>';
    h += '<tr><td>Unterschiedliche Ausgangsformate</td><td class="n">' + fmts.length + '</td></tr>';
    fmts.forEach(([k, v]) => { h += '<tr><td class="sub">' + esc(k) + '</td><td class="n">' + v + '</td></tr>'; });
    h += '<tr class="sum"><td>Erfolgreich konvertiert</td><td class="n">' + rep.converted + '</td></tr>';
    if (rep.reformatted) h += '<tr><td>Neu formatiert</td><td class="n">' + rep.reformatted + '</td></tr>';
    h += '<tr><td>Mit Problemen, <mark>gelb markiert</mark></td><td class="n">' + rep.problems.length + '</td></tr>';
    if (rep.unchanged) h += '<tr><td>Unverändert gelassen</td><td class="n">' + rep.unchanged + '</td></tr>';
    h += '</table>';
    if (rep.problems.length) {
      h += '<p class="hint">Diese Formeln bitte im neuen Dokument neu eingeben:</p><ul>';
      rep.problems.forEach(p => { h += '<li>Formel ' + p.number + ' (' + esc(p.where) + '): ' + esc(p.why) + '</li>'; });
      h += '</ul>';
    }
    box.innerHTML = h; box.hidden = false;
  }
  async function convertAll() {
    if (!inWord()) { msg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    btn.disabled = true; box.hidden = true; prog.hidden = false; prog.value = 0;
    try {
      msg.textContent = 'Dokument wird gelesen …';
      const zip = await documentZip();
      const rep = await FeConvert.convertDocx(zip, renderFormula, {
        font: stdFont.value, pt: +stdSize.value,
        applyFont: $('fe-apply-font').checked, applyPt: $('fe-apply-size').checked,
        onProgress: (k, n) => { prog.max = n; prog.value = k; msg.textContent = 'Formel ' + k + ' von ' + n + ' …'; },
      });
      if (!rep.total) { msg.textContent = 'In diesem Dokument wurden keine Formeln gefunden.'; return; }
      if (!rep.changed) { msg.textContent = 'Es gab nichts zu konvertieren, deshalb wurde kein neues Dokument erstellt.'; showReport(rep); return; }
      msg.textContent = 'Neues Dokument wird geöffnet …';
      const b64 = await zip.generateAsync({type: 'base64', compression: 'DEFLATE'});
      await Word.run(async ctx => { const d = ctx.application.createDocument(b64); await ctx.sync(); d.open(); await ctx.sync(); });
      msg.textContent = 'Fertig. Das Ergebnis ist in einem neuen Fenster geöffnet. Bitte dort mit „Ablage“ → „Speichern unter …“ sichern.';
      showReport(rep);
    } catch (e) {
      msg.textContent = 'Das Konvertieren hat nicht geklappt: ' + (e && e.message ? e.message : e);
    } finally { btn.disabled = false; prog.hidden = true; }
  }
  btn.addEventListener('click', convertAll);

  // ---------- Analysieren ----------
  const anaBtn = $('fe-analyze'), anaMsg = $('fe-anamsg'), anaBox = $('fe-analysis');
  async function analyze() {
    if (!inWord()) { anaMsg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    anaBtn.disabled = true; anaMsg.textContent = 'Dokument wird gelesen …';
    try {
      const an = await FeConvert.analyze(await documentZip());
      anaMsg.textContent = '';
      fillSearch(an);
      const row = (label, n, kind, val) => '<tr><td class="sub"><button type="button" class="linkbtn" data-kind="' + kind + '" data-val="' + esc(val) + '">' + esc(label) + '</button></td><td class="n">' + n + '</td></tr>';
      let h = '<h3>Formeln im Dokument</h3><table><tr><td>Formeln insgesamt</td><td class="n">' + an.total + '</td></tr>';
      h += '<tr class="grp"><td colspan="2">Formate</td></tr>';
      Object.entries(an.formats).forEach(([k, v]) => { h += row(k, v, 'format', k); });
      h += '<tr class="grp"><td colspan="2">Schriften</td></tr>';
      Object.entries(an.fonts).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => { h += row(k, v, 'font', k); });
      h += '<tr class="grp"><td colspan="2">Schriftgrößen</td></tr>';
      Object.entries(an.sizes).sort((a, b) => a[0] - b[0]).forEach(([k, v]) => { h += row(ptText(k), v, 'size', k); });
      h += '</table><p class="hint">Ein Klick auf eine Zeile sucht und markiert genau diese Formeln.</p>';
      anaBox.innerHTML = h; anaBox.hidden = false;
      anaBox.querySelectorAll('button[data-kind]').forEach(b => b.addEventListener('click', () => {
        qFormat.value = ''; qFont.value = ''; qSize.value = '';
        ({format: qFormat, font: qFont, size: qSize})[b.dataset.kind].value = b.dataset.val;
        find();
      }));
    } catch (e) {
      anaMsg.textContent = 'Die Analyse hat nicht geklappt: ' + (e && e.message ? e.message : e);
    } finally { anaBtn.disabled = false; }
  }
  anaBtn.addEventListener('click', analyze);

  // ---------- Suchen, markieren, springen, formatieren ----------
  const findMsg = $('fe-findmsg'), nav = $('fe-nav'), posEl = $('fe-pos');
  let hits = [], pos = 0;
  const NEWFMT = 'Neues Format (Formeleditor)';
  const nameOf = id => (FONTS.find(f => f.id === id) || {}).name || id;

  async function pictures(ctx) {
    const pics = ctx.document.body.inlinePictures;
    pics.load('items/altTextDescription');
    await ctx.sync();
    return pics.items.map((p, i) => ({p, i, meta: findMeta(null, p.altTextDescription || '')}));
  }
  function matches(meta) {
    if (!meta) return false;
    if (qFont.value && nameOf(meta.font) !== qFont.value) return false;
    if (qSize.value && +meta.pt !== +qSize.value) return false;
    return true;
  }
  function showPos() {
    nav.hidden = !hits.length;
    posEl.textContent = hits.length ? 'Formel ' + (pos + 1) + ' von ' + hits.length : '';
  }
  async function setMark(list, color) {
    await Word.run(async ctx => {
      const all = await pictures(ctx);
      list.forEach(h => { const it = all[h.i]; if (it && it.meta && it.meta.tex === h.tex) it.p.getRange('Whole').font.highlightColor = color; });
      await ctx.sync();
    });
  }
  async function find() {
    if (!inWord()) { findMsg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    if (qFormat.value && qFormat.value !== NEWFMT) {
      findMsg.textContent = 'Formeln im Format „' + qFormat.value + '“ lassen sich im Dokument nicht einzeln markieren. Bitte zuerst „Alle Formeln konvertieren“ verwenden.';
      return;
    }
    try {
      if (hits.length) await setMark(hits, null);
      findMsg.textContent = 'Suche läuft …';
      await Word.run(async ctx => {
        const all = await pictures(ctx);
        hits = all.filter(x => matches(x.meta)).map(x => ({i: x.i, tex: x.meta.tex}));
        hits.forEach(h => { all[h.i].p.getRange('Whole').font.highlightColor = MARK; });
        await ctx.sync();
      });
      pos = 0;
      findMsg.textContent = hits.length ? hits.length + ' Formeln gefunden und türkis markiert.' : 'Keine passenden Formeln gefunden.';
      showPos();
      if (hits.length) await goTo(0);
    } catch (e) { findMsg.textContent = 'Die Suche hat nicht geklappt: ' + (e && e.message ? e.message : e); }
  }
  async function current(ctx) {
    const all = await pictures(ctx);
    const h = hits[pos], it = h && all[h.i];
    if (!it || !it.meta || it.meta.tex !== h.tex) throw new Error('Das Dokument hat sich verändert. Bitte noch einmal suchen.');
    return it;
  }
  async function goTo(n) {
    if (!hits.length) return;
    pos = (n + hits.length) % hits.length;
    showPos();
    try { await Word.run(async ctx => { (await current(ctx)).p.select(); await ctx.sync(); }); }
    catch (e) { findMsg.textContent = e.message; }
  }
  async function formatAt(ctx, it) {
    const m = it.meta;
    const meta = {tex: m.tex, mode: ['lin', 'kom'].includes(m.mode) ? m.mode : 'sci', font: stdFont.value, pt: +stdSize.value};
    if (m.pad) meta.pad = m.pad;
    const r = await renderFormula(meta.tex, meta.mode, meta.pt, meta.pad || 0, meta.font, 600);
    it.p.getRange('Whole').insertOoxml(FeWord.ooxml(r, 'FORMEL:' + JSON.stringify(meta)), 'Replace');
    await ctx.sync();
  }
  async function formatOne() {
    try {
      await Word.run(async ctx => { await formatAt(ctx, await current(ctx)); });
      hits.splice(pos, 1);
      findMsg.textContent = 'Formel formatiert. Noch ' + hits.length + ' markiert.';
      if (pos >= hits.length) pos = 0;
      showPos();
      if (hits.length) await goTo(pos);
    } catch (e) { findMsg.textContent = 'Das Formatieren hat nicht geklappt: ' + (e && e.message ? e.message : e); }
  }
  async function formatAll() {
    const total = hits.length;
    nav.querySelectorAll('button').forEach(b => { b.disabled = true; });
    try {
      for (let k = 0; hits.length; k++) {
        pos = 0;
        findMsg.textContent = 'Formel ' + (k + 1) + ' von ' + total + ' wird formatiert …';
        await Word.run(async ctx => { await formatAt(ctx, await current(ctx)); });
        hits.shift();
      }
      findMsg.textContent = (total === 1 ? 'Die Formel ist' : 'Alle ' + total + ' Formeln sind') + ' im Standard formatiert (' + nameOf(stdFont.value) + ', ' + ptText(stdSize.value) + ').';
    } catch (e) {
      findMsg.textContent = 'Abgebrochen: ' + (e && e.message ? e.message : e);
    } finally {
      nav.querySelectorAll('button').forEach(b => { b.disabled = false; });
      showPos();
    }
  }
  async function unmark() {
    try { await setMark(hits, null); } catch (e) {}
    hits = []; showPos();
    findMsg.textContent = 'Markierungen entfernt.';
  }
  $('fe-find').addEventListener('click', find);
  $('fe-prev').addEventListener('click', () => goTo(pos - 1));
  $('fe-next').addEventListener('click', () => goTo(pos + 1));
  $('fe-fmt-one').addEventListener('click', formatOne);
  $('fe-fmt-all').addEventListener('click', formatAll);
  $('fe-unmark').addEventListener('click', unmark);
})();
