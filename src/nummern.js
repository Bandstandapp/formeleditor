/* Gleichungsnummern wie (4.13), so wie sie in Papas Manuskripten stehen: als Text am Ende des Absatzes
   mit der Formel. Neue Formeln bekommen auf Wunsch eine Nummer rechtsbündig am Rand. Liegt die neue
   Nummer mitten in der Zählung, werden die folgenden Nummern im Kapitel weitergezählt, bei Nummern mit
   Punkt (4.13) auch die Verweise im Text. In „Alles bearbeiten“ prüft ein Knopf die Zählung.
   Die reine Logik steht in FeNum (auch für Tests ohne Word), die Word-Anbindung darunter. */
(function (root) {
  // ---------- Logik ----------
  const END = /\((\d+(?:\.\d+)*)([a-z]?)\)\s*$/;          // Nummer am Absatzende
  const REF = /\((\d+(?:\.\d+)+)([a-z]?)\)/g;               // Verweis mit Punkt im Text, etwa (4.13)
  const VALID = /^\d+(?:\.\d+)*[a-z]?$/;
  function split(label) {
    const m = String(label).match(/^((?:\d+\.)*)(\d+)([a-z]?)$/);
    return m ? {prefix: m[1].replace(/\.$/, ''), n: +m[2], letter: m[3]} : null;
  }
  const join = (prefix, n, letter) => (prefix ? prefix + '.' : '') + n + (letter || '');
  // Nummer am Ende eines Absatztextes, sonst null
  // Nur in kurzen Absätzen: Ein langer Textabsatz, der mit einem Verweis endet, ist keine Formelzeile
  function endNumber(text) {
    const t = String(text || '').replace(/[\u0000-\u001f\ufffc]+$/g, '');
    const m = t.match(END);
    return m && t.slice(0, m.index).replace(/[\s\ufffc]/g, '').length <= 40 ? m[1] + m[2] : null;
  }
  // Vorschlag für die nächste Nummer aus den Nummern vor dem Cursor (und notfalls danach)
  function suggest(before, after) {
    const last = before.length ? split(before[before.length - 1]) : null;
    if (last) return join(last.prefix, last.n + 1);
    const first = after.length ? split(after[0]) : null;
    if (first) return first.prefix ? first.prefix.split('.').map(() => '1').join('.') + '.1' : '1';
    return '1';
  }
  // Welche vorhandenen Nummern müssen um eins weiter, wenn „num“ neu dazukommt?
  // Liefert absteigend sortierte Paare [alt, neu]; nur Nummern im selben Kapitel ab num.
  function shifts(existing, num) {
    const s = split(num);
    if (!s || s.letter) return [];
    const out = new Map();
    existing.forEach(l => {
      const x = split(l);
      if (x && x.prefix === s.prefix && x.n >= s.n) out.set(l, join(x.prefix, x.n + 1, x.letter));
    });
    if (![...out.keys()].some(l => split(l).n === s.n)) return [];   // Nummer ist frei, nichts zu tun
    return [...out.entries()].sort((a, b) => { const x = split(a[0]), y = split(b[0]); return y.n - x.n || (y.letter > x.letter ? 1 : -1); });
  }
  // Prüfbericht: je Kapitel Anzahl, Lücken, doppelte Nummern, Verweise ohne Gleichung
  function check(paraTexts) {
    const eqs = [], refs = [];
    paraTexts.forEach(t => {
      const e = endNumber(t);
      let body = String(t || '');
      if (e) { eqs.push(e); body = body.replace(/[\u0000-\u001f\ufffc]+$/g, '').replace(END, ''); }
      for (const m of body.matchAll(REF)) refs.push(m[1] + m[2]);
    });
    const groups = new Map();
    eqs.forEach(l => { const x = split(l); const k = x.prefix; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(x); });
    const chapters = [...groups.entries()].map(([prefix, xs]) => {
      const seen = new Map(); xs.forEach(x => { const k = x.n + x.letter; seen.set(k, (seen.get(k) || 0) + 1); });
      const dup = [...seen.entries()].filter(([, c]) => c > 1).map(([k]) => join(prefix, k));
      const ns = [...new Set(xs.map(x => x.n))].sort((a, b) => a - b);
      const gaps = [];
      for (let n = 1; n <= ns[ns.length - 1]; n++) if (!ns.includes(n)) gaps.push(join(prefix, n));
      const order = [];                          // Stellen, an denen die Zählung rückwärts springt
      for (let i = 1; i < xs.length; i++) if (xs[i].n < xs[i - 1].n) order.push(join(prefix, xs[i].n, xs[i].letter));
      return {prefix, count: xs.length, first: join(prefix, xs[0].n, xs[0].letter), last: join(prefix, xs[xs.length - 1].n, xs[xs.length - 1].letter), dup, gaps, order};
    });
    const known = new Set(eqs);
    const missing = [...new Set(refs.filter(r => !known.has(r)))];
    return {total: eqs.length, refs: refs.length, chapters, missing};
  }
  const FeNum = {endNumber, suggest, shifts, check, split, VALID};
  if (typeof module !== 'undefined' && module.exports) { module.exports = FeNum; return; }
  root.FeNum = FeNum;

  // ---------- Word-Anbindung ----------
  const q = s => document.querySelector(s);
  const vorschau = q('section[aria-label="Ergebnis"]');
  if (!vorschau || !window.FeWord) return;
  const css = document.createElement('style');
  css.textContent = `
.fe-num { display: grid; gap: 6px; }
.fe-num[hidden], .fe-num .fe-num-more[hidden] { display: none !important; }
.fe-num .fe-num-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.fe-num input[type=text] { font: inherit; font-size: 1rem; width: 6.5em; padding: 5px 8px; border: 2px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--fg); }
.fe-num input[type=text]:focus-visible { outline: 3px solid var(--focus); }
.fe-num .fe-check { font-size: .95rem; }
.fe-num #fe-num-why { margin: 0; }
.fe-num #fe-num-why:empty, .fe-num #fe-num-msg:empty { display: none; }`;
  document.head.appendChild(css);
  const box = document.createElement('div');
  box.className = 'fe-num';
  box.innerHTML = '<label class="fe-check"><input type="checkbox" id="fe-num-on"> Gleichungsnummer erzeugen</label>' +
    '<p class="hint" id="fe-num-why"></p>' +
    '<div class="fe-num-more" hidden><div class="fe-num-row">(<input type="text" id="fe-num" inputmode="decimal" autocomplete="off" aria-label="Gleichungsnummer">)' +
    '<button type="button" class="linkbtn" id="fe-num-sug">Vorschlag</button></div>' +
    '<label class="fe-check"><input type="checkbox" id="fe-num-shift" checked> Folgende Nummern weiterzählen, auch in Verweisen</label></div>' +
    '<p class="hint" id="fe-num-msg" role="status"></p>';
  const modeBox = q('#mode');
  (modeBox && modeBox.parentElement === vorschau ? modeBox : vorschau.querySelector('.paper')).after(box);
  const on = box.querySelector('#fe-num-on'), field = box.querySelector('#fe-num'), shiftBox = box.querySelector('#fe-num-shift');
  const more = box.querySelector('.fe-num-more'), msg = box.querySelector('#fe-num-msg'), why = box.querySelector('#fe-num-why');
  try { on.checked = localStorage.getItem('fe-num-on') === '1'; shiftBox.checked = localStorage.getItem('fe-num-shift') !== '0'; } catch (e) {}
  // Einzüge aus dem Lineal (lineal.js); ohne Lineal am Rand
  const L = () => window.FeLineal || {leftPt: () => 0, rightPt: () => 0, text: () => '', use: () => false, visible: () => false};
  const useRuler = () => L().visible() && L().use();
  function indent(para) {
    para.leftIndent = L().leftPt();
    para.firstLineIndent = 0;
    para.rightIndent = L().rightPt();
    para.alignment = 'Left';
  }
  const inWord = () => typeof Word !== 'undefined' && typeof Office !== 'undefined' && Office.context && Office.context.document;

  // Nummer in der Vorschau anzeigen
  function preview() {
    try {
      const d = q('#out').contentDocument, no = d && d.querySelector('.eq .no');
      if (no) no.textContent = !on.disabled && on.checked && field.value.trim() ? '(' + field.value.trim() + ')' : '';
    } catch (e) {}
  }
  // Immer zu sehen; wirkt nur bei neuen abgesetzten Formeln (wissenschaftlich, nicht beim Ersetzen)
  function sync() {
    const w = mode !== 'sci' ? 'Gleichungsnummern gibt es nur bei der Darstellung „Wissenschaftlich“.'
      : FeWord.editing() ? 'Beim Ersetzen bleibt die Nummer der Formel, wie sie ist.' : '';
    if (why.textContent !== w) why.textContent = w;
    on.disabled = !!w;
    more.hidden = !on.checked || !!w;
    preview();
  }
  setInterval(sync, 400);

  async function paragraphsAround(ctx) {
    const sel = ctx.document.getSelection();
    const body = ctx.document.body;
    const before = body.getRange('Start').expandTo(sel.getRange('Start')).paragraphs;
    const all = body.paragraphs;
    before.load('items/text'); all.load('items/text');
    await ctx.sync();
    const nb = before.items.length;
    const nums = all.items.map(p => endNumber(p.text));
    // Der Absatz mit dem Cursor zählt nur mit, wenn seine Nummer vor dem Cursor steht
    const lastBefore = before.items.length ? endNumber(before.items[nb - 1].text) : null;
    const cut = lastBefore ? nb : nb - 1;
    return {before: nums.slice(0, Math.max(0, cut)).filter(Boolean), after: nums.slice(Math.max(0, cut)).filter(Boolean), all: nums.filter(Boolean), items: all.items};
  }
  async function makeSuggestion() {
    if (!inWord()) return;
    try {
      await Word.run(async ctx => {
        const p = await paragraphsAround(ctx);
        field.value = suggest(p.before, p.after);
      });
      msg.textContent = '';
    } catch (e) { msg.textContent = 'Kein Vorschlag möglich: ' + (e && e.message ? e.message : e); }
    preview();
  }
  on.addEventListener('change', () => {
    try { localStorage.setItem('fe-num-on', on.checked ? '1' : '0'); } catch (e) {}
    sync();
    if (on.checked && !field.value.trim()) makeSuggestion();
  });
  shiftBox.addEventListener('change', () => { try { localStorage.setItem('fe-num-shift', shiftBox.checked ? '1' : '0'); } catch (e) {} });
  box.querySelector('#fe-num-sug').addEventListener('click', makeSuggestion);
  field.addEventListener('input', () => { msg.textContent = ''; preview(); });
  // Nach dem Anklicken einer Stelle im Dokument passt der Vorschlag nicht mehr: neu vorschlagen,
  // solange der Vater die Nummer nicht selbst geändert hat
  let auto = true;
  field.addEventListener('input', () => { auto = !field.value.trim(); });
  if (inWord() && Office.context.document.addHandlerAsync) {
    let t;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      clearTimeout(t);
      t = setTimeout(() => { if (on.checked && !on.disabled && auto) makeSuggestion(); }, 600);
    });
  }

  let pendingNote = '';
  // Gewünschte Nummer für die nächste eingefügte Formel, sonst null
  function wanted() {
    sync();
    if (on.disabled || !on.checked) return null;
    const v = field.value.trim().replace(/^\(|\)$/g, '');
    return VALID.test(v) ? v : null;
  }
  // Vor dem Einfügen: folgende Nummern weiterzählen, eigenen Absatz für die Formel anlegen
  async function beforeInsert(ctx, num) {
    pendingNote = '';
    const p = await paragraphsAround(ctx);
    const plan = shiftBox.checked ? shifts(p.all, num) : [];
    let refCount = 0;
    if (!shiftBox.checked && p.all.includes(num)) pendingNote = 'Achtung: Die Nummer (' + num + ') gibt es schon.';
    for (const [from, to] of plan) {
      if (from.includes('.')) {
        // Nummern mit Punkt: überall im Text, also Gleichung und Verweise
        const hits = ctx.document.body.search('(' + from + ')', {matchCase: true});
        hits.load('items');
        await ctx.sync();
        hits.items.forEach(h => h.insertText('(' + to + ')', 'Replace'));
        refCount += hits.items.length;
      } else {
        // Nummern ohne Punkt wie (3) nur am Ende von Formel-Absätzen, sonst träfe es auch Aufzählungen
        for (const para of p.items) {
          if (endNumber(para.text) !== from) continue;
          const hits = para.search('(' + from + ')', {matchCase: true});
          hits.load('items');
          await ctx.sync();
          if (hits.items.length) { hits.items[hits.items.length - 1].insertText('(' + to + ')', 'Replace'); refCount++; }
        }
      }
    }
    await ctx.sync();
    if (plan.length) {
      const refs = refCount - plan.length;
      pendingNote = plan.length + (plan.length === 1 ? ' folgende Nummer' : ' folgende Nummern') + ' weitergezählt' +
        (refs > 0 ? ', dazu ' + refs + (refs === 1 ? ' Verweis' : ' Verweise') + ' im Text' : '') + '.';
    }
    // Die nummerierte Formel bekommt einen eigenen Absatz
    const sel = ctx.document.getSelection();
    const para = sel.paragraphs.getFirst();
    para.load('text');
    await ctx.sync();
    if (para.text.replace(/[\s\ufffc]/g, '')) {
      const np = para.insertParagraph('', 'After');
      if (useRuler()) indent(np);
      return np.getRange('Start');
    }
    if (useRuler()) indent(para);
    return sel;
  }
  // Ohne Nummer: Steht die neue Formel allein in ihrem Absatz? Dann bekommt sie die Einzüge aus dem Lineal
  // (eine Formel im Fließtext nicht). Liefert die Absatzeigenschaften für das OOXML oder ''.
  async function placeUnnumbered(ctx) {
    if (!useRuler()) return '';
    const para = ctx.document.getSelection().paragraphs.getFirst();
    para.load('text');
    await ctx.sync();
    if (para.text.replace(/[\s\ufffc]/g, '')) return '';
    indent(para);
    return L().pPr();
  }
  // Absatzeigenschaften für eine nummerierte Formel (eigener Absatz)
  const pPr = () => useRuler() ? L().pPr() : '';
  // Nach dem Einfügen: Nummer für die nächste Formel vorschlagen
  function done(num) {
    const s = split(num);
    if (s && !s.letter) field.value = join(s.prefix, s.n + 1);
    auto = true;
    preview();
    const note = pendingNote; pendingNote = '';
    return note;
  }
  // Nach dem Einfügen einer abgesetzten Formel: Einzüge wie im Lineal. Mit Nummer hat die Formel ihren
  // eigenen Absatz; ohne Nummer nur, wenn der Absatz außer der Formel nichts enthält (keine Formel im Fließtext).
  async function afterInsert(ctx, range, numbered) {
    if (!useRuler()) return;
    const para = range.paragraphs.getFirst();
    if (!numbered) {
      para.load('text');
      await ctx.sync();
      if (para.text.replace(/[\s\ufffc]/g, '')) return;
    }
    indent(para);
  }
  root.FeNummer = {wanted, beforeInsert, placeUnnumbered, pPr, afterInsert, afterReplace, done};

  // Formeln mit Nummer wie im Lineal setzen: Einzug links für die Formel, Einzug rechts für das Ende der Nummer.
  // Leerraum (Tabs, Leerzeichen) zwischen Formel und Nummer wird durch einen Tabulator am rechten Einzug ersetzt.
  const PTAB = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:ptab w:relativeTo="indent" w:alignment="right" w:leader="none"/></w:r></w:p></w:body></w:document>' +
    '</pkg:xmlData></pkg:part></pkg:package>';
  const blank = t => /^[\s\u00a0]*$/.test(t || '');
  // Einen Absatz mit abgesetzter Formel wie im Lineal setzen; num: seine Nummer oder null. false, wenn die Nummer fehlt.
  async function alignPara(ctx, p, num) {
    const hits = num ? p.search('(' + num + ')', {matchCase: true}) : null;
    const pics = p.inlinePictures;
    const tabs = p.search('^t');
    if (hits) hits.load('items');
    pics.load('items'); tabs.load('items');
    await ctx.sync();
    if (num && !hits.items.length) return false;
    if (num) {
      const at = hits.items[hits.items.length - 1].getRange('Start');
      // Am liebsten alles zwischen der letzten Formel und der Nummer, wenn dort nur Leerraum steht
      let gap = null;
      if (pics.items.length) {
        const g = pics.items[pics.items.length - 1].getRange('End').expandTo(at);
        g.load('text');
        await ctx.sync();
        if (blank(g.text)) gap = g;
      }
      // Sonst die Tabs unmittelbar vor der Nummer
      if (!gap) {
        const before = p.getRange('Start').expandTo(at);
        before.load('text');
        await ctx.sync();
        const trail = (before.text.match(/[\t ]*$/) || [''])[0];
        const k = (trail.match(/\t/g) || []).length;
        if (k && tabs.items.length >= k) gap = tabs.items[tabs.items.length - k].expandTo(at);
      }
      if (gap) gap.insertOoxml(PTAB, 'Replace');
      else at.insertOoxml(PTAB, 'Replace');
    }
    // Leerraum vor der ersten Formel entfernen, damit sie beim Einzug aus dem Lineal beginnt
    if (pics.items.length) {
      const lead = p.getRange('Start').expandTo(pics.items[0].getRange('Start'));
      lead.load('text');
      await ctx.sync();
      if (lead.text && blank(lead.text)) lead.delete();
    }
    indent(p);
    return true;
  }
  // Alle abgesetzten Formeln: Absätze mit Nummer am Ende und Absätze, in denen nur eine Formel steht.
  // progress(i, n) meldet den Fortschritt.
  async function alignAll(progress) {
    let done = 0, numbered = 0, skipped = [];
    await Word.run(async ctx => {
      const ps = ctx.document.body.paragraphs;
      ps.load('items/text');
      await ctx.sync();
      // Absätze ohne Text: nachsehen, ob eine Formel aus diesem Editor darin steht
      const empty = ps.items.filter(p => !endNumber(p.text) && blank(String(p.text || '').replace(/[\u0000-\u001f\ufffc]/g, '')));
      empty.forEach(p => p.inlinePictures.load('items/altTextDescription,items/altTextTitle'));
      if (empty.length) await ctx.sync();
      const isFormula = p => p.inlinePictures.items.some(x => /FORMEL:/.test((x.altTextDescription || '') + (x.altTextTitle || '')));
      const todo = ps.items.map(p => ({p, num: endNumber(p.text)})).filter(x => x.num || (empty.includes(x.p) && isFormula(x.p)));
      if (progress) progress(0, todo.length);
      for (let i = 0; i < todo.length; i++) {
        const {p, num} = todo[i];
        if (await alignPara(ctx, p, num)) { done++; if (num) numbered++; } else skipped.push(num);
        if (progress) progress(i + 1, todo.length);
      }
      await ctx.sync();
    });
    return {done, numbered, skipped};
  }
  // Nach dem Ersetzen einer geladenen Formel: wenn „Werte für Formel übernehmen“ an ist und die Formel
  // allein (mit oder ohne Nummer) in ihrem Absatz steht
  async function afterReplace(ctx, range) {
    if (!useRuler()) return;
    const p = range.paragraphs.getFirst();
    p.load('text');
    await ctx.sync();
    const num = endNumber(p.text);
    const rest = num ? p.text.replace(/[\u0000-\u001f\ufffc]+$/g, '').replace(END, '') : p.text;
    if (rest.replace(/[\s\ufffc]/g, '')) return;
    await alignPara(ctx, p, num);
  }
  // Knopf unter dem Lineal (lineal.js): alle Formeln ausrichten, mit Fortschrittsbalken
  const alignBtn = q('#fe-align'), bar = q('#fe-align-bar'), amsg = q('#fe-align-msg');
  if (alignBtn) alignBtn.addEventListener('click', async () => {
    if (!inWord()) { amsg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    alignBtn.disabled = true;
    amsg.textContent = 'Formeln werden ausgerichtet …';
    bar.value = 0; bar.max = 1; bar.hidden = false;
    try {
      const r = await alignAll((i, n) => {
        bar.max = Math.max(1, n); bar.value = i;
        amsg.textContent = 'Formeln werden ausgerichtet: ' + i + ' von ' + n + ' …';
      });
      const f = n => n === 1 ? '1 Formel' : n + ' Formeln';
      amsg.textContent = r.done || r.skipped.length
        ? 'Ausrichtung abgeschlossen. ' + f(r.done) + ' ausgerichtet' + (r.numbered ? ', davon ' + r.numbered + ' mit Nummer' : '') + ' (' + L().text() + ').' +
          (r.skipped.length ? ' Nicht gefunden: ' + r.skipped.map(x => '(' + x + ')').join(', ') + '.' : '')
        : 'Ausrichtung abgeschlossen. Im Dokument steht keine abgesetzte Formel.';
    } catch (e) { amsg.textContent = 'Das Ausrichten hat nicht geklappt: ' + (e && e.message ? e.message : e); }
    finally { bar.hidden = true; alignBtn.disabled = false; }
  });

  // ---------- Prüfen in „Alles bearbeiten“ ----------
  const sec = q('section[aria-label="Dokument"]');
  if (!sec) return;
  const grp = document.createElement('div');
  grp.className = 'fe-group';
  grp.innerHTML = '<button class="btn" type="button" id="fe-numcheck">Gleichungsnummern prüfen</button>' +
    '<p class="hint" id="fe-numcheck-msg" role="status"></p><div id="fe-numcheck-box" class="fe-box2" hidden></div>';
  sec.appendChild(grp);
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const list = a => a.map(x => '(' + esc(x) + ')').join(', ');
  grp.querySelector('#fe-numcheck').addEventListener('click', async () => {
    const m = grp.querySelector('#fe-numcheck-msg'), out = grp.querySelector('#fe-numcheck-box');
    if (!inWord()) { m.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    m.textContent = 'Dokument wird gelesen …'; out.hidden = true;
    try {
      let texts = [];
      await Word.run(async ctx => { const ps = ctx.document.body.paragraphs; ps.load('items/text'); await ctx.sync(); texts = ps.items.map(p => p.text); });
      const r = check(texts);
      m.textContent = '';
      if (!r.total) { m.textContent = 'Im Dokument stehen keine Gleichungsnummern am Ende von Formel-Absätzen.'; return; }
      let h = '<h3>Gleichungsnummern</h3><table><tr><td>Nummerierte Formeln</td><td class="n">' + r.total + '</td></tr>' +
        '<tr><td>Verweise im Text</td><td class="n">' + r.refs + '</td></tr>';
      r.chapters.forEach(c => {
        h += '<tr class="grp"><td colspan="2">' + (c.prefix ? 'Kapitel ' + esc(c.prefix) : 'Ohne Kapitel') + ': (' + esc(c.first) + ') bis (' + esc(c.last) + ')</td></tr>';
        h += '<tr><td class="sub">Formeln</td><td class="n">' + c.count + '</td></tr>';
        if (c.dup.length) h += '<tr><td class="sub" colspan="2"><mark>Doppelt:</mark> ' + list(c.dup) + '</td></tr>';
        if (c.gaps.length) h += '<tr><td class="sub" colspan="2"><mark>Fehlt:</mark> ' + list(c.gaps) + '</td></tr>';
        if (c.order.length) h += '<tr><td class="sub" colspan="2"><mark>Zählt rückwärts bei:</mark> ' + list(c.order) + '</td></tr>';
      });
      h += '</table>';
      if (r.missing.length) h += '<p class="hint"><mark>Verweise auf Nummern, die es nicht gibt:</mark> ' + list(r.missing) + '</p>';
      if (r.chapters.every(c => !c.dup.length && !c.gaps.length && !c.order.length) && !r.missing.length) h += '<p class="hint">Alles in Ordnung: Die Zählung ist lückenlos.</p>';
      out.innerHTML = h; out.hidden = false;
    } catch (e) { m.textContent = 'Die Prüfung hat nicht geklappt: ' + (e && e.message ? e.message : e); }
  });
})(typeof window !== 'undefined' ? window : globalThis);
