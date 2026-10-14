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
  // Absatz mit nummerierter Formel neu setzen (OOXML des Absatzes, wie getOoxml() es liefert):
  // Einzug links leftTw, ein rechtsbündiger Tabstopp bei tabTw (beides Twips ab linkem Seitenrand),
  // Leerraum vor der ersten Formel weg, zwischen Formel und Nummer genau ein Tabulator.
  // Liefert das geänderte OOXML, null, wenn die Nummer nicht zu finden ist, und '', wenn im Absatz
  // gar keine Formel steht (eine Textzeile, die nur zufällig mit „(1)“ endet, bleibt unberührt).
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
    'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
    'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  function layoutXml(flat, num, leftTw, tabTw) {
    const doc = new DOMParser().parseFromString(flat, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) return null;
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    const p = body && Array.from(body.childNodes).find(n => n.localName === 'p' && n.namespaceURI === W);
    if (!p) return null;
    const el = (name, attrs) => { const e = doc.createElementNS(W, 'w:' + name); Object.entries(attrs || {}).forEach(([k, v]) => e.setAttributeNS(W, 'w:' + k, String(v))); return e; };
    // Bausteine des Absatzes der Reihe nach: Text, Tabs, Umbrüche, Formeln (Bild oder altes Objekt)
    const atoms = [];
    p.querySelectorAll('*').forEach(n => {
      if (n.namespaceURI !== W || !n.parentNode || n.parentNode.localName !== 'r') return;
      const k = n.localName;
      if (k === 't') atoms.push({n, kind: 't', text: n.textContent});
      else if (k === 'tab' || k === 'ptab') atoms.push({n, kind: 'ws', text: '\t'});
      else if (k === 'br' || k === 'cr') atoms.push({n, kind: 'ws', text: '\n'});
      else if (k === 'drawing' || k === 'object' || k === 'pict') atoms.push({n, kind: 'obj', text: '\ufffc'});
    });
    if (!atoms.some(a => a.kind === 'obj') && !p.getElementsByTagNameNS('http://schemas.openxmlformats.org/officeDocument/2006/math', 'oMath').length) return '';
    const all = atoms.map(a => a.text).join('');
    const want = '(' + num + ')';
    const at = all.lastIndexOf(want);
    if (at < 0 || all.slice(at + want.length).trim()) return null;
    // Stelle der Klammer finden
    let pos = 0, ai = -1, off = 0;
    for (let i = 0; i < atoms.length; i++) { const L = atoms[i].text.length; if (at < pos + L) { ai = i; off = at - pos; break; } pos += L; }
    if (ai < 0 || atoms[ai].kind !== 't') return null;
    // Leerraum vor der Nummer entfernen, bis zur Formel oder zu sichtbarem Text
    const a = atoms[ai];
    const head = a.text.slice(0, off).replace(/[\s\u00a0]+$/, '');
    let stop = !!head;
    for (let i = ai - 1; i >= 0 && !stop; i--) {
      const b = atoms[i];
      if (b.kind === 'ws') { b.n.parentNode.removeChild(b.n); continue; }
      if (b.kind === 'obj') break;
      const t = b.text.replace(/[\s\u00a0]+$/, '');
      b.n.textContent = t;
      if (t) stop = true;
    }
    // Ein Tabulator vor der Nummer
    a.n.textContent = a.text.slice(off);
    a.n.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    if (head) { const h = el('t'); h.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); h.textContent = head; a.n.parentNode.insertBefore(h, a.n); }
    a.n.parentNode.insertBefore(el('tab'), a.n);
    // Leerraum vor der ersten Formel entfernen
    const first = atoms.findIndex(x => x.kind === 'obj');
    if (first > 0 && atoms.slice(0, first).every(x => x.kind === 'ws' || (x.kind === 't' && !x.text.trim()))) {
      atoms.slice(0, first).forEach(x => { if (x.kind === 'ws') { if (x.n.parentNode) x.n.parentNode.removeChild(x.n); } else x.n.textContent = ''; });
    }
    // Absatzeigenschaften: Einzug, linksbündig, ein rechter Tabstopp für die Nummer
    let pPr = Array.from(p.childNodes).find(n => n.localName === 'pPr');
    if (!pPr) { pPr = el('pPr'); p.insertBefore(pPr, p.firstChild); }
    Array.from(pPr.childNodes).filter(n => ['tabs', 'ind', 'jc'].includes(n.localName)).forEach(n => pPr.removeChild(n));
    const tabs = el('tabs'); tabs.appendChild(el('tab', {val: 'right', pos: Math.round(tabTw)}));
    const put = e => {
      const i = PPR.indexOf(e.localName);
      const after = Array.from(pPr.childNodes).find(n => PPR.indexOf(n.localName) > i);
      pPr.insertBefore(e, after || null);
    };
    put(tabs);
    put(el('ind', {left: Math.round(leftTw), right: 0, firstLine: 0}));
    put(el('jc', {val: 'left'}));
    // Abschnittsangaben des Dokuments nicht mit einfügen
    Array.from(body.childNodes).filter(n => n.localName === 'sectPr').forEach(n => body.removeChild(n));
    return new XMLSerializer().serializeToString(doc);
  }
  const FeNum = {endNumber, suggest, shifts, check, split, VALID, layoutXml};
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
  const L = () => window.FeLineal || {leftPt: () => 0, tabPt: () => 0, text: () => '', use: () => false, visible: () => false};
  const useRuler = () => L().visible() && L().use();
  function indent(para) {
    para.leftIndent = L().leftPt();
    para.firstLineIndent = 0;
    para.rightIndent = 0;
    para.alignment = 'Left';
  }
  // Eigene Absatzvorlage „Formel (Editor)“ mit „Standard“ als Folgeabsatz: Nach Return hinter einer Formel
  // geht es mit normalem Text weiter, ohne den Einzug und den Tabstopp aus dem Lineal (wie bei MathType).
  const STIL = 'Formel (Editor)';
  let stilOk = false;
  async function stil(ctx, para) {
    try {
      if (!Office.context.requirements.isSetSupported('WordApi', '1.5')) return;
      if (!stilOk) {
        const styles = ctx.document.getStyles();
        styles.load('items/nameLocal,items/type');
        await ctx.sync();
        const names = styles.items.map(s => s.nameLocal);
        if (!names.includes(STIL)) {
          const normal = ['Standard', 'Normal'].find(n => names.includes(n));
          const st = ctx.document.addStyle(STIL, 'Paragraph');
          if (normal) { st.baseStyle = normal; st.nextParagraphStyle = normal; }
          await ctx.sync();
        }
        stilOk = true;
      }
      para.style = STIL;
      await ctx.sync();
    } catch (e) { stilOk = false; }
  }
  async function prepare(ctx, para) { await stil(ctx, para); indent(para); }
  const inWord = () => typeof Word !== 'undefined' && typeof Office !== 'undefined' && Office.context && Office.context.document;

  // Nummer in der Vorschau anzeigen
  function preview() {
    try {
      const d = q('#out').contentDocument, no = d && d.querySelector('.eq .no');
      if (no) no.textContent = FeWord.editing() ? (editNum ? '(' + editNum + ')' : '')
        : !on.disabled && on.checked && field.value.trim() ? '(' + field.value.trim() + ')' : '';
    } catch (e) {}
  }
  // Immer zu sehen; wirkt nur bei neuen abgesetzten Formeln (wissenschaftlich, nicht beim Ersetzen).
  // Bei einer geladenen Formel steht ihre vorhandene Nummer im Kästchen und in der Vorschau.
  let editNum = null, wasEditing = false, ownChoice = false;
  async function readEditNum() {
    editNum = null;
    if (!inWord()) return;
    try {
      await Word.run(async ctx => {
        const para = ctx.document.getSelection().paragraphs.getFirst();
        para.load('text');
        await ctx.sync();
        editNum = endNumber(para.text);
      });
    } catch (e) {}
    sync();
  }
  function sync() {
    const ed = FeWord.editing();
    if (ed !== wasEditing) {
      wasEditing = ed;
      if (ed) { ownChoice = on.checked; editNum = null; readEditNum(); }
      else { on.checked = ownChoice; editNum = null; }
    }
    if (ed) on.checked = !!editNum;
    const w = mode !== 'sci' ? 'Gleichungsnummern gibt es nur bei der Darstellung „Wissenschaftlich“.'
      : ed ? (editNum ? 'Diese Formel hat die Nummer (' + editNum + '). Sie bleibt beim Ersetzen erhalten.'
        : 'Diese Formel hat keine Nummer. Eine Nummer gibt es beim Einfügen einer neuen Formel.') : '';
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
    ownChoice = on.checked;
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
  let nachher = null;                     // Absatz, der vor dem Einfügen auf den Formel-Absatz folgte
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
    nachher = para.getNextOrNullObject();
    nachher.load('isNullObject');
    await ctx.sync();
    if (para.text.replace(/[\s\ufffc]/g, '')) {
      const np = para.insertParagraph('', 'After');
      if (useRuler()) await prepare(ctx, np);
      return np.getRange('Start');
    }
    if (useRuler()) await prepare(ctx, para);
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
    await prepare(ctx, para);
    return L().pPr();
  }
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
  async function afterInsert(ctx, range, num) {
    if (!useRuler()) return;
    const para = range.paragraphs.getFirst();
    if (num) { const n = nachher; nachher = null; await alignNumbered(ctx, para, num, n); return; }
    para.load('text');
    await ctx.sync();
    if (para.text.replace(/[\s\ufffc]/g, '')) return;
    await prepare(ctx, para);
  }
  root.FeNummer = {wanted, beforeInsert, placeUnnumbered, pPr, afterInsert, afterReplace, done};

  // Formeln mit Nummer wie im Lineal setzen: Einzug links für die Formel, rechtsbündiger Tabstopp am Ende der Nummer.
  // Der Absatz wird über sein OOXML neu geschrieben (layoutXml), denn Tabstopps kennt Office.js nicht.
  const blank = t => /^[\s\u00a0]*$/.test(t || '');
  const TW = pt => pt * 20;
  async function alignNumbered(ctx, p, num, next0) {
    const ox0 = p.getOoxml();
    const neu = !!next0;                  // gerade eingefügt: danach Cursor ans Ende der Formelzeile
    if (!next0) { next0 = p.getNextOrNullObject(); next0.load('isNullObject'); }
    await ctx.sync();
    if (layoutXml(ox0.value, num, 0, 0) === '') return 'none';
    await stil(ctx, p);
    const ox = p.getOoxml();
    await ctx.sync();
    const xml = layoutXml(ox.value, num, TW(L().leftPt()), TW(L().tabPt()));
    if (!xml) return false;
    const fp = p.insertOoxml(xml, 'Replace').paragraphs.getFirst();
    await ctx.sync();
    // Bleibt der alte Absatz leer zurück, wird er entfernt
    try {
      p.load('text');
      await ctx.sync();
      if (!p.text.replace(/[\s\u00a0\ufffc]/g, '')) { p.delete(); await ctx.sync(); }
    } catch (e) {}
    await leftover(ctx, fp, next0, neu);
    return true;
  }
  // Word lässt beim Einfügen eines ganzen Absatzes einen leeren Absatz dahinter zurück, mit Formel-Vorlage,
  // Einzug und Tabstopp der Formel. Wer dort weiterschreibt, schreibt eingerückt. Solche Absätze zwischen der
  // Formel und dem Absatz, der vorher folgte, werden entfernt; der Cursor steht danach am Ende der Formelzeile,
  // so dass Return einen normalen Absatz („Standard“) beginnt.
  async function leftover(ctx, fp, next0, cursor) {
    try {
      for (let i = 0; i < 3; i++) {
        const nx = fp.getNextOrNullObject();
        nx.load('isNullObject,text');
        await ctx.sync();
        if (nx.isNullObject) break;
        if (next0 && !next0.isNullObject) {
          const c = nx.getRange('Whole').compareLocationWith(next0.getRange('Whole'));
          await ctx.sync();
          if (c.value !== 'Before') break;
        }
        const pics = nx.inlinePictures;
        pics.load('items');
        await ctx.sync();
        if (nx.text.replace(/[\s\u00a0]/g, '') || pics.items.length) break;
        nx.delete();
        await ctx.sync();
      }
    } catch (e) {}
    if (cursor) try { fp.getRange('End').select(); await ctx.sync(); } catch (e) {}
  }
  // Einen Absatz mit abgesetzter Formel wie im Lineal setzen; num: seine Nummer oder null. false, wenn die Nummer fehlt.
  async function alignPara(ctx, p, num) {
    if (num) return alignNumbered(ctx, p, num);
    const pics = p.inlinePictures;
    pics.load('items');
    await ctx.sync();
    // Leerraum vor der ersten Formel entfernen, damit sie beim Einzug aus dem Lineal beginnt
    if (pics.items.length) {
      const lead = p.getRange('Start').expandTo(pics.items[0].getRange('Start'));
      lead.load('text');
      await ctx.sync();
      if (lead.text && blank(lead.text)) lead.delete();
    }
    await prepare(ctx, p);
    return true;
  }
  async function alignAll(progress) {
    stilOk = false;                       // Vorlage notfalls neu anlegen, falls sie gelöscht wurde
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
        const r = await alignPara(ctx, p, num);
        if (r === 'none') { /* Textzeile ohne Formel */ } else if (r) { done++; if (num) numbered++; } else skipped.push(num);
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
  // Absatzeigenschaften für eine neue nummerierte Formel; ohne Lineal steht die Nummer am rechten Rand
  function pPr() { return useRuler() ? L().pPr() : ''; }
  // Knopf unter dem Lineal (lineal.js): alle Formeln ausrichten, mit Fortschrittsbalken
  const alignBtn = q('#fe-align'), bar = q('#fe-align-bar'), fill = bar && bar.firstElementChild, amsg = q('#fe-align-msg');
  if (alignBtn) alignBtn.addEventListener('click', async () => {
    if (!inWord()) { amsg.textContent = 'Das geht nur, wenn der Editor in Word geöffnet ist.'; return; }
    alignBtn.disabled = true;
    amsg.textContent = 'Formeln werden ausgerichtet …';
    fill.style.width = '0%'; bar.hidden = false;
    try {
      const r = await alignAll((i, n) => {
        fill.style.width = (n ? Math.round(100 * i / n) : 100) + '%';
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
