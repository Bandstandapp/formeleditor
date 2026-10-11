/* Anordnung im Word-Add-in: Kacheln Bausteine, Eingabe, Vorschau, Ganzes Dokument (bulk.js).
   Erklärtexte stehen hinter einem Info-Symbol neben der Kachel-Überschrift. Fett und Kursiv für
   markierte Zeichen. Stellt vorhandene Elemente nur um; Logik bleibt in index.src.html und addin.js. */
(function () {
  const q = s => document.querySelector(s);
  const bausteine = q('section[aria-label="Bausteine"]');
  const eingabe = q('section[aria-label="Eingabe"]');
  const aussehen = q('section[aria-label="Darstellung"]');
  const vorschau = q('section[aria-label="Ergebnis"]');

  const css = document.createElement('style');
  css.textContent = `
body.addin section[aria-label="Bausteine"] { order: 1; }
body.addin section[aria-label="Eingabe"] { order: 2; }
body.addin section[aria-label="Ergebnis"] { order: 3; }
body.addin section[aria-label="Dokument"] { order: 4; }
body.addin .label { position: relative; display: flex; align-items: center; gap: 6px; }
.info { all: unset; box-sizing: border-box; display: inline-grid; place-items: center; width: 1.25em; height: 1.25em; border: 1.5px solid currentColor; border-radius: 50%;
  font: italic 700 .8rem/1 var(--f-ui); text-transform: none; color: var(--accent); cursor: help; flex: none; }
.info:focus-visible { outline: 3px solid var(--focus); outline-offset: 2px; }
.info::after { content: attr(data-tip); display: none; position: absolute; z-index: 20; left: 0; right: 0; top: calc(100% + 6px);
  background: var(--fg); color: var(--surface); font: 400 .92rem/1.4 var(--f-ui); letter-spacing: 0; text-transform: none;
  padding: 8px 10px; border-radius: 8px; box-shadow: 0 6px 18px rgba(0,0,0,.25); white-space: pre-line; }
.info:hover::after, .info:focus::after { display: block; }
body.addin .tabs { flex-wrap: nowrap; gap: 6px; }
body.addin .tab { flex: 1 1 0; padding: 8px 6px; font-size: 1rem; text-align: center; }
body.addin #mfwrap math-field { min-height: 80px; padding-right: 4px; }
body.addin #mfwrap math-field { position: relative; }
body.addin #mfwrap math-field[data-state="saved"] { border: 3px solid #2e8540; }
body.addin #mfwrap math-field[data-state="changed"] { border: 3px solid #e07b00; }
body.addin #mfwrap math-field[data-state]:focus-within { outline: none; }
body.addin #mfwrap math-field::part(content) { padding-right: 34px; }
body.addin #mfwrap math-field::part(virtual-keyboard-toggle) { position: absolute; right: 2px; top: 6px; }
body.addin #mfwrap math-field::part(menu-toggle) { position: absolute; right: 2px; bottom: 6px; }
body.addin .stylebar { display: flex; gap: 6px; }
body.addin .stylebar button { font: inherit; width: 40px; height: 36px; border-radius: 8px; border: 2px solid var(--line); background: transparent; color: var(--fg); cursor: pointer; font-size: 1.1rem; }
body.addin .stylebar button[aria-pressed="true"] { border-color: var(--accent); background: var(--accent-soft); }
body.addin .stylebar .b { font-weight: 800; }
body.addin .stylebar .k { font-style: italic; }
body.addin .stylebar { position: relative; align-items: center; }
body.addin .stylebar .c { width: 52px; display: inline-flex; align-items: center; justify-content: center; gap: 3px; }
body.addin .stylebar .c i { font-style: normal; font-weight: 700; border-bottom: 4px solid var(--cur, #FF0000); line-height: 1; padding: 0 2px; }
body.addin .stylebar .c small { font-size: .7rem; }
body.addin .fe-swatches { display: flex; gap: 6px; padding: 6px; border: 1.5px solid var(--line); border-radius: 8px; background: var(--surface); }
body.addin .fe-swatches[hidden] { display: none !important; }
body.addin .stylebar .fe-swatches button { width: 30px; height: 30px; border-radius: 50%; border: 2px solid var(--line); padding: 0; }
body.addin h1, body.addin h2, body.addin h3 { font-family: var(--f-ui) !important; font-weight: 700; }
.fe-maintabs { order: 0; display: grid; grid-template-columns: 1fr 1fr; gap: 0; border: 2px solid var(--accent); border-radius: var(--r); overflow: hidden; }
.fe-maintabs button { font: inherit; font-weight: 700; font-size: 1rem; padding: 10px 6px; border: 0; background: var(--surface); color: var(--accent); cursor: pointer; }
.fe-maintabs button[aria-selected="true"] { background: var(--accent); color: var(--accent-fg); }
.fe-maintabs button:focus-visible { outline: 3px solid var(--focus); outline-offset: -3px; }
body.addin:not(.tab-doc) section[aria-label="Dokument"] { display: none !important; }
body.addin.tab-doc section[aria-label="Bausteine"], body.addin.tab-doc section[aria-label="Eingabe"], body.addin.tab-doc section[aria-label="Ergebnis"] { display: none !important; }
body.addin .editrow { display: flex; flex-wrap: wrap; gap: 5px; }
body.addin .editrow > * { flex: 1 1 auto; }   /* Knöpfe brechen in eine zweite Zeile um, statt abgeschnitten zu werden */
body.addin .editrow .btn.btn { padding: 6px 5px; font-size: .8rem; white-space: nowrap; }
body.addin .editrow .btn.fe-caret { padding: 0; flex-basis: 20px; }
body.addin button.btn.primary { width: 100%; font-size: 1.15rem; padding: 12px 14px; }
body.addin .copyrow { display: block; }
body.addin .fe-insrow { display: flex; flex-direction: column-reverse; gap: 6px; }
body.addin .fe-insrow .fe-cancel { width: 100%; padding: 7px 10px; font-size: .95rem; }
body.addin .fe-insrow .fe-cancel:disabled { opacity: .45; cursor: default; }
body.addin .hint:empty { display: none; }
body.addin #mode { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
body.addin #mode button { text-align: center; padding: 8px 4px; }
body.addin #mode small { display: none; }
body.addin .zoomrow { margin: 0; gap: 8px; font-weight: 400; font-size: .95rem; }
body.addin .pickrow { display: grid; grid-template-columns: minmax(0, 1.6fr) minmax(0, 1fr); gap: 8px; }
body.addin .pick { display: grid; gap: 2px; font-size: .9rem; color: var(--muted); }
body.addin select { font: inherit; font-size: 1rem; color: var(--fg); background: var(--surface); border: 2px solid var(--line); border-radius: 8px; padding: 6px 8px; min-width: 0; }
body.addin #fonts, body.addin #sizes, body.addin section[aria-label="Darstellung"] { display: none !important; }
body.addin .paper iframe { min-height: 0; }`;
  document.head.appendChild(css);

  // Überschrift mit Info-Symbol
  function heading(section, text, tip) {
    let lbl = section.querySelector(':scope > .label');
    if (!lbl) { lbl = document.createElement('p'); lbl.className = 'label'; section.prepend(lbl); }
    lbl.textContent = text;
    const i = document.createElement('button');
    i.type = 'button'; i.className = 'info'; i.textContent = 'i';
    i.setAttribute('aria-label', 'Info: ' + tip); i.dataset.tip = tip;
    lbl.appendChild(i);
    return lbl;
  }
  window.FeHeading = heading;
  // Info-Symbol an eine vorhandene Zeile hängen
  window.FeInfo = (el, tip) => {
    const i = document.createElement('button');
    i.type = 'button'; i.className = 'info'; i.textContent = 'i';
    i.setAttribute('aria-label', 'Info: ' + tip); i.dataset.tip = tip;
    el.style.position = 'relative'; el.style.display = 'flex'; el.style.alignItems = 'center'; el.style.gap = '6px';
    el.appendChild(i);
  };

  // Zwei Reiter: „Formel bearbeiten“ (Standard) und „Alles bearbeiten“
  const tabs = document.createElement('div');
  tabs.className = 'fe-maintabs'; tabs.setAttribute('role', 'tablist');
  tabs.innerHTML = '<button type="button" role="tab" data-t="formel" aria-selected="true">Formel bearbeiten</button>' +
                   '<button type="button" role="tab" data-t="doc" aria-selected="false">Alles bearbeiten</button>';
  (q('.wrap') || document.body).prepend(tabs);
  tabs.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    tabs.querySelectorAll('button').forEach(x => x.setAttribute('aria-selected', x === b));
    document.body.classList.toggle('tab-doc', b.dataset.t === 'doc');
    window.scrollTo({top: 0});
  }));

  heading(bausteine, 'Bausteine', 'Ein Klick setzt den Baustein an der Stelle des blinkenden Strichs in die Eingabe. Leere Kästchen, etwa Zähler und Nenner, der Reihe nach ausfüllen.');
  heading(eingabe, 'Eingabe',
    'Hier entsteht die Formel. Zum Ändern einer Formel sie im Dokument anklicken, dann erscheint sie hier.\n\n' +
    'Tippen: / macht einen Bruch, ^ eine Hochzahl, _ einen Index. Die Pfeiltasten wechseln zwischen den Feldern.\n\n' +
    'F macht markierte Zeichen fett, K schaltet zwischen kursiv und aufrecht um, A färbt sie. Ohne Markierung gilt das für die nächsten Zeichen. Die Leertaste setzt einen kleinen Abstand, etwa in 1.159 652 180 46.');
  heading(vorschau, 'Vorschau', 'So erscheint die Formel in Word. Der Zoom vergrößert nur diese Ansicht, nicht die Formel im Dokument. Schrift und Schriftgröße gelten für die Formel in der Eingabe.');

  // Erklärtexte entfernen; Rückmeldungen (pastemsg, copymsg, Status) bleiben, starten aber leer
  eingabe.querySelectorAll('.hint').forEach(p => { if (!p.id && !p.getAttribute('role')) p.remove(); });
  ['pastemsg', 'copymsg'].forEach(id => { const el = q('#' + id); if (el) el.textContent = ''; });
  const fromWord = q('#fromword');
  if (fromWord) fromWord.remove();          // Formeln laden sich beim Anklicken von selbst

  // Fett und Kursiv über dem Eingabefeld
  const bar = document.createElement('div');
  bar.className = 'stylebar';
  bar.innerHTML = '<button type="button" class="b" title="Fett" aria-label="Fett" aria-pressed="false">F</button>' +
                  '<button type="button" class="k" title="Kursiv" aria-label="Kursiv" aria-pressed="false">K</button>';
  // Farbe: kleiner Knopf mit Farbauswahl (Farben wie in Papas Dokumenten); Schwarz entfernt die Farbe
  const COLORS = [['Schwarz', 'none', '#111'], ['Rot', '#FF0000'], ['Grün', '#00B050'], ['Hellblau', '#00B0F0'], ['Blau', '#4472C4']];
  const cBtn = document.createElement('button');
  cBtn.type = 'button'; cBtn.className = 'c'; cBtn.title = cBtn.ariaLabel = 'Farbe';
  cBtn.setAttribute('aria-expanded', 'false');
  cBtn.innerHTML = '<i>A</i><small>▾</small>';
  const sw = document.createElement('div');
  sw.className = 'fe-swatches'; sw.hidden = true;
  COLORS.forEach(([name, val, swatch]) => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'fe-swatch'; b.title = b.ariaLabel = name; b.dataset.color = val;
    b.style.background = swatch || val;
    b.addEventListener('mousedown', e => e.preventDefault());
    b.addEventListener('click', () => {
      mf.applyStyle({color: val});
      if (val !== 'none') bar.style.setProperty('--cur', val);
      sw.hidden = true; cBtn.setAttribute('aria-expanded', 'false');
      mf.focus(); show();
    });
    sw.appendChild(b);
  });
  cBtn.addEventListener('mousedown', e => e.preventDefault());
  cBtn.addEventListener('click', () => { sw.hidden = !sw.hidden; cBtn.setAttribute('aria-expanded', String(!sw.hidden)); });
  bar.append(cBtn, sw);
  q('#mfwrap').before(bar);
  const [bBtn, kBtn] = bar.querySelectorAll('button');
  const all = v => v === 'all';
  const selLatex = () => { try { return mf.getValue(mf.selection, 'latex'); } catch (e) { return ''; } };
  function styleState() {
    if (!mf || !mf.queryStyle) return {bold: false, italic: false};
    const bold = all(mf.queryStyle({variantStyle: 'bold'}));
    const up = all(mf.queryStyle({variantStyle: 'up'}));
    const it = all(mf.queryStyle({variantStyle: 'italic'}));
    const letters = /[A-Za-zα-ωΑ-Ω]|\\[a-zA-Z]+/.test(selLatex().replace(/\\math[a-z]+/g, ''));
    return {bold, italic: it || (!up && !bold && letters)};
  }
  function syncStyle() { const s = styleState(); bBtn.setAttribute('aria-pressed', s.bold); kBtn.setAttribute('aria-pressed', s.italic); }
  bBtn.addEventListener('mousedown', e => e.preventDefault());
  kBtn.addEventListener('mousedown', e => e.preventDefault());
  bBtn.addEventListener('click', () => {
    if (styleState().bold) mf.applyStyle({variantStyle: ''}); else mf.applyStyle({variantStyle: 'bold'});
    mf.focus(); syncStyle(); show();
  });
  kBtn.addEventListener('click', () => {
    const up = all(mf.queryStyle({variantStyle: 'up'})), it = all(mf.queryStyle({variantStyle: 'italic'}));
    if (up || it) mf.applyStyle({variantStyle: ''});
    else if (/[A-Za-zα-ωΑ-Ω]|\\[a-zA-Z]+/.test(selLatex())) mf.applyStyle({variantStyle: 'up'});
    else mf.insert('\\mathit{#@}');
    mf.focus(); syncStyle(); show();
  });
  const hookMf = () => { if (window.mf && mf.addEventListener) mf.addEventListener('selection-change', syncStyle); else setTimeout(hookMf, 200); };
  hookMf();

  // Eingabe-Knöpfe in einer Reihe
  const editRow = eingabe.querySelector('.row');
  editRow.classList.add('editrow');

  // Vorschau: Anzeige, Umschalter, Zoom, Schrift und Größe, Einfügen
  const paper = vorschau.querySelector('.paper');
  const modeBox = q('#mode');
  const zoomRow = vorschau.querySelector('.zoomrow');
  paper.after(modeBox);
  modeBox.after(zoomRow);
  const zl = zoomRow.querySelector('label'); if (zl) zl.textContent = 'Zoom';

  const pick = document.createElement('div');
  pick.className = 'pickrow';
  pick.innerHTML = '<label class="pick">Schrift<select id="fe-font"></select></label><label class="pick">Schriftgröße<select id="fe-size"></select></label>';
  zoomRow.after(pick);
  const fontSel = pick.querySelector('#fe-font'), sizeSel = pick.querySelector('#fe-size');
  FONTS.forEach(f => fontSel.add(new Option(f.name, f.id)));
  const SIZES = [10, 11, 12, 14, 16, 20, 24];
  function syncPick() {
    fontSel.value = font;
    const want = SIZES.includes(sizePt) ? SIZES : SIZES.concat(sizePt).sort((x, y) => x - y);
    if ([...sizeSel.options].map(o => +o.value).join() !== want.join()) {
      sizeSel.innerHTML = '';
      want.forEach(p => sizeSel.add(new Option(String(p).replace('.', ',') + ' pt', p)));
    }
    sizeSel.value = sizePt;
  }
  fontSel.addEventListener('change', () => { const b = q('#font-' + fontSel.value); if (b) b.click(); });
  sizeSel.addEventListener('change', () => {
    const v = +sizeSel.value, b = q('#size-' + v);
    if (b) b.click(); else { sizePt = v; syncSeg(); }
  });
  syncPick();
  new MutationObserver(syncPick).observe(document.body, {subtree: true, attributes: true, attributeFilter: ['aria-pressed']});

  // Der untere Einfügen-Knopf ans Ende der Vorschau, gleich gestaltet wie der obere
  const copyRow = vorschau.querySelector('.copyrow');
  pick.after(copyRow);
  const copymsg = q('#copymsg');
  if (copymsg) copyRow.after(copymsg);

  // Aussehen-Kachel entfällt; die alten Knopfreihen bleiben unsichtbar erhalten, weil die Editor-Logik sie nutzt
  const keep = document.createElement('div');
  keep.hidden = true;
  const sizesWrap = q('.sizes');
  keep.append(q('#fonts'), q('#sizes'));
  if (sizesWrap) sizesWrap.remove();
  document.body.appendChild(keep);
  if (aussehen) aussehen.remove();
})();
