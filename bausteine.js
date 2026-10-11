/* Bausteine im Word-Add-in: Suchfeld für Zeichen nach Namen, die zuletzt benutzten Bausteine und
   „Meine Formeln“ (häufige Ausdrücke unter einem Namen speichern und mit einem Klick einsetzen).
   Baut auf TOOLS, buildTools und useTool aus index.src.html auf. */
(function () {
  const q = s => document.querySelector(s);
  const sec = q('section[aria-label="Bausteine"]');
  const tabsEl = sec && sec.querySelector('.tabs');
  if (!sec || !tabsEl) return;
  const store = {
    get: (k, d) => { try { const v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } },
    set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
  };
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // ---------- Namen zum Suchen ----------
  // Deutsche Namen und gebräuchliche Bezeichnungen aus der Physik; Schlüssel ist der eingesetzte Text
  const GREEK = {alpha: 'Alpha Feinstrukturkonstante', beta: 'Beta', gamma: 'Gamma Lorentzfaktor', delta: 'Delta', varepsilon: 'Epsilon Dielektrizitätskonstante',
    zeta: 'Zeta', eta: 'Eta', theta: 'Theta Winkel', vartheta: 'Theta geschwungen Winkel', kappa: 'Kappa', lambda: 'Lambda Wellenlänge', mu: 'My Mü Permeabilität magnetische Feldkonstante',
    nu: 'Ny Frequenz', xi: 'Xi', uppi: 'Pi Kreiszahl', rho: 'Rho Dichte', sigma: 'Sigma', tau: 'Tau', varphi: 'Phi Winkel', phi: 'Phi gerade', chi: 'Chi', psi: 'Psi Wellenfunktion',
    omega: 'Omega Kreisfrequenz', Gamma: 'Gamma groß', Delta: 'Delta groß Differenz Laplace', Theta: 'Theta groß', Lambda: 'Lambda groß', Xi: 'Xi groß', Pi: 'Pi groß Produkt',
    Sigma: 'Sigma groß Summe', Phi: 'Phi groß Fluss', Psi: 'Psi groß', Omega: 'Omega groß Ohm'};
  const NAMES = {
    '\\hbar ': 'h quer Planck reduziert Wirkungsquantum hbar', '\\hslash ': 'h quer schräg Planck', '\\lambdabar ': 'Lambda quer Compton-Wellenlänge', '\\lambdaslash ': 'Lambda quer schräg Compton',
    '\\partial ': 'Partiell del Ableitung rund d', '\\nabla ': 'Nabla Gradient Divergenz Rotation', '\\infty ': 'Unendlich', '\\pm ': 'Plus minus', '\\mp ': 'Minus plus',
    '\\cdot ': 'Mal Punkt multiplizieren Skalarprodukt', '\\times ': 'Mal Kreuz Kreuzprodukt Vektorprodukt', '=': 'Gleich', '\\approx ': 'Ungefähr gleich etwa', '\\ne ': 'Ungleich',
    '\\equiv ': 'Identisch definiert kongruent', '\\sim ': 'Proportional ähnlich Tilde Größenordnung', '\\propto ': 'Proportional', '<': 'Kleiner', '>': 'Größer', '\\le ': 'Kleiner gleich',
    '\\ge ': 'Größer gleich', '\\ll ': 'Viel kleiner', '\\gg ': 'Viel größer', '\\to ': 'Pfeil nach rechts gegen strebt', '\\Rightarrow ': 'Folgepfeil daraus folgt Implikation',
    '\\Leftrightarrow ': 'Äquivalenz genau dann Doppelpfeil', '\\in ': 'Element von', '^{\\circ}': 'Grad Winkel', "'": 'Strich Ableitung Apostroph', '^{\\ast}': 'Stern konjugiert komplex',
    '\\cdots ': 'Punkte mittig Pünktchen', '\\ldots ': 'Punkte unten Pünktchen', '\\perp ': 'Senkrecht orthogonal', '\\parallel ': 'Parallel', '\\mathrm{Å}': 'Ångström Angström',
    '^{\\dagger}': 'Dagger Kreuz adjungiert hermitesch', '\\otimes ': 'Tensorprodukt Kreuz im Kreis', '\\forall ': 'Für alle', '\\exists ': 'Es existiert', '\\notin ': 'Nicht Element',
    '\\subset ': 'Teilmenge', '\\cup ': 'Vereinigung', '\\cap ': 'Schnittmenge', '\\leftarrow ': 'Pfeil nach links', '\\leftrightarrow ': 'Pfeil beidseitig Doppelpfeil',
    '^{-1}': 'Hoch minus eins Kehrwert pro', '^{2}': 'Quadrat hoch zwei', '^{3}': 'Kubik hoch drei', '/': 'Schrägstrich pro geteilt', '\\Omega ': 'Ohm Widerstand Omega',
    '\\upmu\\mathrm{m}': 'Mikrometer', '\\upmu\\mathrm{s}': 'Mikrosekunde', '\\upmu ': 'Mikro my', '^{\\circ}\\mathrm{C}': 'Grad Celsius Temperatur',
  };
  const UNITS = {m: 'Meter', s: 'Sekunde', kg: 'Kilogramm', g: 'Gramm', N: 'Newton Kraft', J: 'Joule Energie', W: 'Watt Leistung', V: 'Volt Spannung', A: 'Ampere Strom',
    T: 'Tesla Magnetfeld', Hz: 'Hertz Frequenz', Pa: 'Pascal Druck', C: 'Coulomb Ladung', K: 'Kelvin Temperatur', mol: 'Mol Stoffmenge', eV: 'Elektronenvolt',
    Wb: 'Weber Fluss', F: 'Farad Kapazität', H: 'Henry Induktivität', cm: 'Zentimeter', mm: 'Millimeter', nm: 'Nanometer', km: 'Kilometer', ms: 'Millisekunde', ns: 'Nanosekunde', rad: 'Radiant Bogenmaß'};
  const TAB_NAMES = {aufbau: 'Aufbau', greek: 'Griechisch', zeichen: 'Zeichen', einheiten: 'Einheiten'};
  const plain = s => String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ß/g, 'ss');
  const keyOf = t => t.cmd || t.ins;
  const all = [];
  Object.keys(TAB_NAMES).forEach(tab => (TOOLS[tab] || []).forEach(t => {
    const ins = t.ins || '';
    const um = ins.match(/^\\mathrm\{([A-Za-z]+)\}$/);
    const gm = ins.match(/^\\([A-Za-z]+) $/);
    const words = [t.n ? t.n.replace(/&shy;/g, '').replace(/<[^>]+>/g, '') : '', NAMES[ins] || '',
      tab === 'greek' && gm ? GREEK[gm[1]] || gm[1] : '', tab === 'einheiten' && um ? (UNITS[um[1]] || '') + ' Einheit' : '', t.sym ? t.g.replace(/<[^>]+>/g, '') : ''].join(' ');
    all.push({t, tab, words: plain(words)});
  }));
  const byKey = new Map(all.map(x => [keyOf(x.t), x]));

  // ---------- Suchfeld und „Zuletzt benutzt“ ----------
  const css = document.createElement('style');
  css.textContent = `
.fe-find { position: relative; }
.fe-find input { font: inherit; font-size: 1rem; width: 100%; padding: 8px 34px 8px 10px; border: 2px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--fg); }
.fe-find input:focus-visible { outline: 3px solid var(--focus); border-color: var(--accent); }
.fe-find button { all: unset; position: absolute; right: 6px; top: 50%; transform: translateY(-50%); width: 26px; height: 26px; display: grid; place-items: center; border-radius: 50%; cursor: pointer; color: var(--muted); font-size: 1.1rem; }
.fe-find button[hidden] { display: none; }
.fe-recent { display: flex; align-items: center; gap: 5px; flex-wrap: wrap; }
.fe-recent[hidden] { display: none; }
.fe-recent > span { font-size: .85rem; color: var(--muted); margin-right: 2px; }
.fe-recent .tool.tool { min-height: 0; padding: 3px 8px; min-width: 34px; }
.fe-recent .tool .g { font-size: 1.15rem; min-height: 1.3rem; }
.fe-recent .tool .n { display: none; }
.fe-hits-note { margin: 0; font-size: .9rem; color: var(--muted); }
body.addin .tabs.tabs { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; }
body.addin .tabs .tab { padding: 7px 4px; font-size: .95rem; white-space: nowrap; min-width: 0; }
body.addin .tabs #tab-meine { grid-column: span 2; }
body.fe-searching .tabs .tab[aria-selected="true"] { background: transparent; color: var(--fg); border-color: var(--line); }
#fe-search::-webkit-search-cancel-button { -webkit-appearance: none; display: none; }
.fe-mine { display: grid; gap: 8px; }
.fe-mine-save { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 6px; }
.fe-mine-save input { font: inherit; font-size: 1rem; padding: 6px 8px; border: 2px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--fg); min-width: 0; }
.fe-mine-list { display: grid; gap: 6px; }
.fe-mine-item { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 4px; align-items: stretch; }
.fe-mine-item .fe-use { font: inherit; text-align: left; display: grid; gap: 2px; padding: 6px 10px; border: 1px solid var(--line); border-radius: var(--r); background: var(--bg); color: var(--fg); cursor: pointer; min-width: 0; overflow-x: auto; }
.fe-mine-item .fe-use:hover { border-color: var(--accent); background: var(--accent-soft); }
.fe-mine-item .fe-use b { font-size: .9rem; }
.fe-mine-item .fe-use .f { font-size: 1.3rem; min-height: 1.4em; }
.fe-mine-item .fe-del { font: inherit; width: 38px; border: 1px solid var(--line); border-radius: var(--r); background: transparent; color: var(--muted); cursor: pointer; font-size: 1.1rem; }
.fe-mine-item .fe-del.sure { width: auto; padding: 0 8px; color: #fff; background: #b42318; border-color: #b42318; font-size: .85rem; font-weight: 700; }`;
  document.head.appendChild(css);

  const find = document.createElement('div');
  find.className = 'fe-find';
  find.innerHTML = '<input type="search" id="fe-search" placeholder="Suchen, z. B. Wurzel, Planck" aria-label="Zeichen suchen" autocomplete="off">' +
    '<button type="button" aria-label="Suche löschen" hidden>×</button>';
  const recent = document.createElement('div');
  recent.className = 'fe-recent'; recent.hidden = true;
  tabsEl.before(find);
  tabsEl.after(recent);
  const input = find.querySelector('input'), clearBtn = find.querySelector('button');
  const grid = q('#tools');

  function toolButton(t) {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'tool';
    b.innerHTML = '<span class="g"><span class="w">' + t.g + '</span></span>' + (t.n ? '<span class="n">' + t.n + '</span>' : '');
    if (!t.n) b.setAttribute('aria-label', t.g.replace(/<[^>]+>/g, ''));
    b.addEventListener('mousedown', e => e.preventDefault());
    b.addEventListener('click', () => useTool(t));
    return b;
  }
  const selectedTab = () => { const t = tabsEl.querySelector('.tab[aria-selected="true"]'); return t ? t.dataset.tab : 'aufbau'; };

  function runSearch() {
    const words = plain(input.value).split(/\s+/).filter(Boolean);
    clearBtn.hidden = !input.value;
    document.body.classList.toggle('fe-searching', !!words.length);
    if (!words.length) { buildTools(selectedTab()); showRecent(); return; }
    recent.hidden = true;
    const hits = all.filter(x => words.every(w => x.words.includes(w) || x.words.split(' ').some(v => v.startsWith(w))));
    grid.className = 'tools sym';
    grid.innerHTML = '';
    hits.slice(0, 40).forEach(x => grid.appendChild(toolButton(x.t)));
    if (!hits.length) {
      const p = document.createElement('p'); p.className = 'fe-hits-note';
      p.textContent = 'Nichts gefunden. Andere Bezeichnung versuchen, etwa „Pfeil“ oder „Quadrat“.';
      grid.appendChild(p);
    }
  }
  input.addEventListener('input', runSearch);
  input.addEventListener('keydown', e => {
    if (e.key === 'Escape') { input.value = ''; runSearch(); }
    if (e.key === 'Enter') { const b = grid.querySelector('.tool'); if (b) { e.preventDefault(); b.click(); } }
  });
  clearBtn.addEventListener('click', () => { input.value = ''; runSearch(); input.focus(); });
  tabsEl.addEventListener('click', e => { if (e.target.closest('.tab') && input.value) { input.value = ''; clearBtn.hidden = true; document.body.classList.remove('fe-searching'); } });

  // Zuletzt benutzte Bausteine (bis zu acht), gemerkt auf diesem Mac
  function showRecent() {
    const keys = store.get('fe-zuletzt', []).filter(k => byKey.has(k));
    recent.innerHTML = '';
    if (!keys.length || input.value) { recent.hidden = true; return; }
    const lbl = document.createElement('span'); lbl.textContent = 'Zuletzt:';
    recent.appendChild(lbl);
    keys.forEach(k => { const b = toolButton(byKey.get(k).t); b.title = (byKey.get(k).t.n || '').replace(/&shy;/g, ''); recent.appendChild(b); });
    recent.hidden = false;
  }
  document.addEventListener('fe-tool', e => {
    const k = keyOf(e.detail);
    if (!byKey.has(k) || e.detail.cmd) return;
    store.set('fe-zuletzt', [k].concat(store.get('fe-zuletzt', []).filter(x => x !== k)).slice(0, 8));
    if (!input.value) showRecent();
  });
  showRecent();

  // ---------- Meine Formeln ----------
  const tab = document.createElement('button');
  tab.className = 'tab'; tab.type = 'button'; tab.setAttribute('role', 'tab'); tab.id = 'tab-meine';
  tab.dataset.tab = 'meine'; tab.setAttribute('aria-selected', 'false'); tab.textContent = 'Meine Formeln';
  tab.addEventListener('mousedown', e => e.preventDefault());
  tab.addEventListener('click', () => {
    tabsEl.querySelectorAll('.tab').forEach(x => x.setAttribute('aria-selected', x === tab));
    buildTools('meine');
  });
  tabsEl.appendChild(tab);

  const mine = () => store.get('fe-meine', []).filter(x => x && typeof x.tex === 'string');
  let askDel = -1;
  function renderMine() {
    grid.className = 'fe-mine';
    recent.hidden = true;
    grid.innerHTML = '<div class="fe-mine-save"><input type="text" id="fe-mine-name" placeholder="Name, z. B. Feinstrukturkonstante" aria-label="Name für die Formel" autocomplete="off">' +
      '<button class="btn" type="button" id="fe-mine-add">Speichern</button></div><p class="fe-hits-note" id="fe-mine-msg" role="status"></p><div class="fe-mine-list"></div>';
    const list = grid.querySelector('.fe-mine-list'), msg = grid.querySelector('#fe-mine-msg');
    const items = mine();
    if (!items.length) msg.textContent = 'Noch nichts gespeichert. Formel in der Eingabe bauen (oder einen Teil markieren), Namen eintragen und „Speichern“ drücken.';
    items.forEach((it, i) => {
      const row = document.createElement('div'); row.className = 'fe-mine-item';
      row.innerHTML = '<button type="button" class="fe-use"><b>' + esc(it.name) + '</b><span class="f"></span></button>' +
        '<button type="button" class="fe-del' + (askDel === i ? ' sure' : '') + '" aria-label="' + (askDel === i ? 'Wirklich löschen' : 'Löschen') + '">' + (askDel === i ? 'Löschen?' : '×') + '</button>';
      const use = row.querySelector('.fe-use');
      texToSvg(it.tex).then(h => { row.querySelector('.f').innerHTML = h; }).catch(() => { row.querySelector('.f').textContent = it.tex; });
      use.addEventListener('mousedown', e => e.preventDefault());
      use.addEventListener('click', () => {
        leaveLatex();
        mf.focus();
        mf.executeCommand(['insert', it.tex, {selectionMode: 'placeholder', format: 'latex'}]);
        document.dispatchEvent(new CustomEvent('fe-tool', {detail: {n: 'Meine Formel ' + it.name, ins: it.tex}}));
        show();
      });
      row.querySelector('.fe-del').addEventListener('click', () => {
        if (askDel !== i) { askDel = i; renderMine(); return; }      // zweiter Klick löscht wirklich
        const rest = mine(); rest.splice(i, 1); store.set('fe-meine', rest);
        askDel = -1; renderMine();
      });
      list.appendChild(row);
    });
    const name = grid.querySelector('#fe-mine-name');
    grid.querySelector('#fe-mine-add').addEventListener('click', () => {
      let tex = '';
      try { tex = mf.selectionIsCollapsed ? mf.getValue('latex') : mf.getValue(mf.selection, 'latex'); } catch (e) { tex = mf.getValue('latex'); }
      if (!tex.trim()) { msg.textContent = 'Die Eingabe ist leer. Erst eine Formel bauen.'; return; }
      const items2 = mine();
      const n = name.value.trim() || 'Formel ' + (items2.length + 1);
      items2.push({name: n, tex});
      store.set('fe-meine', items2);
      askDel = -1;
      renderMine();
      grid.querySelector('#fe-mine-msg').textContent = '„' + n + '“ gespeichert. Ein Klick darauf setzt sie an der Stelle des Cursors ein.';
    });
    name.addEventListener('keydown', e => { if (e.key === 'Enter') grid.querySelector('#fe-mine-add').click(); });
  }
  const origBuild = window.buildTools;
  window.buildTools = function (t) {
    if (t === 'meine') { askDel = -1; renderMine(); return; }
    origBuild(t);
    if (!input.value) showRecent();
  };
  window.FeBausteine = {search: v => { input.value = v; runSearch(); }};
})();
