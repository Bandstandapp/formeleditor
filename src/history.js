/* Verlauf für „Rückgängig“ und „Wiederholen“: Ein kleiner Pfeil neben jedem Knopf öffnet die letzten
   20 Schritte mit Häkchen. Ausgewählte Schritte lassen sich einzeln rückgängig machen oder auf die
   Formel in der Eingabe anwenden, auch auf eine andere, später geladene Formel.
   Ein Schritt merkt sich nur die geänderte Stelle samt etwas Umgebung. Beim Anwenden wird diese
   Stelle in der aktuellen Formel gesucht; passt sie nicht, wird reiner neuer Inhalt an der
   Cursorposition eingefügt. */
(function () {
  const MAX = 20;
  const q = s => document.querySelector(s);
  const undoBtn = q('#undo'), redoBtn = q('#redo');
  if (!undoBtn || !redoBtn) return;

  let steps = [];             // {label, kind: 'tex'|'mode'|'font'|'size', before, after, undone, t}
  let lastTex = null;         // Stand der Eingabe nach dem letzten erfassten Schritt
  let quiet = false;          // eigene Änderungen nicht als neuen Schritt erfassen
  let nextLabel = null;       // Name des Bausteins bzw. Knopfs, der die nächste Änderung auslöst

  const tex = () => (window.mf && mf.getValue ? mf.getValue('latex') : '');
  const short = t => { t = t.split('\\placeholder{}').join('□').replace(/\\textcolor\{[^{}]*\}\{([^{}]*)\}/g, '$1').replace(/\\(mathbf|mathrm|mathit|boldsymbol|bm)\{([^{}]*)\}/g, '$2').replace(/\s+/g, ' ').trim(); return t.length > 22 ? t.slice(0, 21) + '…' : t; };
  const fontName = id => { const f = (typeof FONTS !== 'undefined' ? FONTS : []).find(f => f.id === id); return f ? f.name : id; };
  const valueText = (kind, v) => kind === 'mode' ? (v === 'lin' ? 'einzeilig' : 'wissenschaftlich')
    : kind === 'font' ? fontName(v) : String(v).replace('.', ',') + ' pt';

  // Geänderte Stelle zwischen zwei Ständen: gemeinsamer Anfang und gemeinsames Ende bleiben stehen
  function diff(a, b) {
    let p = 0;
    while (p < a.length && p < b.length && a[p] === b[p]) p++;
    let s = 0;
    while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
    return {left: a.slice(0, p), removed: a.slice(p, a.length - s), inserted: b.slice(p, b.length - s), right: a.slice(a.length - s), pos: a.length ? p / a.length : 0};
  }
  // Wendet die Änderung a→b auf cur an. Ergebnis: neuer Text, 'cursor' (an der Cursorposition einfügen) oder null
  const PH = '\\placeholder{}';
  const esc = t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const FILL = '((?:[^{}]|\\{(?:[^{}]|\\{[^{}]*\\})*\\})*?)';   // Inhalt eines inzwischen ausgefüllten Platzhalters
  function patch(cur, a, b) {
    const d = diff(a, b);
    for (const n of [Infinity, 24, 12, 6, 3, 1, 0]) {
      const L = n === Infinity ? d.left : d.left.slice(Math.max(0, d.left.length - n)), R = n === Infinity ? d.right : d.right.slice(0, n);
      if (!(L + d.removed + R)) break;
      // Platzhalter im zu ersetzenden Teil dürfen inzwischen ausgefüllt sein; ihr Inhalt bleibt erhalten
      const re = new RegExp(esc(L) + d.removed.split(PH).map(esc).join(FILL) + esc(R), 'g');
      let best = null, bestDist = Infinity;
      for (const m of cur.matchAll(re)) {
        const dist = Math.abs((cur.length ? (m.index + L.length) / cur.length : 0) - d.pos);
        if (dist < bestDist) { best = m; bestDist = dist; }
        if (!m[0].length) re.lastIndex++;
      }
      if (best) {
        const fills = best.slice(1).filter(f => f && f !== PH).join('');
        return cur.slice(0, best.index) + L + d.inserted + fills + R + cur.slice(best.index + best[0].length);
      }
    }
    return d.removed ? null : (d.inserted ? 'cursor' : cur);
  }

  function texLabel(a, b) {
    const d = diff(a, b);
    d.removed = d.removed.split(PH).join('');
    if (nextLabel) return nextLabel + (/^(Fett|Kursiv|Aufrecht|Farbe)/.test(nextLabel) && d.removed ? ' „' + short(d.removed) + '“' : '');
    if (!d.removed) return 'Eingabe „' + short(d.inserted) + '“';
    if (!d.inserted) return 'Gelöscht „' + short(d.removed) + '“';
    return 'Geändert „' + short(d.removed) + '“ → „' + short(d.inserted) + '“';
  }
  function add(step) {
    step.t = Date.now(); step.undone = false;
    steps.push(step);
    if (steps.length > MAX) steps.shift();
    refresh();
  }
  function onInput() {
    const cur = tex();
    if (quiet || lastTex === null) { lastTex = cur; return; }
    if (cur === lastTex) return;
    const top = steps[steps.length - 1];
    const typing = !nextLabel && top && top.kind === 'tex' && !top.undone && top.typing && top.after === lastTex && Date.now() - top.t < 1500;
    if (typing) {                                // schnell hintereinander getippte Zeichen sind ein Schritt
      top.after = cur; top.t = Date.now();
      top.label = texLabel(top.before, cur);
      refresh();
    } else {
      add({kind: 'tex', before: lastTex, after: cur, label: texLabel(lastTex, cur), typing: !nextLabel});
    }
    nextLabel = null;
    lastTex = cur;
  }
  // Änderungen ohne Eingabe-Ereignis (Formel aus Word geladen) übernehmen, ohne sie zu erfassen
  setInterval(() => { if (window.mf && mf.getValue && !quiet) { const cur = tex(); if (lastTex === null || (cur !== lastTex && !mfBusy())) lastTex = cur; } }, 400);
  let inputAt = 0;
  const mfBusy = () => Date.now() - inputAt < 50;

  const hook = () => {
    if (!(window.mf && mf.addEventListener && mf.getValue)) return setTimeout(hook, 200);
    lastTex = tex();
    mf.addEventListener('input', () => { inputAt = Date.now(); onInput(); });
  };
  hook();

  // Namen der Bausteine und der Knöpfe F und K für die Liste merken
  document.addEventListener('click', e => {
    const tool = e.target.closest('#tools .tool');
    if (tool) { const n = tool.querySelector('.n'); nextLabel = 'Baustein ' + (n ? n.textContent : tool.getAttribute('aria-label') || ''); return; }
    const st = e.target.closest('.stylebar button');
    if (e.target.closest('.fe-swatch')) { nextLabel = 'Farbe ' + e.target.closest('.fe-swatch').title; return; }
    if (st && st.classList.contains('c')) return;
    if (st) nextLabel = st.classList.contains('b') ? 'Fett' : (st.getAttribute('aria-pressed') === 'true' ? 'Aufrecht' : 'Kursiv');
    const mb = e.target.closest('#mode button');
    if (mb && !quiet) { const before = mode; setTimeout(() => { if (mode !== before) add({kind: 'mode', before, after: mode, label: 'Darstellung: ' + valueText('mode', mode)}); }, 0); }
  }, true);
  document.addEventListener('change', e => {
    if (quiet) return;   // gilt auch für Änderungen, die run() selbst auslöst
    const kind = e.target.id === 'fe-font' ? 'font' : e.target.id === 'fe-size' ? 'size' : null;
    if (!kind) return;
    const before = kind === 'font' ? font : sizePt;
    setTimeout(() => {
      const after = kind === 'font' ? font : sizePt;
      if (after !== before) add({kind, before, after, label: (kind === 'font' ? 'Schrift: ' : 'Größe: ') + valueText(kind, after)});
    }, 0);
  }, true);

  // Einzelne Werte zurücksetzen bzw. setzen (Darstellung, Schrift, Größe)
  function setValue(kind, v) {
    if (kind === 'mode') { const b = q('#mode button[data-mode="' + v + '"]'); if (b) b.click(); }
    else if (kind === 'font') { const b = q('#font-' + v); if (b) b.click(); }
    else { const b = q('#size-' + v); if (b) b.click(); else { sizePt = v; if (window.syncSeg) syncSeg(); } }
  }
  // Wendet ausgewählte Schritte an (forward) oder macht sie rückgängig; liefert die Zahl der nicht passenden
  function run(list, forward, label) {
    let cur = tex(), misses = 0, cursorIns = [];
    quiet = true;
    try {
      list.forEach(s => {
        if (s.kind === 'tex') {
          const r = forward ? patch(cur, s.before, s.after) : patch(cur, s.after, s.before);
          if (r === null) { misses++; return; }
          if (r === 'cursor') cursorIns.push(diff(forward ? s.before : s.after, forward ? s.after : s.before).inserted);
          else cur = r;
        } else setValue(s.kind, forward ? s.after : s.before);
        s.undone = !forward;
      });
      const was = tex();
      if (cur !== was) mf.setValue(cur, {silenceNotifications: true});
      cursorIns.forEach(t => mf.insert(t, {format: 'latex', selectionMode: 'placeholder'}));
      // Übertragene Schritte werden selbst ein Schritt, damit „Rückgängig“ sie wieder zurücknimmt
      if (forward && label && tex() !== was) setTimeout(() => add({kind: 'tex', before: was, after: tex(), label}), 0);
    } finally {
      lastTex = tex();
      quiet = false;
    }
    if (window.show) show();
    return misses;
  }

  // Oberfläche: Pfeil neben den Knöpfen und eine aufklappende Liste darunter
  const css = document.createElement('style');
  css.textContent = `
.fe-split { display: flex; }
.fe-split > .btn:first-child { flex: 1; border-top-right-radius: 0; border-bottom-right-radius: 0; }
.fe-split > .fe-caret { flex: 0 0 22px; padding: 0 !important; border-left: 0; border-top-left-radius: 0; border-bottom-left-radius: 0; font-size: .8rem; }
.fe-split > .fe-caret[aria-expanded="true"] { background: var(--accent); color: #fff; }
.fe-hist { margin-top: 8px; border: 1.5px solid var(--line); border-radius: var(--r); background: var(--surface); padding: 8px 10px; font-family: var(--f-ui); }
.fe-hist h3 { margin: 0 0 6px; font-size: .9rem; }
.fe-hist ol { list-style: none; margin: 0 0 8px; padding: 0; max-height: 260px; overflow-y: auto; }
.fe-hist li label { display: flex; gap: 8px; align-items: center; padding: 4px 2px; font-size: .88rem; cursor: pointer; border-radius: 6px; }
.fe-hist li label:hover { background: var(--line); }
.fe-hist li input { flex: none; width: 16px; height: 16px; }
.fe-hist li .off { color: var(--muted, #6b7280); font-size: .78rem; margin-left: auto; white-space: nowrap; }
.fe-hist .fe-hrow { display: flex; gap: 6px; flex-wrap: wrap; }
.fe-hist .fe-hrow .btn.btn { flex: 1 1 auto; width: auto; min-height: 0; padding: 7px 10px; font-size: .88rem; }
.fe-hist .fe-hmsg { margin: 6px 0 0; font-size: .82rem; }
.fe-hist .fe-empty { font-size: .85rem; color: var(--muted, #6b7280); margin: 2px 0 8px; }`;
  document.head.appendChild(css);

  function split(btn, title) {
    const wrap = document.createElement('div');
    wrap.className = 'fe-split';
    btn.replaceWith(wrap);
    wrap.appendChild(btn);
    const c = document.createElement('button');
    c.type = 'button'; c.className = 'btn fe-caret'; c.textContent = '▾';
    c.title = c.ariaLabel = title;
    c.setAttribute('aria-expanded', 'false');
    wrap.appendChild(c);
    return c;
  }
  const undoCaret = split(undoBtn, 'Letzte Schritte zum Rückgängigmachen anzeigen');
  const redoCaret = split(redoBtn, 'Letzte Schritte zum Wiederholen und Übertragen anzeigen');
  const panel = document.createElement('div');
  panel.className = 'fe-hist';
  panel.hidden = true;
  (undoCaret.closest('.row') || undoCaret.parentNode).after(panel);
  let open = null;            // 'undo' | 'redo' | null
  let lastMsg = null;         // Rückmeldung nach dem letzten Klick, bleibt beim Neuzeichnen stehen

  function render() {
    if (!open) return;
    const forward = open === 'redo';
    const list = steps.map((s, i) => ({s, i})).filter(x => forward || !x.s.undone).reverse();   // neueste oben
    panel.innerHTML = '<h3>' + (forward ? 'Schritte wiederholen oder übertragen' : 'Schritte rückgängig machen') + '</h3>' +
      (list.length ? '<ol>' + list.map(x => '<li><label><input type="checkbox" data-i="' + x.i + '"><span></span>' +
        (forward && x.s.undone ? '<span class="off">rückgängig gemacht</span>' : '') + '</label></li>').join('') + '</ol>'
        : '<p class="fe-empty">Noch keine Schritte.</p>') +
      '<div class="fe-hrow"><button type="button" class="btn primary fe-go">' + (forward ? 'Anwenden' : 'Rückgängig machen') + '</button>' +
      '<button type="button" class="btn fe-close">Schließen</button></div><p class="fe-hmsg hint" role="status"></p>';
    if (lastMsg) { const msg = panel.querySelector('.fe-hmsg'); msg.textContent = lastMsg.text; msg.classList.toggle('ok', lastMsg.ok); }
    panel.querySelectorAll('li span:not(.off)').forEach((sp, k) => { sp.textContent = list[k].s.label; });
    const go = panel.querySelector('.fe-go');
    const sync = () => { go.disabled = !panel.querySelector('input:checked'); };
    panel.querySelectorAll('input').forEach(c => c.addEventListener('change', sync));
    sync();
    go.addEventListener('click', () => {
      const idx = [...panel.querySelectorAll('input:checked')].map(c => +c.dataset.i);
      const chosen = idx.map(i => steps[i]).filter(Boolean);
      chosen.sort((a, b) => forward ? a.t - b.t : b.t - a.t);   // anwenden: älteste zuerst; zurück: neueste zuerst
      lastMsg = null;
      const misses = run(chosen, forward, forward && !chosen.every(s => s.undone) ? 'Übertragen: ' + chosen.map(s => s.label).join(', ') : null);
      lastMsg = {ok: !misses, text: misses ? (misses === 1 ? 'Ein Schritt passte nicht zu dieser Formel und wurde ausgelassen.' : misses + ' Schritte passten nicht zu dieser Formel und wurden ausgelassen.')
        : (forward ? 'Erledigt.' : 'Rückgängig gemacht.')};
      render();
    });
    panel.querySelector('.fe-close').addEventListener('click', () => toggle(null));
  }
  function toggle(which) {
    lastMsg = null;
    open = open === which ? null : which;
    panel.hidden = !open;
    undoCaret.setAttribute('aria-expanded', open === 'undo');
    redoCaret.setAttribute('aria-expanded', open === 'redo');
    render();
  }
  function refresh() { if (open) render(); }
  undoCaret.addEventListener('click', () => toggle('undo'));
  redoCaret.addEventListener('click', () => toggle('redo'));

  // Die normalen Knöpfe gehen weiter einen Schritt zurück bzw. vor und halten die Liste stimmig
  undoBtn.addEventListener('click', e => {
    e.stopImmediatePropagation();
    const s = [...steps].reverse().find(s => !s.undone);
    if (s) { run([s], false); refresh(); } else { mf.executeCommand('undo'); if (window.show) show(); lastTex = tex(); }
  }, true);
  redoBtn.addEventListener('click', e => {
    e.stopImmediatePropagation();
    const s = steps.find(s => s.undone && s.t >= Math.max(0, ...steps.filter(x => !x.undone).map(x => x.t)));
    if (s) { run([s], true); refresh(); } else { mf.executeCommand('redo'); if (window.show) show(); lastTex = tex(); }
  }, true);

  window.FeHistory = {steps: () => steps, patch};
})();
