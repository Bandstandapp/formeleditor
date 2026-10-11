/* „Problem melden“ in der Fußzeile: packt eine kurze Notiz, die Formel aus der Eingabe (als Link, der
   sie in der Web-Fassung des Editors öffnet), die letzten Meldungen des Editors und die Version in eine
   E-Mail. Die Empfänger-Adresse wird einmal eingetragen und auf diesem Mac gemerkt. */
(function () {
  const q = s => document.querySelector(s);
  const bar = q('.fe-bar');
  if (!bar) return;
  const local = window.FE_VERSION || {version: '?', datum: ''};
  const css = document.createElement('style');
  css.textContent = `
.fe-bar { flex-wrap: wrap; row-gap: 0; }
.fe-bar > span { flex: 1 1 100%; }
.fe-bar .fe-report { all: unset; cursor: pointer; font-size: .95rem; color: var(--accent); text-decoration: underline; padding: 4px 2px; margin-left: auto; margin-right: 6px; white-space: nowrap; }
.fe-bar .fe-report:focus-visible { outline: 3px solid var(--focus); }
.fe-box label.fe-field { display: grid; gap: 3px; font-size: .95rem; }
.fe-box textarea, .fe-box input[type=email] { font: inherit; font-size: 1rem; padding: 6px 8px; border: 2px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--fg); width: 100%; }
.fe-box textarea { min-height: 90px; resize: vertical; }
.fe-box pre { margin: 0; white-space: pre-wrap; word-break: break-word; font-size: .8rem; background: var(--bg); padding: 8px; border-radius: 6px; max-height: 160px; overflow: auto; }`;
  document.head.appendChild(css);

  const btn = document.createElement('button');
  btn.type = 'button'; btn.className = 'fe-report'; btn.textContent = 'Problem melden';
  bar.insertBefore(btn, bar.querySelector('.fe-gear'));

  const dlg = document.createElement('div');
  dlg.className = 'fe-dlg'; dlg.hidden = true;
  dlg.innerHTML = `<div class="fe-box" role="dialog" aria-modal="true" aria-labelledby="fe-rep-h">
  <h2 id="fe-rep-h">Problem melden</h2>
  <label class="fe-field">Was stimmt nicht?<textarea id="fe-rep-note" placeholder="Zum Beispiel: Das Integral sieht anders aus als im alten Dokument."></textarea></label>
  <label class="fe-field">An (E-Mail-Adresse)<input type="email" id="fe-rep-to" autocomplete="off"></label>
  <details><summary>Das wird mitgeschickt</summary><pre id="fe-rep-pre"></pre></details>
  <p class="hint" id="fe-rep-msg" role="status"></p>
  <div class="row">
    <button class="btn primary" type="button" id="fe-rep-mail">E-Mail vorbereiten</button>
    <button class="btn" type="button" id="fe-rep-copy">Text kopieren</button>
    <button class="btn" type="button" id="fe-rep-close">Schließen</button>
  </div>
</div>`;
  document.body.appendChild(dlg);
  const $ = id => dlg.querySelector('#' + id);
  try { $('fe-rep-to').value = localStorage.getItem('fe-melden-an') || ''; } catch (e) {}
  $('fe-rep-to').addEventListener('change', () => { try { localStorage.setItem('fe-melden-an', $('fe-rep-to').value.trim()); } catch (e) {} });

  const b64url = t => { const u = new TextEncoder().encode(t); let s = ''; u.forEach(c => { s += String.fromCharCode(c); }); return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
  const modeName = m => m === 'lin' ? 'einzeilig' : m === 'kom' ? 'kompakt' : 'wissenschaftlich';
  const fontName = id => (FONTS.find(f => f.id === id) || {}).name || id;
  const msgs = () => ['pastemsg', 'copymsg', 'fe-allmsg', 'fe-findmsg', 'fe-anamsg', 'fe-num-msg', 'fe-align-msg']
    .map(id => { const el = document.getElementById(id); return el && el.textContent.trim(); }).filter(Boolean);
  function details() {
    const tex = window.mf ? mf.getValue('latex') : '';
    const lines = [];
    if (tex.trim()) {
      const link = new URL('editor.html', location.href);
      link.search = '';
      link.hash = 'f=' + b64url(JSON.stringify({tex, mode, font, pt: sizePt}));
      lines.push('Formel in der Eingabe (der Link zeigt sie im Browser):', link.href, '', 'Quelltext: ' + tex,
        'Darstellung: ' + modeName(mode) + ', Schrift ' + fontName(font) + ', ' + String(sizePt).replace('.', ',') + ' pt');
    } else lines.push('Die Eingabe war leer.');
    const m = msgs();
    if (m.length) lines.push('', 'Letzte Meldungen des Editors:', ...m.map(x => '- ' + x));
    let word = '';
    try { const d = Office.context.diagnostics; word = d ? ', Word ' + d.version + ' (' + d.platform + ')' : ''; } catch (e) {}
    lines.push('', 'Papas Formeleditor, Version ' + local.version + (local.datum ? ' vom ' + local.datum.split('-').reverse().join('.') : '') + word);
    return lines.join('\n');
  }
  function fullText() {
    const note = $('fe-rep-note').value.trim();
    return 'Was ist passiert:\n' + (note || '(keine Beschreibung)') + '\n\n' + details();
  }
  function open() {
    dlg.hidden = false; $('fe-rep-msg').textContent = '';
    $('fe-rep-pre').textContent = details();
    $('fe-rep-note').focus();
  }
  function close() { dlg.hidden = true; }
  btn.addEventListener('click', open);
  $('fe-rep-close').addEventListener('click', close);
  dlg.addEventListener('click', e => { if (e.target === dlg) close(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !dlg.hidden) close(); });

  $('fe-rep-mail').addEventListener('click', () => {
    const to = $('fe-rep-to').value.trim();
    try { localStorage.setItem('fe-melden-an', to); } catch (e) {}
    let body = fullText();
    // Mail-Programme kappen sehr lange Links; dann bleibt der Quelltext, der Link entfällt
    if (body.length > 1800) body = body.replace(/\nFormel in der Eingabe[^\n]*\n[^\n]*\n/, '\n');
    if (body.length > 1800) body = body.slice(0, 1790) + ' …';
    const url = 'mailto:' + encodeURIComponent(to).replace(/%40/g, '@') + '?subject=' + encodeURIComponent('Formeleditor: Problem') + '&body=' + encodeURIComponent(body);
    try { Office.context.ui.openBrowserWindow(url); }
    catch (e) { window.open(url, '_blank'); }
    $('fe-rep-msg').textContent = 'Das E-Mail-Programm sollte sich jetzt mit der fertigen Nachricht öffnen. Bitte dort auf „Senden“ klicken. ' +
      'Passiert nichts, „Text kopieren“ drücken und den Text in eine neue E-Mail einfügen (⌘V).';
  });
  $('fe-rep-copy').addEventListener('click', async () => {
    const text = fullText();
    const ok = () => { $('fe-rep-msg').textContent = 'Kopiert. Jetzt eine neue E-Mail öffnen und mit ⌘V einfügen.'; };
    try { await navigator.clipboard.writeText(text); ok(); return; } catch (e) {}
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy') ? ok() : ($('fe-rep-msg').textContent = 'Kopieren ging nicht.'); } catch (e) { $('fe-rep-msg').textContent = 'Kopieren ging nicht.'; }
    ta.remove();
  });
})();
