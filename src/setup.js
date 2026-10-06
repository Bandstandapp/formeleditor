/* Versionsanzeige und Einrichtung: vergleicht die laufende Version mit der auf GitHub Pages
   (version.json) und lädt den Editor auf Wunsch neu, damit Word die neuen Dateien holt. */
(function () {
  const local = window.FE_VERSION || {version: '0', datum: ''};
  const datum = d => d ? d.split('-').reverse().join('.') : '';
  const newer = (a, b) => {           // ist a neuer als b?
    const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number);
    for (let i = 0; i < Math.max(x.length, y.length); i++) {
      if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0);
    }
    return false;
  };

  const css = document.createElement('style');
  css.textContent = `
.fe-bar { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: .9rem; color: var(--muted); }
.fe-bar .btn { padding: 4px 10px; font-size: .9rem; }
.fe-news { background: var(--accent-soft); border: 2px solid var(--accent); border-radius: var(--r); padding: 8px 10px; display: grid; gap: 6px; }
.fe-news[hidden], .fe-dlg[hidden] { display: none; }
.fe-dlg { position: fixed; inset: 0; z-index: 50; background: rgba(0,0,0,.45); display: grid; place-items: start center; padding: 12px; overflow-y: auto; }
.fe-box { background: var(--surface); color: var(--fg); border: 1px solid var(--line); border-radius: var(--r); padding: 14px; width: 100%; max-width: 420px; display: grid; gap: 10px; }
.fe-box h2 { font-family: var(--f-head); font-size: var(--s-3); margin: 0; }
.fe-box table { border-collapse: collapse; width: 100%; }
.fe-box td { padding: 4px 0; vertical-align: top; }
.fe-box td:first-child { color: var(--muted); padding-right: 10px; white-space: nowrap; }
.fe-box .row { display: flex; flex-wrap: wrap; gap: 6px; }`;
  document.head.appendChild(css);

  const bar = document.createElement('div');
  bar.className = 'fe-bar';
  bar.innerHTML = '<span>Formeleditor, Version ' + local.version + '</span><button class="btn" type="button">Einrichtung</button>';
  const news = document.createElement('div');
  news.className = 'fe-news'; news.hidden = true; news.setAttribute('role', 'status');
  const dlg = document.createElement('div');
  dlg.className = 'fe-dlg'; dlg.hidden = true;
  dlg.innerHTML = `<div class="fe-box" role="dialog" aria-modal="true" aria-labelledby="fe-dlg-h">
  <h2 id="fe-dlg-h">Einrichtung</h2>
  <table>
    <tr><td>Diese Version</td><td id="fe-loc"></td></tr>
    <tr><td>Auf GitHub</td><td id="fe-rem">noch nicht geprüft</td></tr>
  </table>
  <p class="hint" id="fe-msg" role="status"></p>
  <div class="row" id="fe-ask" hidden>
    <button class="btn primary" type="button" id="fe-yes">Ja, jetzt aktualisieren</button>
    <button class="btn" type="button" id="fe-no">Nein, später</button>
  </div>
  <div class="row">
    <button class="btn" type="button" id="fe-check">Nach Updates suchen</button>
    <button class="btn" type="button" id="fe-reload">Editor neu laden</button>
    <button class="btn" type="button" id="fe-close">Schließen</button>
  </div>
</div>`;
  const wrap = document.querySelector('.wrap') || document.body;
  wrap.prepend(news);
  wrap.prepend(bar);
  document.body.appendChild(dlg);
  const $ = id => dlg.querySelector('#' + id);
  $('fe-loc').textContent = local.version + (local.datum ? ' vom ' + datum(local.datum) : '');

  let remote = null;
  async function fetchRemote() {
    const url = new URL('version.json', location.href);
    url.search = 't=' + Date.now();
    const r = await fetch(url, {cache: 'no-store'});
    if (!r.ok) throw new Error('GitHub antwortet nicht (' + r.status + ')');
    return r.json();
  }

  function update() {
    // Neu laden mit neuer Adresse, damit Word nichts aus dem Zwischenspeicher nimmt
    const url = new URL(location.href);
    url.search = 'v=' + encodeURIComponent(remote ? remote.version : local.version) + '&t=' + Date.now();
    location.replace(url.toString());
  }

  async function check(inDialog) {
    const msg = $('fe-msg');
    if (inDialog) { msg.textContent = 'Prüfe …'; $('fe-ask').hidden = true; }
    try {
      remote = await fetchRemote();
    } catch (e) {
      if (inDialog) { $('fe-rem').textContent = 'nicht erreichbar'; msg.textContent = 'Keine Verbindung zu GitHub. Bitte Internet prüfen. ' + e.message; }
      return;
    }
    $('fe-rem').textContent = remote.version + (remote.datum ? ' vom ' + datum(remote.datum) : '');
    if (newer(remote.version, local.version)) {
      const text = 'Es gibt eine neue Version ' + remote.version + '.' + (remote.neu ? ' Neu: ' + remote.neu : '');
      msg.textContent = text + ' Jetzt aktualisieren? Eine gerade eingegebene Formel geht dabei verloren.';
      $('fe-ask').hidden = false;
      news.innerHTML = '';
      const p = document.createElement('span'); p.textContent = 'Neue Version ' + remote.version + ' verfügbar.';
      const b = document.createElement('button'); b.className = 'btn'; b.type = 'button'; b.textContent = 'Ansehen';
      b.addEventListener('click', open);
      news.append(p, b); news.hidden = false;
    } else {
      news.hidden = true;
      if (inDialog) msg.textContent = 'Der Editor ist auf dem neuesten Stand.';
    }
  }

  function open() { dlg.hidden = false; check(true); $('fe-close').focus(); }
  function close() { dlg.hidden = true; }

  bar.querySelector('button').addEventListener('click', open);
  $('fe-close').addEventListener('click', close);
  $('fe-no').addEventListener('click', () => { $('fe-ask').hidden = true; close(); });
  $('fe-yes').addEventListener('click', update);
  $('fe-reload').addEventListener('click', update);
  $('fe-check').addEventListener('click', () => check(true));
  dlg.addEventListener('click', e => { if (e.target === dlg) close(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !dlg.hidden) close(); });

  window.FeSetup = {newer, check};
  check(false);   // stille Prüfung beim Start: zeigt nur einen Hinweis, wenn es etwas Neues gibt
})();
