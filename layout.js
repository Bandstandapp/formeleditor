/* Anordnung im Word-Add-in: Eingabe und Einfügen oben, alle Einstellungen zum Aussehen an einer Stelle,
   eine Zeile darunter zeigt, mit welchen Einstellungen eingefügt wird. Nur Umstellen, keine neue Logik. */
(function () {
  const q = s => document.querySelector(s);
  const eingabe = q('section[aria-label="Eingabe"]');
  const aussehen = q('section[aria-label="Darstellung"]');
  const ergebnis = q('section[aria-label="Ergebnis"]');

  const css = document.createElement('style');
  css.textContent = `
body.addin .tabs { flex-wrap: nowrap; gap: 6px; }
body.addin .tab { flex: 1 1 0; padding: 8px 6px; font-size: 1rem; text-align: center; }
body.addin .editrow { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; }
body.addin .editrow .btn { padding: 6px 2px; font-size: .85rem; white-space: nowrap; }
body.addin .editrow .btn.wide { grid-column: 1 / -1; }
body.addin .row .btn.primary, body.addin .copyrow .btn.primary { width: 100%; font-size: 1.15rem; padding: 12px 14px; }
body.addin #sizes { flex-wrap: nowrap; gap: 4px; }
body.addin #sizes button { flex: 1 1 0; min-width: 0; padding: 8px 0; text-align: center; font-size: .85rem; white-space: nowrap; }
body.addin .zoomrow { margin: 0; gap: 8px; font-weight: 400; font-size: .95rem; }
body.addin .copyrow { display: block; }
body.addin #mfwrap math-field { min-height: 80px; }
body.addin .paper iframe { min-height: 0; }
body.addin #fe-summary { font-size: .92rem; display: flex; flex-wrap: wrap; gap: 4px 8px; align-items: baseline; }
body.addin #fe-summary b { color: var(--fg); }
body.addin details.tipps summary { font-size: .92rem; }
body.addin details.tipps p { margin-top: 6px; }
#fe-report[hidden] { display: none !important; }
body.addin details.ganz > summary { font-size: var(--s-0); font-weight: 700; color: var(--muted); letter-spacing: .04em; text-transform: uppercase; }
body.addin details.ganz[open] { display: grid; gap: 10px; }`;
  document.head.appendChild(css);

  // Schriftgröße zu Darstellung und Schrift: alles zum Aussehen an einer Stelle
  const sizes = q('.sizes');
  if (sizes && aussehen) {
    const box = document.createElement('div');
    box.innerHTML = '<p class="label">Schriftgröße im Dokument</p>';
    box.appendChild(q('#sizes'));
    q('section[aria-label="Darstellung"] .settings').prepend(box);
    sizes.remove();
  }
  const aussLabel = document.createElement('p');
  aussLabel.className = 'label'; aussLabel.style.fontSize = '1.05rem';
  aussLabel.textContent = 'Aussehen der Formel';
  aussehen.prepend(aussLabel);

  // Unter dem oberen Einfügen-Knopf: mit welchen Einstellungen eingefügt wird
  const summary = document.createElement('p');
  summary.className = 'hint'; summary.id = 'fe-summary';
  const firstText = el => el ? Array.from(el.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join('').trim() : '';
  function updateSummary() {
    const pt = q('#sizes button[aria-pressed="true"]');
    const m = q('#mode button[aria-pressed="true"]');
    const f = q('#fonts button[aria-pressed="true"]');
    summary.innerHTML = '';
    const t = document.createElement('span');
    t.innerHTML = 'Einfügen mit <b></b>';
    t.querySelector('b').textContent = [pt && pt.textContent.trim(), firstText(m), firstText(f)].filter(Boolean).join(' · ');
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'linkbtn'; b.textContent = 'ändern';
    b.addEventListener('click', () => aussehen.scrollIntoView({behavior: 'smooth', block: 'start'}));
    summary.append(t, b);
  }
  const btn2 = q('#copy2');
  const status2 = btn2 && btn2.parentElement.nextElementSibling;
  (status2 || btn2.parentElement).after(summary);
  updateSummary();
  new MutationObserver(updateSummary).observe(aussehen, {subtree: true, attributes: true, attributeFilter: ['aria-pressed'], childList: true});

  // Bearbeiten-Knöpfe kompakt in einer Reihe, „Markierte Formel bearbeiten“ darunter
  const editRow = eingabe.querySelector('.row');
  editRow.classList.add('editrow');
  const fromWord = editRow.querySelector('#fromword');
  if (fromWord) fromWord.classList.add('wide');

  // Reihenfolge in der Eingabe: Feld, Bearbeiten-Knöpfe, großer Einfügen-Knopf, dann Hinweise
  const tipp = Array.from(eingabe.querySelectorAll('.hint')).find(p => /^Tipp/.test(p.textContent));
  if (tipp) {
    const d = document.createElement('details');
    d.className = 'tipps';
    d.innerHTML = '<summary>Tipps zum Tippen</summary>';
    tipp.textContent = tipp.textContent.replace(/^Tipp:\s*/, '');
    tipp.replaceWith(d);
    d.appendChild(tipp);
  }

  // „Ganzes Dokument“ ist selten nötig: eingeklappt
  const ganz = q('section[aria-label="Dokument"]');
  if (ganz) {
    const d = document.createElement('details');
    d.className = 'ganz';
    const sum = document.createElement('summary');
    sum.textContent = 'Ganzes Dokument umwandeln';
    d.appendChild(sum);
    const lbl = ganz.querySelector('.label');
    if (lbl) lbl.remove();
    while (ganz.firstChild) d.appendChild(ganz.firstChild);
    ganz.appendChild(d);
  }

  // Vorschau-Überschrift kürzer, Quelltext für Jens kommt in die Einrichtung (setup.js)
  const lbl = ergebnis.querySelector('.label');
  if (lbl) lbl.textContent = 'Vorschau';
})();
