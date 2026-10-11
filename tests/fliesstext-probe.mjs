// Probedokument für Formeln im Fließtext: setzt einige Formeln mit dem echten Editor (wissenschaftlich,
// kompakt, einzeilig) und baut daraus eine Word-Datei mit Text in Times New Roman 12 pt. Darin sieht man
// in Word, ob die Formeln auf der Grundlinie sitzen und wie stark sie den Zeilenabstand vergrößern.
//
//   node fliesstext-probe.mjs Fliesstext-Probe.docx
import {chromium} from 'playwright';
import fs from 'fs';
import path from 'path';
import {createRequire} from 'module';
import {fileURLToPath} from 'url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.dirname(HERE);
const JSZip = createRequire(import.meta.url)('jszip');
const out = process.argv[2] || 'Fliesstext-Probe.docx';

const FORMELN = [
  ['E=mc^2', 'eine flache Formel'],
  ['\\frac{1}{2}mv^2', 'ein Bruch'],
  ['\\alpha=\\frac{e^2}{4\\uppi\\varepsilon_0\\hbar c}', 'ein Bruch mit Indizes'],
  ['\\int_0^{\\infty}e^{-x^2}\\,\\mathrm{d}x', 'ein Integral'],
  ['\\sqrt{\\frac{a}{b}}', 'eine Wurzel über einem Bruch'],
];
const MODI = [['sci', 'Wissenschaftlich'], ['kom', 'Kompakt'], ['lin', 'Einzeilig']];
const PT = 12;

const types = {'.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json', '.woff2': 'font/woff2'};
const launch = process.env.CHROMIUM ? {executablePath: process.env.CHROMIUM} : {};
const browser = await chromium.launch(launch);
const page = await browser.newPage();
await page.route('**/*', route => {
  const u = new URL(route.request().url());
  if (u.hostname !== 'editor.test') return route.abort();
  const f = path.join(REPO, decodeURIComponent(u.pathname));
  if (!fs.existsSync(f) || !fs.statSync(f).isFile()) return route.fulfill({status: 404, body: ''});
  route.fulfill({path: f, contentType: types[path.extname(f)] || 'application/octet-stream'});
});
await page.goto('https://editor.test/editor.html');
await page.waitForFunction(() => window.mf && document.getElementById('out').contentWindow.__mjReady === true, null, {timeout: 60000});

// Genau wie beim Einfügen im Add-in: Bild der Vorschau, 0,12 em Luft links und rechts im Fließtext
const bilder = {};
for (const [m] of MODI) for (const [tex] of FORMELN) {
  bilder[m + tex] = await page.evaluate(async ({tex, m, pt}) => {
    const w = document.getElementById('out').contentWindow;
    const disp = render(tex, m, true);
    w.__show(disp, disp, m);
    await w.__done;
    const r = await w.__png(pt, 600, 0.12);
    return {png: r.dataUrl.split(',')[1], svg: r.svg, wpt: r.wpt, hpt: r.hpt, dpt: r.dpt};
  }, {tex, m, pt: PT});
}
await browser.close();

// ---------- Word-Datei ----------
const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const zip = new JSZip();
let rels = '', n = 0;
const run = t => '<w:r><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
function bild(b) {
  n++;
  zip.file('word/media/f' + n + '.png', b.png, {base64: true});
  zip.file('word/media/f' + n + '.svg', b.svg);
  rels += '<Relationship Id="p' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/f' + n + '.png"/>' +
          '<Relationship Id="s' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/f' + n + '.svg"/>';
  const cx = Math.round(b.wpt * 12700), cy = Math.round(b.hpt * 12700);
  return '<w:r><w:rPr><w:position w:val="' + (-Math.round(b.dpt * 2)) + '"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + n + '" name="Formel ' + n + '"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="f.svg"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="p' + n + '"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="s' + n + '"/></a:ext></a:extLst></a:blip>' +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}
const para = (inner, extra) => '<w:p><w:pPr>' + (extra || '') + '</w:pPr>' + inner + '</w:p>';
const head = t => para(run(t), '<w:pStyle w:val="Heading2"/>');
const fill = 'The magnetic moment of the circulating charge follows directly from the current and the enclosed area, and the same reasoning applies to the orbiting electron in the hydrogen atom. ';
let body = head('Formeln im Fließtext: Probe') +
  para(run('Jeder Absatz unten enthält Fließtext in Times New Roman 12 pt mit einer Formel mittendrin. Bitte in Word ansehen: Sitzen die Formeln auf derselben Linie wie der Text? Wie stark weiten sie den Abstand der Zeile, in der sie stehen?'));
for (const [m, name] of MODI) {
  body += head('Darstellung „' + name + '“');
  for (const [tex, was] of FORMELN) {
    body += para(run(fill + 'Here is ' + was + ' ') + bild(bilder[m + tex]) + run(' in the middle of a line, followed by more text. ' + fill));
  }
}
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
zip.file('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '><w:docDefaults><w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="left"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="28"/></w:rPr></w:style></w:styles>');
zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="st" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + rels + '</Relationships>');
zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' + body +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>');
fs.writeFileSync(out, await zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'}));
console.log('Probedokument: ' + out);
