import json, os, sys
# Aufruf: python3 build.py [ZIELORDNER] [--samples samples.json]
HERE = os.path.dirname(os.path.abspath(__file__))
args = [a for a in sys.argv[1:]]
SAMPLES = 'samples.json'
if '--samples' in args:
    i = args.index('--samples'); SAMPLES = os.path.abspath(args[i + 1]); del args[i:i + 2]
REPO = os.path.abspath(args[0]) if args else None
os.chdir(HERE)
if REPO and os.path.abspath(REPO) == HERE:
    sys.exit('Zielordner darf nicht src/ sein')
s = open('index.src.html').read()
s = s.replace('/*@@MLFONTS@@*/', open('mlfonts.css').read())
s = s.replace('/*@@SAMPLES@@*/', json.dumps((json.load(open(SAMPLES)) if os.path.exists(SAMPLES) else []), ensure_ascii=False))
web = s.replace('/*@@ADDIN@@*/false', 'false')
open('formeleditor.html', 'w').write(web); print('artifact', len(s))
if REPO:
    body = s.replace('/*@@ADDIN@@*/false', 'true')
    title = '<title>Formeleditor für Papa</title>'
    assert title in body
    body = body.replace(title, '')
    css = '''<style>
body.addin header, body.addin section[aria-label="Beispiele"], body.addin #fallback,
body.addin .hint:has(#copypng) { display: none !important; }
body.addin .wrap { padding: 8px 8px 24px; gap: 10px; }
body.addin .panel { padding: 10px; gap: 10px; }
body.addin section[aria-label="Bausteine"] { order: 1; }
body.addin section[aria-label="Eingabe"] { order: 2; }
body.addin section[aria-label="Ergebnis"] { order: 3; }
body.addin section[aria-label="Darstellung"] { order: 4; }
body.addin .tools { grid-template-columns: repeat(auto-fill, minmax(62px, 1fr)); gap: 6px; max-height: 38vh; overflow-y: auto; padding-right: 2px; }
body.addin .tools.sym { grid-template-columns: repeat(auto-fill, minmax(46px, 1fr)); }
body.addin .tool, body.addin .tools.sym .tool { min-height: 54px; padding: 4px 2px; gap: 2px; }
body.addin .tool .g { font-size: 1.45rem; min-height: 1.7rem; }
body.addin .tool .n { font-size: .74rem; }
body.addin .row { gap: 6px; }
body.addin .row .btn { padding: 6px 10px; font-size: .92rem; }
body.addin .settings { grid-template-columns: minmax(0, 1fr); gap: 10px; }
body.addin .seg button { padding: 6px 10px; }
body.addin #fonts small { display: none; }
</style>'''
    page = ('<!doctype html>\n<html lang="de">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
            '<title>Formeleditor</title>\n'
            '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>\n'
            '</head>\n<body class="addin">\n' + body + css +
            '\n<script src="eqn3.js"></script>\n<script src="addin.js"></script>\n</body>\n</html>\n')
    open(os.path.join(REPO, 'taskpane.html'), 'w').write(page)
    for f in ('addin.js', 'eqn3.js'):
        open(os.path.join(REPO, f), 'w').write(open(f).read())
    print('taskpane', len(page))
    open(os.path.join(REPO, 'editor.html'), 'w').write('<!doctype html>\n<html lang="de">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n</head>\n<body>\n' + web + '\n</body>\n</html>\n')
