# Mitgelieferte Bibliotheken und Schriften

Der Editor lädt in der Word-Fassung nichts mehr aus dem Netz. Alles, was er braucht, liegt hier.

| Ordner | Paket | Version | Lizenz |
|---|---|---|---|
| `mathjax/` | mathjax (nur Kern, TeX-Eingabe, SVG-Ausgabe) | 4.1.3 | Apache-2.0 |
| `mathjax-fonts/` | @mathjax/mathjax-{stix2,termes,newcm,pagella,schola,bonum,fira,dejavu}-font (nur SVG) | 4.1.3 | Apache-2.0, Schriften OFL-1.1 |
| `mathlive/` | mathlive | 0.111.0 | MIT |
| `jszip/` | jszip | 3.10.1 | MIT oder GPL-3.0 |
| `schrift/` | @fontsource/atkinson-hyperlegible | 5.3.0 | OFL-1.1 |

Zum Aktualisieren die Pakete mit npm holen und dieselben Dateien hierher kopieren. Danach `python3 src/build.py .` laufen lassen und mit `tests/pruefen.mjs` alle Beispielformeln prüfen.
