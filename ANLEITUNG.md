# Papas Formeleditor: Einrichtung und Bedienung (Mac)

Diese Anleitung richtet den Formeleditor als Add-in in Word auf dem Mac ein. Sie ist so geschrieben, dass Jens sie über TeamViewer auf Papas Rechner durchgehen kann. Die Einrichtung dauert ein paar Minuten und muss nur einmal gemacht werden. Danach steht der Editor in Word im Reiter „Start“ als Knopf „Papas Formeleditor“ bereit.

Der Editor selbst liegt im Internet auf GitHub Pages. Word lädt ihn bei jedem Öffnen von dort. Verbesserungen kommen deshalb ohne neue Installation an; manchmal reicht es, Word einmal zu beenden und neu zu starten.

## Voraussetzungen

- Ein Mac mit Microsoft Word aus Microsoft 365 (aktuelle Version).
- Eine Internetverbindung, solange der Editor benutzt wird.
- Word sollte während der Einrichtung **beendet** sein (⌘Q).

## Schritt 1: Die Add-in-Datei an die richtige Stelle legen

Word sucht Add-ins zum Ausprobieren in einem bestimmten Ordner. Am einfachsten legt man die Datei mit dem Terminal dorthin, weil der Ordner im Finder versteckt ist.

1. Das Programm **Terminal** öffnen: ⌘ und Leertaste drücken, „Terminal“ tippen, Eingabetaste.
2. Die folgende Zeile vollständig kopieren, im Terminal mit ⌘V einfügen und die Eingabetaste drücken:

   ```
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef && curl -fsSL https://bandstandapp.github.io/formeleditor/manifest.xml -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/formeleditor.xml && echo "Fertig"
   ```

3. Wenn im Terminal „Fertig“ erscheint, hat es geklappt. Das Terminal kann man danach schließen.

Die Zeile legt den Ordner an, falls er fehlt, und lädt die kleine Beschreibungsdatei des Add-ins hinein. Sie verändert sonst nichts am Rechner.

**Ohne Terminal** geht es auch über den Finder: die Datei [manifest.xml](https://bandstandapp.github.io/formeleditor/manifest.xml) herunterladen, im Finder „Gehe zu“ → „Gehe zum Ordner …“ (⇧⌘G) wählen, `~/Library/Containers/com.microsoft.Word/Data/Documents/` eingeben, dort einen Ordner `wef` anlegen, falls es ihn noch nicht gibt, und die Datei hineinziehen.

## Schritt 2: Das Add-in in Word einschalten

1. Word starten und ein Dokument öffnen (ein neues, leeres reicht).
2. Im Reiter **Start** ganz rechts auf **Add-ins** klicken. In manchen Word-Versionen heißt der Weg **Einfügen** → **Add-ins** → **Meine Add-ins**.
3. Im Fenster, das sich öffnet, erscheint unter **Entwickler-Add-ins** (englisch „Developer Add-ins“) der Eintrag **Papas Formeleditor**. Ihn einmal anklicken.
4. Rechts öffnet sich der Formeleditor als Seitenleiste. Im Reiter **Start** gibt es von nun an außerdem den Knopf **Papas Formeleditor**.

Falls der Eintrag nicht erscheint, Word mit ⌘Q ganz beenden und neu starten. Hilft das nicht, ist die Datei aus Schritt 1 nicht im richtigen Ordner gelandet.

## So wird der Editor benutzt

Beim Öffnen ist die Eingabe leer. Der Editor hat oben zwei Reiter. **Formel bearbeiten** ist beim Öffnen gewählt und enthält die Kacheln Bausteine, Eingabe und Vorschau. **Alles bearbeiten** enthält die Werkzeuge für das ganze Dokument. Hinter jeder Überschrift erklärt ein kleines **i** die Kachel, sobald man mit der Maus darauf zeigt.

- **Neue Formel:** Den Cursor im Text an die gewünschte Stelle setzen. Die Formel aus den Bausteinen zusammensetzen oder tippen und auf **In Word einfügen** klicken. Den Knopf gibt es unter der Eingabe und unter der Vorschau.
- **Fett, kursiv und Farbe:** Zeichen in der Eingabe markieren und auf **F** (fett), **K** (kursiv bzw. aufrecht) oder **A** (Farbe) klicken. Ohne Markierung gilt die Wahl für die nächsten Zeichen.
- **Zahlen mit Abständen:** Die Leertaste setzt einen kleinen Abstand, etwa in 1.159 652 180 46. Nach einem Komma entsteht kein Abstand, also 1,5 und nicht 1, 5.
- **Aussehen:** In der Vorschau zwischen **Wissenschaftlich** (Brüche übereinander, für abgesetzte Gleichungen), **Kompakt** (Brüche übereinander, aber flacher, für den Fließtext) und **Einzeilig** (Brüche mit Schrägstrich) umschalten und Schrift und Schriftgröße wählen. Der Zoom vergrößert nur die Vorschau.
- **Formel ändern:** Die Formel im Dokument einmal anklicken. Sie erscheint in der Eingabe. Solange nichts geändert ist, heißt der große Knopf **Ohne Änderung übernehmen**; er wandelt eine alte Formel unverändert ins neue Format um. Nach einer Änderung heißt er **Formel in Word ersetzen**. Ein grüner Rahmen um die Eingabe zeigt: Die Formel ist so wie im Dokument. Nach einer Änderung wird der Rahmen orange, die Änderung ist also noch nicht in Word. Nach dem Klick auf den Knopf ist er wieder grün. **Abbrechen** verwirft alle Änderungen und stellt die Formel wieder her, wie sie im Dokument steht.
- **Verlauf:** Der kleine Pfeil neben **Rückgängig** zeigt die letzten 20 Schritte. Ausgewählte Schritte werden mit **Rückgängig machen** einzeln zurückgenommen, auch wenn danach noch weitergearbeitet wurde. Der Pfeil neben **Wiederholen** zeigt dieselben Schritte. Mit **Anwenden** kommen ausgewählte Schritte zurück, oder sie werden auf eine andere Formel übertragen, zum Beispiel „v fett“ auf die nächste angeklickte Formel. Passt ein Schritt nicht zur Formel, wird er ausgelassen und der Editor sagt das.
- **Gleichungsnummer:** Bei Wissenschaftlich lässt sich unter der Vorschau **Mit Gleichungsnummer** ankreuzen. Der Editor schlägt die nächste freie Nummer vor, etwa (4.13), und setzt sie rechtsbündig hinter die Formel. Ist **Folgende Nummern weiterzählen** angekreuzt, werden alle späteren Nummern und die Verweise im Text darauf um eins erhöht. **Abstand zum rechten Rand** (in mm) rückt die Nummer vom rechten Rand ein; **Alle Nummern im Dokument so ausrichten** setzt alle vorhandenen Nummern rechtsbündig mit diesem Abstand. **Gleichungsnummern prüfen** im Reiter Alles bearbeiten meldet doppelte, fehlende und ungeordnete Nummern sowie Verweise ohne Gleichung.
- **Bausteine:** Unter Aufbau gibt es mehrzeilige Formeln, Gleichungssysteme, Fallunterscheidungen, Matrizen, Vektoren und Determinanten, dazu Knöpfe für weitere Zeilen und Spalten. Der Reiter **Einheiten** setzt Einheiten aufrecht und mit dem richtigen kleinen Abstand hinter die Zahl.
- **Suchen und Zuletzt:** Über den Bausteinen sucht ein Feld nach Namen, etwa „Wurzel“, „Planck“ oder „Ohm“. Die zuletzt benutzten Bausteine stehen in der Zeile „Zuletzt“.
- **Meine Formeln:** Im Reiter **Meine Formeln** lassen sich häufig gebrauchte Formeln oder Teile davon unter einem Namen speichern und mit einem Klick an der Cursorstelle einsetzen.
- **Entwurf:** Der Editor merkt sich die Formel in der Eingabe laufend. Wird er geschlossen oder stürzt Word ab, bietet er sie beim nächsten Öffnen oben zum **Wiederherstellen** an.
- **Neu anfangen:** **Neue Formel** leert die Eingabe. Danach fügt der große Knopf wieder neu ein, statt zu ersetzen.

Die Formel steckt unsichtbar im Alternativtext des Bildes. Deshalb lässt sie sich auch nach Wochen noch bearbeiten, solange niemand diesen Alternativtext von Hand ändert.

## Alles bearbeiten

- **Standard für dieses Dokument:** Schrift und Schriftgröße, die beim Konvertieren und Formatieren verwendet werden.
- **Alle Formeln konvertieren:** wandelt alte Formeln aus dem Formel-Editor 3.0 ins neue Format um. Mit **Schriftart übernehmen** und **Schriftgröße übernehmen** bekommen alle Formeln den Standard. Alte Formeln werden wie im alten Formel-Editor wissenschaftlich gesetzt, auch im Fließtext. Bei neuen Formeln bleibt die Darstellung (wissenschaftlich oder einzeilig) erhalten. Das Ergebnis öffnet sich als neues Dokument, das man mit „Speichern unter …“ sichert; das Original bleibt unverändert. Eine Übersicht zeigt danach Ausgangsformate, konvertierte Formeln und Probleme. Problemfälle sind gelb markiert und müssen neu eingegeben werden.
- **Formeln im Dokument analysieren:** zählt die Formeln nach Format, Schrift und Schriftgröße. Ein Klick auf eine Zeile sucht genau diese Formeln.
- **Formeln suchen:** markiert passende Formeln im Dokument türkis. Mit **Zurück** und **Weiter** springt man von Formel zu Formel, formatiert sie einzeln oder alle auf einmal im Standard und entfernt am Ende die Markierungen.

## Gut zu wissen

- Die Formeln sind Vektorgrafiken. Sie bleiben beim Vergrößern und im Druck scharf.
- Der Editor bringt alle Schriften und Bibliotheken selbst mit (Ordner `vendor/`) und fragt beim Arbeiten keine fremden Server mehr. Internet braucht er nur, um sich selbst von GitHub Pages zu laden.
- **Problem melden** unten im Editor bereitet eine E-Mail mit Beschreibung, Formel und Version vor. Die Empfänger-Adresse wird beim ersten Mal im Fenster eingetragen und dann auf diesem Mac gemerkt.
- Meine Formeln, der Entwurf und die Adresse für Problem melden liegen im Speicher von Word. Der Befehl zum Leeren des Zwischenspeichers unten löscht auch sie.

## Prüfung vor jeder Version

`tests/pruefen.mjs` setzt alle Beispielformeln mit dem Editor und vergleicht sie mit dem guten Stand in `tests/stand-beispiele.json`. Es meldet Formeln, die sich nicht setzen lassen, die sich beim Öffnen verändern oder die anders aussehen als vorher. Auf GitHub läuft es bei jedem Pull Request.

```
cd tests && npm install && node pruefen.mjs
node pruefen.mjs --neu --stand ~/stand.json Dokument1.docx Dokument2.docx   # guten Stand aus eigenen Dokumenten anlegen
node pruefen.mjs --stand ~/stand.json --bericht bericht.html Dokument1.docx Dokument2.docx   # später dagegen prüfen
```

Formeln aus privaten Manuskripten gehören nicht in dieses öffentliche Repository. Ihr guter Stand liegt außerhalb.

`tests/fliesstext-probe.mjs` baut ein Word-Dokument mit Formeln im Fließtext in allen drei Darstellungen. Darin sieht man, ob sie auf der Grundlinie sitzen und wie stark sie den Zeilenabstand weiten.

## Version und Aktualisierung

Unten im Editor stehen Name und Version, zum Beispiel „Papas Formeleditor · Version 1.7.1“. Das **Zahnrad** daneben öffnet die Einrichtung. Sie vergleicht diese Version mit der aktuellen Version auf GitHub und bietet die Dokumentation zum Herunterladen an.

- Gibt es eine neuere Version, erscheint beim Start oben ein blauer Hinweis. In der Einrichtung fragt der Editor dann „Jetzt aktualisieren?“. Mit **Ja, jetzt aktualisieren** lädt er sich neu und holt dabei die neuen Dateien von GitHub. Eine gerade eingegebene, noch nicht eingefügte Formel geht dabei verloren.
- **Editor neu laden** lädt den Editor auch ohne neue Version frisch von GitHub.
- **Dokumentation herunterladen** öffnet die bebilderte Bedienungsanleitung als Word-Datei im Browser. Ihre Quelle liegt in `docs/anleitung/`.
- Die Einrichtungsdatei aus Schritt 1 muss für ein Update nicht neu geladen werden. Nur wenn sich der Name des Knopfs im Reiter Start ändern soll, Schritt 1 einmal wiederholen.

Zeigt der Editor gar keine Versionsnummer an, läuft noch eine ältere Fassung aus dem Zwischenspeicher von Word. Dann Word mit ⌘Q beenden und neu starten. Hilft das nicht, bei beendetem Word diese Zeile im Terminal ausführen und Word wieder starten:

```
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/* ~/Library/Containers/com.microsoft.Word/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/*
```

Sie löscht nur den Zwischenspeicher der Word-Add-ins, nicht die Einrichtungsdatei und keine Dokumente.

## Wieder entfernen

Die Datei `formeleditor.xml` aus dem Ordner `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` löschen und Word neu starten.
