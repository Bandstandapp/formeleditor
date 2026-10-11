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

Der Editor hat oben zwei Reiter. **Formel bearbeiten** ist beim Öffnen gewählt und enthält die Kacheln Bausteine, Eingabe und Vorschau. **Alles bearbeiten** enthält die Werkzeuge für das ganze Dokument. Hinter jeder Überschrift erklärt ein kleines **i** die Kachel, sobald man mit der Maus darauf zeigt.

- **Neue Formel:** Den Cursor im Text an die gewünschte Stelle setzen. Die Formel aus den Bausteinen zusammensetzen oder tippen und auf **In Word einfügen** klicken. Den Knopf gibt es unter der Eingabe und unter der Vorschau.
- **Fett, kursiv und Farbe:** Zeichen in der Eingabe markieren und auf **F** (fett), **K** (kursiv bzw. aufrecht) oder **A** (Farbe) klicken. Ohne Markierung gilt die Wahl für die nächsten Zeichen.
- **Zahlen mit Abständen:** Die Leertaste setzt einen schmalen Abstand, etwa in 1.159 652 180 46.
- **Aussehen:** In der Vorschau zwischen **Wissenschaftlich** (Brüche übereinander) und **Einzeilig** (Brüche mit Schrägstrich, für den Fließtext) umschalten und Schrift und Schriftgröße wählen. Der Zoom vergrößert nur die Vorschau.
- **Formel ändern:** Die Formel im Dokument einmal anklicken. Sie erscheint in der Eingabe. Solange nichts geändert ist, heißt der große Knopf **Ohne Änderung übernehmen**; er wandelt eine alte Formel unverändert ins neue Format um. Nach einer Änderung heißt er **Formel in Word ersetzen**. Ein grüner Rahmen um die Eingabe zeigt: Die Formel ist so wie im Dokument. Nach einer Änderung wird der Rahmen orange, die Änderung ist also noch nicht in Word. Nach dem Klick auf den Knopf ist er wieder grün.
- **Verlauf:** Der kleine Pfeil neben **Rückgängig** zeigt die letzten 20 Schritte. Ausgewählte Schritte werden mit **Rückgängig machen** einzeln zurückgenommen, auch wenn danach noch weitergearbeitet wurde. Der Pfeil neben **Wiederholen** zeigt dieselben Schritte. Mit **Anwenden** kommen ausgewählte Schritte zurück, oder sie werden auf eine andere Formel übertragen, zum Beispiel „v fett“ auf die nächste angeklickte Formel. Passt ein Schritt nicht zur Formel, wird er ausgelassen und der Editor sagt das.
- **Neu anfangen:** **Neue Formel** leert die Eingabe. Danach fügt der große Knopf wieder neu ein, statt zu ersetzen.

Die Formel steckt unsichtbar im Alternativtext des Bildes. Deshalb lässt sie sich auch nach Wochen noch bearbeiten, solange niemand diesen Alternativtext von Hand ändert.

## Alles bearbeiten

- **Standard für dieses Dokument:** Schrift und Schriftgröße, die beim Konvertieren und Formatieren verwendet werden.
- **Alle Formeln konvertieren:** wandelt alte Formeln aus dem Formel-Editor 3.0 ins neue Format um. Mit **Schriftart übernehmen** und **Schriftgröße übernehmen** bekommen alle Formeln den Standard. Alte Formeln werden wie im alten Formel-Editor wissenschaftlich gesetzt, auch im Fließtext. Bei neuen Formeln bleibt die Darstellung (wissenschaftlich oder einzeilig) erhalten. Das Ergebnis öffnet sich als neues Dokument, das man mit „Speichern unter …“ sichert; das Original bleibt unverändert. Eine Übersicht zeigt danach Ausgangsformate, konvertierte Formeln und Probleme. Problemfälle sind gelb markiert und müssen neu eingegeben werden.
- **Formeln im Dokument analysieren:** zählt die Formeln nach Format, Schrift und Schriftgröße. Ein Klick auf eine Zeile sucht genau diese Formeln.
- **Formeln suchen:** markiert passende Formeln im Dokument türkis. Mit **Zurück** und **Weiter** springt man von Formel zu Formel, formatiert sie einzeln oder alle auf einmal im Standard und entfernt am Ende die Markierungen.

## Gut zu wissen

- Die Formeln sind Vektorgrafiken. Sie bleiben beim Vergrößern und im Druck scharf.
- Für das Einfügen braucht der Editor Internet, weil er die Schriften für den Formelsatz aus dem Netz lädt. Bereits eingefügte Formeln funktionieren im Dokument auch ohne Internet.

## Version und Aktualisierung

Unten im Editor stehen Name und Version, zum Beispiel „Papas Formeleditor · Version 1.4.0“. Das **Zahnrad** daneben öffnet die Einrichtung. Sie vergleicht diese Version mit der aktuellen Version auf GitHub und bietet die Dokumentation zum Herunterladen an.

- Gibt es eine neuere Version, erscheint beim Start oben ein blauer Hinweis. In der Einrichtung fragt der Editor dann „Jetzt aktualisieren?“. Mit **Ja, jetzt aktualisieren** lädt er sich neu und holt dabei die neuen Dateien von GitHub. Eine gerade eingegebene, noch nicht eingefügte Formel geht dabei verloren.
- **Editor neu laden** lädt den Editor auch ohne neue Version frisch von GitHub.
- **Dokumentation herunterladen** öffnet diese Anleitung als Word-Datei im Browser.
- Die Einrichtungsdatei aus Schritt 1 muss für ein Update nicht neu geladen werden. Nur wenn sich der Name des Knopfs im Reiter Start ändern soll, Schritt 1 einmal wiederholen.

Zeigt der Editor gar keine Versionsnummer an, läuft noch eine ältere Fassung aus dem Zwischenspeicher von Word. Dann Word mit ⌘Q beenden und neu starten. Hilft das nicht, bei beendetem Word diese Zeile im Terminal ausführen und Word wieder starten:

```
rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/* ~/Library/Containers/com.microsoft.Word/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/*
```

Sie löscht nur den Zwischenspeicher der Word-Add-ins, nicht die Einrichtungsdatei und keine Dokumente.

## Wieder entfernen

Die Datei `formeleditor.xml` aus dem Ordner `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` löschen und Word neu starten.
