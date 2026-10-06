# Formeleditor in Word einrichten (Mac)

Diese Anleitung richtet den Formeleditor als Add-in in Word auf dem Mac ein. Sie ist so geschrieben, dass Jens sie über TeamViewer auf Papas Rechner durchgehen kann. Die Einrichtung dauert ein paar Minuten und muss nur einmal gemacht werden. Danach steht der Editor in Word im Reiter „Start“ als Knopf „Formeleditor“ bereit.

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
3. Im Fenster, das sich öffnet, erscheint unter **Entwickler-Add-ins** (englisch „Developer Add-ins“) der Eintrag **Formeleditor**. Ihn einmal anklicken.
4. Rechts öffnet sich der Formeleditor als Seitenleiste. Im Reiter **Start** gibt es von nun an außerdem den Knopf **Formeleditor**.

Falls der Eintrag nicht erscheint, Word mit ⌘Q ganz beenden und neu starten. Hilft das nicht, ist die Datei aus Schritt 1 nicht im richtigen Ordner gelandet.

## So wird der Editor benutzt

- **Neue Formel:** Den Cursor im Text an die gewünschte Stelle setzen. Im Editor die Formel aus den Bausteinen zusammensetzen oder tippen, Darstellung, Schrift und Schriftgröße wählen und auf **In Word einfügen** klicken.
- **Formel ändern:** Die Formel im Dokument einmal anklicken. Sie erscheint automatisch im Editor, und der große Knopf heißt jetzt **Formel in Word ersetzen**. Nach dem Ändern darauf klicken, dann wird die alte Formel an derselben Stelle ausgetauscht. Falls sie nicht von selbst erscheint, hilft der Knopf **Markierte Formel bearbeiten**.
- **Neu anfangen:** **Neue Formel** leert das Eingabefeld. Danach fügt der große Knopf wieder neu ein, statt etwas zu ersetzen.

Die Formel steckt unsichtbar im Alternativtext des Bildes. Deshalb lässt sie sich auch nach Wochen noch bearbeiten, solange niemand diesen Alternativtext von Hand ändert.

## Gut zu wissen

- Die Formeln sind Vektorgrafiken. Sie bleiben beim Vergrößern und im Druck scharf.
- Für das Einfügen braucht der Editor Internet, weil er die Schriften für den Formelsatz aus dem Netz lädt. Bereits eingefügte Formeln funktionieren im Dokument auch ohne Internet.
- Alte Formeln aus dem Formel-Editor 3.0 (in Dokumenten, die von .doc nach .docx umgewandelt wurden) lassen sich genauso anklicken. Der Editor übernimmt sie, und „Formel in Word ersetzen“ speichert sie im neuen Format.

## Wieder entfernen

Die Datei `formeleditor.xml` aus dem Ordner `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` löschen und Word neu starten.
