# Sende individuelle Anhänge an deine Empfänger mit Serienmails

Mit diesen VBA-Code kannst du mit Outlook, Excel und Word individuelle Dateianhänge an deine Empfänger per Serienmail senden.
Kurzüberblick: Diese README erklärt, wie mit dem Word‑Makro "Serienmails mit individuellem Anhang" E‑Mails über Outlook erstellt/versendet werden, Datensätze aus Excel gelesen, Platzhalter im Word‑Text ersetzt, individuelle Anhänge pro Zeile hinzugefügt und optional ein Sendezeitpunkt gesetzt wird. Sie führt durch alle Dialog‑Abfragen und markiert Pflicht‑/Optionalfelder.
Neben dieser Anleitung gibt es noch meinen bebilderten Blogbeitrag: [hier geht's zu meinen Blog mit der Anleitung](https://blogs.urz.uni-halle.de/simpletricks/2023/03/serien-e-mails-mit-individuellen-anhaengen/)

> [!IMPORTANT]
> Diese Anleitung ist 1:1 an den vorliegenden VBA‑Code angepasst: Erforderlich sind die Spalten "E‑Mail" und "Betreff". Alle anderen Spalten sind optional.

Supporte meinen frei verfügbaren Content :)

<a href="https://www.buymeacoffee.com/justanotherjurastudent" target="_blank">
    <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" >
</a>

***


## Voraussetzungen

- Windows mit Desktop‑Outlook, ‑Word und ‑Excel (Office 2016+ oder Microsoft 365).
- Outlook ist bereits gestartet und mit einem sendefähigen Konto verbunden.
- Makros sind erlaubt (Vertrauensstellungscenter) und die Office‑Bibliotheken sind referenziert.

> [!NOTE]
> In Word im VBA‑Editor unter "Extras → Verweise…" die folgenden Verweise aktivieren:
> - Microsoft Outlook xx.0 Object Library
> - Microsoft Word xx.0 Object Library
> - Microsoft Excel xx.0 Object Library
> - Microsoft Office xx.0 Object Library

***

## Installation

1. Word öffnen und die spätere E‑Mail‑Vorlage als neue Datei anlegen.
2. VBA‑Editor mit Alt+F11 öffnen → "Datei → Datei importieren…" → die .bas‑Datei "Serienmails mit individuellem Anhang.bas" importieren.
3. Unter "Extras → Verweise…" die Bibliotheken (siehe oben) aktivieren.
4. Word‑Vorlage als ".docm" speichern, z. B. "Vorlage_Serienmail.docm".

> [!WARNING]
> Makros nur aus vertrauenswürdigen Quellen ausführen. Falls Makros blockiert werden, Datei an einen vertrauenswürdigen Speicherort legen oder die Signatur vertrauen.

***

## Word‑Vorlage anlegen

- Den kompletten E‑Mail‑Text in Word gestalten; Formatierungen, Listen, Links und Bilder werden als HTML in die E‑Mail übernommen.
- Folgende Platzhalter werden unterstützt und automatisch ersetzt:
  - `%Anrede%`, `%Titel%`, `%Vorname%`, `%Nachname%`, `%Unternehmen%`
- Beispiel Kopf (optional):  
  "%Anrede% %Titel% %Vorname% %Nachname%,"  
  Danach der eigentliche Nachrichtentext.
- **Automatische Textbereinigung:** Der Code entfernt automatisch doppelte Leerzeichen und Leerzeichen vor Satzzeichen (z. B. `, . ! ? : ;`), die durch leere Platzhalter entstehen könnten.

> [!TIP]
> Für Serien mit vielen Einträgen zuerst einen Testlauf mit wenigen Zeilen im Modus "nur generieren" durchführen; anschließend Inhalte prüfen.

***

## Excel‑Tabelle erstellen

- Jede Zeile entspricht genau einer E‑Mail.
- Der Code erkennt Spaltenköpfe automatisch und bietet eine interaktive Korrektur an.
- Unterstützte/gesuchte Spalten (Kopfzeilen) und Status:

| Spalte (Kopf)           | Zweck                                 | Pflicht?                          | Beispiel/Inhaltshinweise |
|---|---|---|---|
| E‑Mail                  | Empfängeradresse (.To)                | Ja                                | `alice@example.org` |
| Betreff                 | Betreffzeile                          | Ja                                | "Ihre Unterlagen 2025" |
| Anhang / Anhänge        | Dateipfade pro Zeile                  | Optional                          | `C:\A\1.pdf; C:\A\2.pdf` |
| CC                      | Kopie‑Empfänger                       | Optional                          | `team@example.org; buchhaltung@example.org` |
| BCC                     | Blindkopie                            | Optional                          | `leitung@example.org` |
| Anrede                  | Anrede‑Quelle                         | Optional                          | "Herr"/"m"/"Frau"/"w" (für formelle Logik) |
| Titel                   | Titel vor dem Namen                   | Optional                          | "Dr." |
| Vorname                 | Vorname                               | Optional                          | "Max" |
| Nachname                | Nachname                              | Optional                          | "Mustermann" |
| Unternehmen/Unternehmensname | Firmenname                      | Optional                          | "Beispiel GmbH" |
| Sendezeitpunkt          | Geplanter Versand                     | Optional                          | Gültiges Datum/Uhrzeit, siehe unten |

> [!NOTE]
> - **Anrede & Titel:** Diese Felder werden automatisch getrimmt (führende/folgende Leerzeichen entfernt). Ist ein Titel in Excel leer, wird der Platzhalter `%Titel%` im Word-Dokument restlos entfernt, ohne dass ein störendes Leerzeichen zurückbleibt.
> - **Datenbereich:** Der Code erkennt das Ende der Tabelle automatisch, auch wenn in der ersten Spalte (A) einzelne Zellen leer sind.

> [!IMPORTANT]
> - Wenn Sie die formelle Anrede (Sehr geehrte ...) benutzen, erkennt der Code automatisch verschiedene Angaben in der Spalte "Anrede":
>   - **Männlich:** "Herr", "m", "Mann", "männlich" → *Sehr geehrter Herr*
>   - **Weiblich:** "Frau", "w", "f", "weiblich" → *Sehr geehrte Frau*
> - Mehrere Anhänge werden in EINER Zelle durch ein **Komma** oder **Semikolon** getrennt, z. B.:  
>   `C:\Rechnungen\RE-4711.pdf; C:\Rechnungen\AGB.pdf`  
> - Als Dateipfadseparator werden sowohl der Windows-Standard `\` als auch `/` unterstützt. Das **Komma** oder **Semikolon** trennt nur mehrere Pfade innerhalb derselben Zelle.  
> - Der Code verarbeitet zuverlässig in Anführungszeichen gesetzte Pfade, **Kommas im Dateinamen**, `file:///`‑URLs, UNC‑Pfade (`\\Server\Freigabe\...`) und **relative Pfade** relativ zum Speicherort der Excel‑Datei. Hyperlinks in Zellen werden berücksichtigt.
> - Dateipfade müssen existieren und lesbar sein. Fehler werden gesammelt angezeigt und der Vorgang bricht zur Korrektur ab.

***

## Sendezeitpunkt korrekt formatieren

- Spalte "Sendezeitpunkt" ist optional. Ist der Zellenwert ein von Excel erkennbares Datum/Uhrzeit und liegt in der Zukunft, wird die E‑Mail mit Verzögerung (.DeferredDeliveryTime) geplant.
- Empfohlene Formate:
  - `YYYY-MM-DD HH:MM` (z. B. `2025-09-15 09:00`)
  - `DD.MM.YYYY HH:MM` (z. B. `15.09.2025 14:30`)
- Die Zellen sollten in Excel ausdrücklich als Datum/Uhrzeit formatiert sein. Leere Zellen bedeuten "keine Verzögerung".

> [!WARNING]
> Liegt der Wert in der Vergangenheit oder ist er ungültig, wird die E‑Mail nicht verzögert. Es erscheint ein Hinweisdialog.

***

## Schritt‑für‑Schritt: Ablauf & Dialoge

1) Outlook‑Check  
- Beim Start prüft das Makro, ob Outlook läuft. Wenn nicht, erscheint ein Hinweis und der Ablauf wird beendet.

2) Excel‑Datei wählen  
- Ein Dateidialog öffnet sich ("Excel‑Liste auswählen"). Die gewählte Arbeitsmappe wird im Hintergrund geöffnet.

3) Arbeitsblatt auswählen  
- Gibt es mehrere Blätter, wird eine nummerierte Liste angezeigt und nach der Blattnummer gefragt. Leere Eingabe beendet den Vorgang.

4) Spalten finden & bestätigen  
- Der Code sucht automatisch: "Anrede", "Titel", "Vorname", "Nachname", "Unternehmen/Unternehmensname", "E‑Mail", "Betreff", "Anhang/Anhänge", "CC", "BCC" sowie "Sendezeitpunkt".  
- Danach zeigt er eine Übersicht der erkannten Spaltenbuchstaben und fragt: "Sind Sie einverstanden?"  
  - "Ja": weiter  
  - "Nein": gezielte Korrektur einzelner Spalten (Buchstaben eingeben)  
  - "Abbrechen": beendet  
- Wichtig: "E‑Mail" und "Betreff" müssen gesetzt sein; fehlen sie, wird abgebrochen.

5) Startzeile festlegen  
- Abfrage: "Beginnen die Daten ab Zeile 2?"  
  - "Ja": Start bei Zeile 2  
  - "Nein": gewünschte Startzeile eingeben (Kopfzeile bleibt außerhalb)
- Mindestprüfung: Wenn E-Mail-Adresse oder Betreff in den Datensätzen fehlen, erscheint eine Warnung mit Optionen:
  - "Abbrechen": Vorgang beenden
  - "Überspringen": Datensätze ohne E-Mail/Betreff werden übersprungen

6) Anrede‑Variante wählen  
- Abfrage: "Formelle Anrede übernehmen?"  
  - "Ja": Wenn in Excel "Frau" → "Sehr geehrte Frau", "Herr" → "Sehr geehrter Herr"  
  - "Nein": Anrede wird aus Excel unverändert übernommen (auch freier Text möglich)

7) Sendezeitpunkt bestätigen (falls nötig)
- Falls die Spalte "Sendezeitpunkt" im Schritt 4 nicht automatisch gefunden oder manuell korrigiert wurde, erscheint hier eine zusätzliche Abfrage.

8) Outlook‑Kontenauswahl (falls mehrere Konten)  
- Sind mehrere sendende Konten in Outlook konfiguriert, wird eine Auswahl angezeigt: "Mit welchem Konto sollen die E‑Mails versendet werden?"  
  - Wählen Sie Ihr Konto aus der Liste
  - "Abbrechen": Versand wird beendet
  - Nach Auswahl des Kontos wird nochmal zur Bestätigung aufgefordert. Bei "Nein" kann erneut ausgewählt werden. Bei "Abbrechen" wird der Vorgang beendet.
- Hat man nur ein Konto oder keine Mehrfachkonten, entfällt dieser Schritt.
- Der Debug-Bereich im VBA-Editor zeigt, welche Konten erkannt und welches letztlich verwendet wurde.

8) Versandmodus wählen  
- Abfrage: "E‑Mails direkt versenden?"  
  - "Ja": zusätzliche Sicherheitsbestätigung; E‑Mails werden automatisch gesendet  
  - "Nein": E‑Mails werden nur generiert und im Editor angezeigt (Entwürfe prüfen/senden)

9) Anhang‑Validierung  
- Der Code prüft für jede Zeile die Existenz der angegebenen Dateien. Fehlende Dateien werden zeilenweise gelistet.  
- Bei Fehlern: Dialog mit Zusammenfassung; Ablauf wird beendet, damit die Pfade korrigiert werden können.

10) E‑Mail‑Erstellung  
- Für jede Zeile wird ein temporäres Word‑Dokument erstellt, die Platzhalter (%Anrede% etc.) ersetzt und der Inhalt als HTML in eine neue Outlook‑Mail kopiert.  
- Dann werden Anhänge aus der Zelle hinzugefügt. `file:///`‑URLs werden in Pfade umgewandelt; Anführungszeichen an den Enden werden entfernt; relative Pfade werden relativ zum Workbook‑Ordner aufgelöst.  
- Ist ein gültiger zukünftiger "Sendezeitpunkt" gesetzt, wird die verzögerte Zustellung aktiviert; bei ungültigen Werten erscheint eine Warnung.

11) Versand/Anzeige & Abschluss  
- Je nach Modus werden Mails gesendet oder nur generiert (das Versenden liegt dann in Ihrer Hand).  
- Am Ende erscheint eine Zusammenfassung (erfolgreich verarbeitet/Fehler) und die Objekte werden aufgeräumt.

> [!TIP]
> Für die erste Serie immer "Nur generieren" wählen, Entwürfe prüfen (Empfänger, Betreff, Text, Anhänge, verzögerte Zustellung), dann erst den Direktversand verwenden.

***

## Beispiele

### Beispiel‑Word (Kopf und Platzhalter)
```
%Anrede% %Titel% %Vorname% %Nachname%,

anbei erhalten Sie die gewünschten Unterlagen für %Unternehmen%.

Freundliche Grüße
```

### Beispiel‑Excel (kleine Tabelle)

| E‑Mail             | Betreff                     | Anhang                                                     | CC                      | BCC               | Sendezeitpunkt  | Anrede | Titel | Vorname | Nachname | Unternehmen      |
|---|---|---|---|---|---|---|---|---|---|---|
| alice@beispiel.de  | Ihre Rechnung RE‑4711       | C:\Rechnungen\RE‑4711.pdf; C:\Rechnungen\AGB.pdf          | buchhaltung@beispiel.de |                   | 2025-09-15 09:00 | Frau   | Dr.   | Alice   | Beispiel  | Beispiel GmbH    |
| bob@beispiel.de    | Einladung zum Webinar       | C:/Einladungen/Bob.pdf                                    |                         |                   |                 | Herr   |       | Bob     | Muster    | Muster AG        |
| clara@beispiel.de  | Dokumente zur Vertragsänderung | \\server\share\Clara\Aenderung.pdf                       | team@beispiel.de        | chef@beispiel.de  | 15.09.2025 14:30 |        |       | Clara   | Meyer    | ACME SE          |

> [!NOTE]
> - Leer gelassene "Sendezeitpunkt"‑Zellen bedeuten Sofortversand (bzw. keine verzögerte Zustellung).  
> - In "Anhang" können Pfade in Anführungszeichen stehen. Kommas oder Semikolons im Dateinamen sind erlaubt; die Aufteilung trennt zuverlässig zwischen Trennzeichen und Zeichen im Namen.

***

## Häufige Fehler & Lösungen

- "Outlook ist nicht geöffnet"  
  → Outlook vor Start des Makros öffnen.

- "Benutzerdefinierter Typ nicht definiert"  
  → In Word unter "Extras → Verweise…" die Office‑Bibliotheken aktivieren.

- "Datei nicht gefunden" in der Anhang‑Prüfung  
  → Pfade korrigieren, Berechtigungen prüfen, Netzlaufwerke eingebunden, `file:///`‑URLs korrekt, relative Pfade relativ zum Excel‑Dateiordner verstehen. Sowohl `\` als auch `/` sind als Pfadtrenner zulässig.

- Ungültiger "Sendezeitpunkt"  
  → Zellenformat auf Datum/Uhrzeit setzen; lokal gültige Eingaben verwenden; nur zukünftige Zeitpunkte verzögern den Versand.

- CC/BCC werden nicht gezogen  
  → Mehrere Adressen per Semikolon `;` trennen; Spaltenkopf korrekt benennen und bei der Spaltenbestätigung prüfen.

***

## Best Practices

- Zuerst Testlauf mit 3–5 Zeilen im Modus "nur generieren".
- Eindeutige Dateinamen.
- Spaltenköpfe so benennen, dass die Auto‑Erkennung greift; ansonsten bei der Spaltenbestätigung sauber nachtragen.
- Große Serien in Batches versenden und ggf. zwischen den Läufen kurze Pausen einlegen (organisatorische Limits/Spam‑Regeln beachten).

***

## Was ist Pflicht, was optional? (Kurzfassung)

- Pflichtspalten:  
  - `E‑Mail`  
  - `Betreff`

- Optionale Spalten:  
  - `Anhang`, `CC`, `BCC`, `Anrede`, `Titel`, `Vorname`, `Nachname`, `Unternehmen/Unternehmensname`, `Sendezeitpunkt`

> [!IMPORTANT]
> - Mehrere Anhänge in EINER Zelle per Komma oder Semikolon trennen.  
> - Sowohl `\` als auch `/` werden als Dateipfadseparator unterstützt. Das Komma/Semikolon dient ausschließlich als Trennzeichen zwischen mehreren Pfaden in einer Zelle.

***

## Changelog

### Update am 25.03.2025
Der Code wurde stark angepasst, ist nun entschlackter und robuster. Zusätzlich wurden viele Kommentare an den Code geschrieben, um ihn verstehen zu können. Eine zusätzliche Sicherheitsabfrage vor dem Direktversand wurde implementiert, um versehentliches Senden zu verhindern.

### Update am 28.03.2025
Verbesserung der Formatübernahme: Die Schriftgröße aus der Word-Vorlage wird nun zuverlässiger in die E-Mail übernommen.

### Update am 08.05.2025
Zusätzlich kann nun noch ein Unternehmensname verwendet werden. In der Excel-Tabelle muss diese Spalte hierfür den Namen "Unternehmen" oder "Unternehmensnamen" haben, um automatisch erkannt zu werden. In dem Word-Dokument wird der Unternehmensname über den Platzhalter %Unternehmen% eingefügt.

### Update am 09.05.2025
Wichtiges Update: Die formelle Anrede wird nun korrekt gesetzt, wenn sich in der Excel-Zelle "Herr" oder "Frau" befindet. Auch müssen nicht alle möglichen Spalten (z. B. Vorname oder BCC) existieren, um die E-Mail zu generieren.

### Update am 11.05.2025
Nun werden auch (un)geordnete Listen (ggf. mit mehreren Ebenen) sowie abweichende Schriftfarben und Texthervorhebungsfarben in die E-Mail übernommen. Zudem startet der Dateiauswahldialog für die Excel-Liste nun standardmäßig im Dokumenten-Ordner des Nutzers.

### Update am 12.05.2025
Wichtiges Update: Statt das Word-Dokument mit HTML-Tags zu versehen, wird das Dokument temporär als HTML-Dokument abgespeichert und dessen Inhalt wird in die E-Mail eingefügt. So sollten die allermeisten Formatierungen aus Word erhalten bleiben. 
In einer zusätzlichen Spalte in Excel kann ein Sendezeitpunkt für jede Nachricht angegeben werden, wann die E-Mail versendet werden soll. Der Sendezeitpunkt sollte als TT.MM.JJJJ HH:MM formatiert sein. Die Spalte sollte mit "Sendezeitpunkt" bezeichnet werden, um automatisch erkannt zu werden. Außerdem wird am Anfang der Code-Durchführung ein Check gemacht, ob Outlook im Hintergrund aktiv ist. Wenn nicht, dann sollte Outlook noch gestartet werden.

### Update vom 13.09.2025
Wichtiges Update: Nun können auch Bilder in die E-Mail eingefügt werden. Möglich macht dies die technische Änderung, dass das Word-Dokument nicht mehr temporär als HTML-Datei abgespeichert wird, sondern der Dokumenteninhalt in die E-Mail hinein kopiert wird (mit den Platzhalterersetzungen).
Außerdem können Dateinamen nun auch Kommas enthalten - davor war das Komma das unmissverständliche Trennzeichen zwischen zwei Dateipfaden.
Zuletzt wurden die (Warn)Meldungen verbessert und Debug-Logs in dem Direktbereich im VBA-Editor hinzugefügt.

### Update vom 23.12.2025
Wichtiges Update zur Textqualität: Der Code bereinigt nun automatisch doppelte Leerzeichen und entfernt Leerzeichen vor Satzzeichen, die oft durch optionale, aber leere Platzhalter (wie `%Titel%`) entstehen. Zudem werden Anreden und Titel nun konsequent getrimmt. Die Erkennung der letzten Zeile in Excel wurde verbessert, sodass leere Zellen in der ersten Spalte nicht mehr zum vorzeitigen Abbruch führen.
Zudem werden nun sowohl `/` als auch `\` als Dateipfadseparatoren unterstützt und Anhänge können flexibel durch Komma oder Semikolon getrennt werden. Die formelle Anrede erkennt nun zudem flexibel verschiedene Geschlechtsangaben (m/w/f/Mann/weiblich etc.). Die Anhang-Spalte ist nun zudem vollkommen optional und muss nicht mehr zwingend in der Excel-Tabelle existieren.


### Update vom 29.01.2026
**Großes Funktions- und Qualitätsupdate:**
- **Outlook-Kontenauswahl:** Bei mehreren Outlook-Konten wird der User nun vor dem Generieren oder Versenden darauf hingewiesen, mit welchem Konto versendet wird und kann dieses gezielt auswählen.
- **Mindestdaten-Check:** Wenn die E-Mail-Adresse oder der Betreff in einem Datensatz fehlt, wirst du darauf hingewiesen. Du kannst dann abbrechen oder diesen Datensatz überspringen beim Versand.
- **Fehlerbehandlung & Logging:** Die Fehlerbehandlung wurde deutlich erweitert, u. a. mit einer LogAbort-Funktion und ausführlicher Protokollierung (Debug.Print) zur besseren Nachvollziehbarkeit.
- **Find- und Variablen-Optimierung:** Die Erkennung der letzten Zeile in Excel ist robuster (Find-Objektprüfung). Alle Variablen werden nun explizit deklariert, was die Wartbarkeit und Fehlersicherheit erhöht.
- **Direktversand stabiler:** Der Modus "DirectlySend" funktioniert nun zuverlässig, indem vor dem Versand der Outlook‑Inspector gezielt initialisiert wird. Damit sollte der Laufzeitfehler 5 verschwinden.

