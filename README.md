# Sende individuelle Anhänge an deine Empfänger mit Serienmails

Mit diesen VBA-Codes kannst du mit Outlook, Excel und Word individuelle Dateianhänge an deine Empfänger per Serienmail senden.
Für eine ausführliche Beschreibung, wie damit umzugehen ist, schaue dafür auf meinen Blog: [hier geht's zu meinen Blog mit der Anleitung](https://blogs.urz.uni-halle.de/simpletricks/2023/03/serien-e-mails-mit-individuellen-anhaengen/)

### Update am 25.03.2025
Der Code wurde stark angepasst, ist nun entschlackter und robuster. Zusätzlich wurden viele Kommentare an den Code geschrieben, um ihn verstehen zu können.

### Update am 08.05.2025
Zusätzlich kann nun noch ein Unternehmensname verwendet werden. In der Excel-Tabelle muss diese Spalte hierfür den NAmen "Unternehmen" oder "Unternehmensnamen" haben, um automatisch erkannt zu werden. In dem Word-Dokument wird der Unternehmensname über den Platzhalter %Unternehmen% eingefügt.

### Update am 09.05.2025
Wichtiges Update: Die formelle Anrede wird nun korrekt gesetzt, wenn sich in der Excel-Zelle Herr oder Frau befindet. Auch müssen nicht alle möglichen Spalten (zB Vorname oder BCC) existieren, um die E-Mail zu generieren.

### Update am 11.05.2025
Nun werden auch (un)geordnete Listen (ggf. mit mehreren Ebenen) sowie abweichende Schriftfarben und Texthervorhebungsfarben in die E-Mail übernommen.

### Update am 12.05.2025
Wichtiges Update: Statt das Word-Dokument mit HTML-Tags zu versehen, wird das Dokument temporär als HTML-Dokument abgespeichert und dessen Inhalt wird in die E-Mail eingefügt. So sollten die allermeisten Formatierungen aus Word erhalten bleiben. 
In einer zusätzlichen Spalte in Excel kann ein Sendezeitpunkt für jede Nachricht angegeben werden, wann die E-Mail versendet werden soll. Die Spalte sollte mit "Sendezeitpunkt" bezeichnet werden, um automatisch erkannt zu werden. Außerdem wird am Anfang der Code-Durchführung ein Check gemacht, ob Outlook im Hintergrund aktiv ist. Wenn nicht, dann sollte Outlook noch gestartet werden.

Supporte meinen frei verfügbaren Content :)

<a href="https://www.buymeacoffee.com/justanotherjurastudent" target="_blank">
    <img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" >
</a>

