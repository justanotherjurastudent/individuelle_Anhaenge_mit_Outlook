'******************************************************************************
' ** MODIFIKATIONS-HINWEISE (ZUR EINFÜGUNG AM ANFANG DES CODES) **
'******************************************************************************
' 1. **Sprachanpassung**: Alle Texte in deutschen Dialogen können an Ihre Sprache angepasst werden (z.B. "E-Mail auswählen" in englisch).
' 2. **Spaltenanpassung**: Die Suchbegriffe für Spaltenköpfe (z.B. "Anrede", "E-Mail") können erweitert oder geändert werden.
' 3. **Formatierung**: Die HTML-Codezeichen für Formate (z.B. <b> für Fettdruck) können durch andere Tags ersetzt werden.
' 4. **Anhang-Validierung**: Die Anhang-Prüfung kann erweitert werden, um Ordnerpfade oder Netzwerkpfade zu unterstützen.
' 5. **Vorlagenverwaltung**: Mehrere E-Mail-Vorlagen könnten über ein Auswahlmenü eingeführt werden.
' 6. **Automatisierung**: Die E-Mail-Versendung könnte über einen Timer oder Terminplaner gesteuert werden.
'******************************************************************************

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub SendEmailsFromWordWithExcelWithAbfrage()
    '******************************************************************************
    ' ** 1. Variablen für die Verbindung zu Outlook, Word und Excel **
    '******************************************************************************
    Dim objOutlook As Object          ' Outlook-Objekt für E-Mail-Funktionen
    Dim objMail As Object             ' Einzelne E-Mail-Nachricht
    Dim doc As Document               ' Aktuelles Word-Dokument
    Dim xlApp As Excel.Application    ' Excel-Anwendung
    Dim xlWB As Excel.Workbook        ' Excel-Datei
    Dim xlWS As Excel.Worksheet       ' Excel-Arbeitsblatt
    Dim Pfad As Variant               ' Pfad zur ausgewählten Excel-Datei
    Dim fd As Office.FileDialog       ' Dialog für Dateiauswahl
    Dim objUndo As UndoRecord         ' Rückgängigmachung für Word-Änderungen
    
    Set objUndo = Application.UndoRecord
    
    Dim fso As Object                 ' Dateisystem-Objekt für Dateiprüfung
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Prüfen ob Outlook läuft
    Dim outlookRunning As Boolean
    outlookRunning = False
    
    ' Prozesse durchsuchen nach OUTLOOK.EXE
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'OUTLOOK.EXE'")
    
    If colProcesses.Count > 0 Then
        outlookRunning = True
    End If
    
    If Not outlookRunning Then
        Dim outlookResponse As VbMsgBoxResult
        outlookResponse = MsgBox("Outlook scheint nicht geöffnet zu sein. Normalerweise muss Outlook geöffnet sein, um E-Mails zu versenden." & vbCrLf & _
                                "Möchten Sie trotzdem fortfahren?", vbQuestion + vbYesNo, "Outlook-Prüfung")
        
        If outlookResponse = vbNo Then
            MsgBox "Bitte starten Sie Outlook und versuchen Sie es erneut.", vbInformation
            GoTo Cleanup
        End If
    End If
    
    '******************************************************************************
    ' ** 2. Excel-Datei auswählen (MODIFIKATIONSMÖGLICHKEIT: Weitere Dateitypen hinzufügen) **
    '******************************************************************************
    Set xlApp = New Excel.Application
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Excel-Liste auswählen"
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xl*" ' <--- Hier können Sie weitere Dateitypen hinzufügen (z.B. "*.csv")
        .AllowMultiSelect = False
        .ButtonName = "Auswählen"
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\"
        If .Show = -1 Then
            Pfad = .SelectedItems(1)
            Set xlWB = xlApp.Workbooks.Open(Pfad)
            
            '******************************************************************************
            ' ** 3. Arbeitsblatt auswählen (MODIFIKATIONSMÖGLICHKEIT: Name-Eingabe statt Nummer) **
            '******************************************************************************
            Dim selectedSheet As Integer
            Dim validInput As Boolean
            
            If xlWB.Worksheets.Count = 1 Then
                selectedSheet = 1
                validInput = True
            Else
                validInput = False
            End If

            Dim sheetList As String
            Dim i As Integer
            Do While Not validInput
                sheetList = "" ' Zurücksetzen, um Duplikate zu vermeiden
                For i = 1 To xlWB.Worksheets.Count
                    sheetList = sheetList & i & " - " & xlWB.Worksheets(i).Name & vbCrLf
                Next i

                Dim selectedNumberStr As String
                selectedNumberStr = InputBox( _
                    Prompt:=sheetList & vbCrLf & vbCrLf & _
                           "Geben Sie die Nummer des gewünschten Arbeitsblatts ein:" & vbCrLf & _
                           "(Zum Abbrechen bitte das Eingabefeld leer lassen und OK klicken)", _
                    Title:="Arbeitsblatt auswählen", _
                    Default:="")

                ' Abbruch, wenn keine Eingabe erfolgt
                If selectedNumberStr = "" Then
                    GoTo Cleanup
                End If

                If IsNumeric(selectedNumberStr) Then
                    selectedSheet = CInt(selectedNumberStr)
                    If selectedSheet >= 1 And selectedSheet <= xlWB.Worksheets.Count Then
                        validInput = True
                    Else
                        MsgBox "Ungültige Eingabe! Bitte Zahl zwischen 1 und " & xlWB.Worksheets.Count & " eingeben.", vbExclamation
                    End If
                Else
                    MsgBox "Ungültige Eingabe! Bitte eine Zahl eingeben.", vbExclamation
                End If
            Loop
            
            '******************************************************************************
            ' ** 4. Spaltenfindung (MODIFIKATIONSMÖGLICHKEIT: Suchbegriffe erweitern) **
            '******************************************************************************
            Set xlWS = xlWB.Worksheets(selectedSheet)
            
            ' Restliche Initialisierung
            Set doc = ActiveDocument
            Set objOutlook = CreateObject("Outlook.Application")
            
            ' Spaltenfindung
            Dim SpalteAnrede As String, SpalteTitel As String
            Dim SpalteVorname As String, SpalteNachname As String
            Dim SpalteTo As String, SpalteSubj As String, SpalteAttach As String
            Dim SpalteUnternehmen As String
            Dim SpalteSendezeitpunkt As String
            
            ' Anrede-Spalte finden
            Dim AnredeRange As Excel.Range
            Set AnredeRange = xlWS.Range("A1:Z1").Find("Anrede", LookIn:=xlValues, LookAt:=xlWhole)
            If AnredeRange Is Nothing Then
                SpalteAnrede = InputBox("Spalte für ""Anrede"" (z.B. A oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteAnrede = Chr(AnredeRange.Column + 64)
            End If
            
            ' Titel-Spalte finden
            Dim TitelRange As Excel.Range
            Set TitelRange = xlWS.Cells.Find("Titel", LookIn:=xlValues, LookAt:=xlWhole)
            If TitelRange Is Nothing Then
                SpalteTitel = InputBox("Spalte für ""Titel"" (z.B. B oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteTitel = Chr(TitelRange.Column + 64)
            End If
            
            ' Vorname-Spalte finden
            Dim VornameRange As Excel.Range
            Set VornameRange = xlWS.Cells.Find("Vorname", LookIn:=xlValues, LookAt:=xlWhole)
            If VornameRange Is Nothing Then
                SpalteVorname = InputBox("Spalte für ""Vorname"" (z.B. C oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteVorname = Chr(VornameRange.Column + 64)
            End If
            
            ' Nachname-Spalte finden
            Dim NachnameRange As Excel.Range
            Set NachnameRange = xlWS.Cells.Find("Nachname", LookIn:=xlValues, LookAt:=xlWhole)
            If NachnameRange Is Nothing Then
                SpalteNachname = InputBox("Spalte für ""Nachname"" (z.B. D oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteNachname = Chr(NachnameRange.Column + 64)
            End If

            ' Unternehmen-Spalte finden
            Dim SuchbegriffeUnternehmen As Variant
            SuchbegriffeUnternehmen = Array("Unternehmen", "Unternehmensname")
            Dim UnternehmenRange As Excel.Range
            For Each term In SuchbegriffeUnternehmen
                Set UnternehmenRange = xlWS.Cells.Find(term, LookIn:=xlValues, LookAt:=xlWhole)
                If Not UnternehmenRange Is Nothing Then
                    SpalteUnternehmen = Chr(UnternehmenRange.Column + 64)
                    Exit For
                End If
            Next term
            If UnternehmenRange Is Nothing Then
                SpalteUnternehmen = InputBox("Spalte für ""Unternehmen"" (z.B. X oder leer lassen, wenn nicht vorhanden):")
            End If
            
            ' E-Mail-Spalte finden
            Dim Suchbegriffe As Variant
            Suchbegriffe = Array("E-Mail", "email", "e-Mail", "e-mail")
            Dim ToRange As Excel.Range
            For Each term In Suchbegriffe
                Set ToRange = xlWS.Cells.Find(term, LookIn:=xlValues, LookAt:=xlWhole)
                If Not ToRange Is Nothing Then
                    SpalteTo = Chr(ToRange.Column + 64)
                    Exit For
                End If
            Next term
            If ToRange Is Nothing Then
                SpalteTo = InputBox("Spalte für ""E-Mail"" (z.B. E):")
                If SpalteTo = "" Then
                    MsgBox "Die Spalte mit den E-Mail-Adressen muss definiert sein. Vorgang wurde abgebrochen.", vbExclamation
                    GoTo Cleanup
                End If
            End If

            ' CC-Spalte finden
            Dim CCRange As Excel.Range
            Dim SpalteCC As String
            Set CCRange = xlWS.Cells.Find("CC", LookIn:=xlValues, LookAt:=xlWhole)
            If CCRange Is Nothing Then
                SpalteCC = InputBox("Spalte für ""CC"" (z.B. H oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteCC = Chr(CCRange.Column + 64)
            End If

            ' BCC-Spalte finden
            Dim BCCRange As Excel.Range
            Dim SpalteBCC As String
            Set BCCRange = xlWS.Cells.Find("BCC", LookIn:=xlValues, LookAt:=xlWhole)
            If BCCRange Is Nothing Then
                SpalteBCC = InputBox("Spalte für ""BCC"" (z.B. I oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteBCC = Chr(BCCRange.Column + 64)
            End If
            
            ' Betreff-Spalte finden
            Dim BetreffRange As Excel.Range
            Set BetreffRange = xlWS.Cells.Find("Betreff", LookIn:=xlValues, LookAt:=xlWhole)
            If BetreffRange Is Nothing Then
                SpalteSubj = InputBox("Spalte für ""Betreff"" (z.B. F oder leer lassen, wenn nicht vorhanden):")
                If SpalteSubj = "" Then
                    MsgBox "Die Spalte mit dem E-Mail-Betreff muss definiert sein. Vorgang wurde abgebrochen.", vbExclamation
                    GoTo Cleanup
                End If
            Else
                SpalteSubj = Chr(BetreffRange.Column + 64)
            End If
            
            ' Anhang-Spalte finden
            Dim Suchbegriffe2 As Variant
            Suchbegriffe2 = Array("Anhang", "Anhänge")
            Dim AnhangRange As Excel.Range
            For Each term In Suchbegriffe2
                Set AnhangRange = xlWS.Cells.Find(term, LookIn:=xlValues, LookAt:=xlWhole)
                If Not AnhangRange Is Nothing Then
                    SpalteAttach = Chr(AnhangRange.Column + 64)
                    Exit For
                End If
            Next term
            If AnhangRange Is Nothing Then
                SpalteAttach = InputBox("Spalte für Anhänge (z.B. G oder leer lassen, wenn nicht vorhanden):")
            End If

            Dim SendezeitpunktRange As Excel.Range
            Set SendezeitpunktRange = xlWS.Cells.Find("Sendezeitpunkt", LookIn:=xlValues, LookAt:=xlWhole)
            If SendezeitpunktRange Is Nothing Then
                SpalteSendezeitpunkt = InputBox("Spalte für Sendezeitpunkt (z.B. K oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteSendezeitpunkt = Chr(SendezeitpunktRange.Column + 64)
            End If
            
            ' Benutzerbestätigung der Spalten
            Dim confirmColumns As VbMsgBoxResult
            Dim correctionLoop As Boolean
            correctionLoop = True
            
            Do While correctionLoop
                ' Spaltenanzeige mit fixen Abständen
                Dim msg As String
                msg = "Datengruppen:" & vbCrLf & _
                      "Anrede:             " & SpalteAnrede & vbCrLf & _
                      "Titel:                  " & SpalteTitel & vbCrLf & _
                      "Vorname:          " & SpalteVorname & vbCrLf & _
                      "Nachname:       " & SpalteNachname & vbCrLf & _
                      "Unternehmen:    " & SpalteUnternehmen & vbCrLf & vbCrLf & _
                      "E-Mail:               " & SpalteTo & vbCrLf & _
                      "CC:                     " & SpalteCC & vbCrLf & _
                      "BCC:                   " & SpalteBCC & vbCrLf & vbCrLf & _
                      "Betreff:              " & SpalteSubj & vbCrLf & _
                      "Anhang:           " & SpalteAttach & vbCrLf & _
                      "Sendezeitpunkt:   " & SpalteSendezeitpunkt & vbCrLf & vbCrLf & _
                      "Diese Spalten wurden zu den Kontaktinformationen gefunden. Sind Sie einverstanden?"
                
                confirmColumns = MsgBox(msg, vbQuestion + vbYesNoCancel, "Spaltenbestätigung")
                
                Select Case confirmColumns
                    Case vbYes
                        correctionLoop = False ' Beenden
                    Case vbCancel
                        GoTo Cleanup
                    Case vbNo
                        ' Aktualisierte Liste inkl. Unternehmen (Nummer 5)
                        Dim columnList As String
                        columnList = "Wählen Sie die Spalte zur Korrektur:" & vbCrLf & _
                                     "1 - Anrede" & vbCrLf & _
                                     "2 - Titel" & vbCrLf & _
                                     "3 - Vorname" & vbCrLf & _
                                     "4 - Nachname" & vbCrLf & _
                                     "5 - Unternehmen" & vbCrLf & _
                                     "6 - E-Mail" & vbCrLf & _
                                     "7 - CC" & vbCrLf & _
                                     "8 - BCC" & vbCrLf & _
                                     "9 - Betreff" & vbCrLf & _
                                     "10 - Anhang" & vbCrLf & _
                                     "11 - Sendezeitpunkt" & vbCrLf & vbCrLf

                        Dim selectedColumn As String
                        selectedColumn = InputBox(Prompt:=columnList & vbCrLf & "Geben Sie die Nummer der Spalte ein:", _
                                                  Title:="Spalte korrigieren")
                        
                        If selectedColumn = "" Then ' Abbruch über 'Abbrechen'-Button
                            GoTo Cleanup
                        End If
                        
                        Dim num As Integer
                        Dim neueSpalte As String
                        num = Val(selectedColumn)
                        
                        Select Case num
                            Case 1 To 10
                                Select Case num
                                    Case 1
                                        neueSpalte = InputBox("Neue Spalte für Anrede (z.B. A):", "Anrede")
                                        If neueSpalte = "" Then
                                            SpalteAnrede = ""
                                        Else
                                            SpalteAnrede = neueSpalte
                                        End If
                                    Case 2
                                        neueSpalte = InputBox("Neue Spalte für Titel (z.B. B):", "Titel")
                                        If neueSpalte = "" Then
                                            SpalteTitel = ""
                                        Else
                                            SpalteTitel = neueSpalte
                                        End If
                                    Case 3
                                        neueSpalte = InputBox("Neue Spalte für Vorname (z.B. C):", "Vorname")
                                        If neueSpalte = "" Then
                                            SpalteVorname = ""
                                        Else
                                            SpalteVorname = neueSpalte
                                        End If
                                    Case 4
                                        neueSpalte = InputBox("Neue Spalte für Nachname:", "Nachname")
                                        If neueSpalte = "" Then
                                            SpalteNachname = ""
                                        Else
                                            SpalteNachname = neueSpalte
                                        End If
                                    Case 5
                                        neueSpalte = InputBox("Neue Spalte für Unternehmen:", "Unternehmen")
                                        If neueSpalte = "" Then
                                            SpalteUnternehmen = ""
                                        Else
                                            SpalteUnternehmen = neueSpalte
                                        End If
                                    Case 6
                                        neueSpalte = InputBox("Neue Spalte für E-Mail:", "E-Mail")
                                        If neueSpalte = "" Then
                                            SpalteTo = ""
                                        Else
                                            SpalteTo = neueSpalte
                                        End If
                                    Case 7
                                        neueSpalte = InputBox("Neue Spalte für CC:", "CC")
                                        If neueSpalte = "" Then
                                            SpalteCC = ""
                                        Else
                                            SpalteCC = neueSpalte
                                        End If
                                    Case 8
                                        neueSpalte = InputBox("Neue Spalte für BCC:", "BCC")
                                        If neueSpalte = "" Then
                                            SpalteBCC = ""
                                        Else
                                            SpalteBCC = neueSpalte
                                        End If
                                    Case 9
                                        neueSpalte = InputBox("Neue Spalte für Betreff:", "Betreff")
                                        If neueSpalte = "" Then
                                            SpalteSubj = ""
                                        Else
                                            SpalteSubj = neueSpalte
                                        End If
                                    Case 10
                                        neueSpalte = InputBox("Neue Spalte für Anhänge:", "Anhang")
                                        If neueSpalte = "" Then
                                            SpalteAttach = ""
                                        Else
                                            SpalteAttach = neueSpalte
                                        End If
                                    Case 11
                                        neueSpalte = InputBox("Neue Spalte für Sendezeitpunkt:", "Sendezeitpunkt")
                                        If neueSpalte = "" Then
                                            SpalteSendezeitpunkt = ""
                                        Else
                                            SpalteSendezeitpunkt = neueSpalte
                                        End If
                                End Select
                            Case Else
                                MsgBox "Ungültige Eingabe!"
                        End Select
                End Select
            Loop
            
            '******************************************************************************
            ' ** 5. Anfangszeile des Excel-Arbeitsblattes (MODIFIKATIONSMÖGLICHKEIT: standardmäßig erst ab einer bestimmten Zeile beginnen lassen) **
            '******************************************************************************
            Dim useCustomStartRes As VbMsgBoxResult
            useCustomStartRes = MsgBox("Befinden sich die Kontaktinformationen ab der 2. Excelzeile?", vbYesNoCancel)
            If useCustomStartRes = vbCancel Then GoTo Cleanup

            Dim startRow As Long
            If useCustomStartRes = vbNo Then
                Dim inputRow As String
                inputRow = InputBox("Geben Sie die Startzeile ein (z.B. 2), beachten Sie jedoch die Kopfzeile mit:" & vbCrLf & _
                                    "(Zum Abbrechen das Feld leer lassen und OK klicken)", "Startzeile")
                If inputRow = "" Then
                    If MsgBox("Keine Eingabe. Möchten Sie den Vorgang abbrechen?", vbYesNoCancel) = vbYes Then GoTo Cleanup
                    startRow = 2
                Else
                    On Error Resume Next
                    startRow = CLng(inputRow)
                    On Error GoTo 0
                    If startRow < 1 Then
                        MsgBox "Ungültige Zeile. Standardwert 2 wird verwendet."
                        startRow = 2
                    End If
                End If
            Else
                startRow = 2
            End If
            
            '******************************************************************************
            ' ** 6. Anrede auswählen **
            '******************************************************************************
            Dim useCustomAnredeRes As VbMsgBoxResult
            useCustomAnredeRes = MsgBox("Möchten Sie die voreingestellte formelle Anrede übernehmen?" & vbCrLf & _
            "Diese wäre für Herr = 'Sehr geehrter Herr' und für Frau = 'Sehr geehrte Frau'.", vbYesNoCancel + vbQuestion, "Formelle Anrede")
            If useCustomAnredeRes = vbCancel Then GoTo Cleanup
            Dim useCustomAnrede As Boolean
            useCustomAnrede = (useCustomAnredeRes = vbYes)
            
            '******************************************************************************
            ' ** 7. E-Mail-Versandoption: Direkt versenden oder nur generieren lassen **
            '******************************************************************************
            Dim sendDirectlyRes As VbMsgBoxResult
            sendDirectlyRes = MsgBox("Möchten Sie die E-Mails direkt versenden? Wenn nein, dann werden die E-Mails nur generiert und Sie senden jede E-Mail einzeln ab.", vbYesNoCancel + vbQuestion, "Versandoption")
            If sendDirectlyRes = vbCancel Then GoTo Cleanup
            Dim sendDirectly As Boolean
            sendDirectly = (sendDirectlyRes = vbYes)

            If sendDirectly Then
                Dim confirmSend As VbMsgBoxResult
                confirmSend = MsgBox("Sind Sie sicher, dass alle E-Mails sofort nach ihrer Erstellung automatisch versendet werden sollen?", vbYesNoCancel + vbQuestion, "Bestätigung E-Mail-Versand")
                If confirmSend = vbCancel Then GoTo Cleanup
                If confirmSend = vbNo Then
                    sendDirectly = False
                End If
            End If
            
            '******************************************************************************
            ' ** 8. HTML-Formatierung (MODIFIKATIONSMÖGLICHKEIT: CSS-Stile hinzufügen) **
            '*****************************************************************************' Temporäre HTML-Datei erstellen
            Dim htmlContent As String
            htmlContent = ExportWordToHTML(ActiveDocument)

            ' Standard-Schriftart und -größe aus dem Dokument ermitteln
            Dim fontName As String
            Dim fontSize As Single
            fontName = ActiveDocument.Styles(wdStyleNormal).Font.Name
            fontSize = ActiveDocument.Styles(wdStyleNormal).Font.Size

            ' CSS-Styles einbetten
            Dim htmlTemplate As String
            Dim bodyContent As String
            Dim splitContent As Variant
            
            splitContent = Split(htmlContent, "<body>")
            If UBound(splitContent) >= 1 Then
                bodyContent = Split(splitContent(1), "</body>")(0)
            Else
                bodyContent = htmlContent ' Fallback: Verwende gesamten Content
            End If
            
            htmlTemplate = "<html>" & _
                        "<head>" & _
                        "<meta charset=""UTF-8"">" & _
                        "<style type=""text/css"">" & _
                        "body { font-family: " & fontName & "; font-size: " & fontSize & "pt; }" & _
                        "</style>" & _
                        "</head>" & _
                        "<body>" & bodyContent & "</body></html>"

            '******************************************************************************
            ' ** 9. Anhang-Validierung (MODIFIKATIONSMÖGLICHKEIT: erweiterter Umgang mit Netzwerkpfaden) **
            '******************************************************************************
            Dim fehlerListe As String
            Dim lastRow As Long
            lastRow = xlWS.Cells(xlWS.Rows.Count, 1).End(xlUp).Row
            
            For d = startRow To lastRow
                Dim DateipfadCheck As String
                DateipfadCheck = xlWS.Range(SpalteAttach & d).Value
                Dim arrFileNames() As String
                arrFileNames = Split(DateipfadCheck, ",")
                
                For Each file In arrFileNames
                    file = Trim(file)
                    ' Entferne führende und abschließende Anführungszeichen, falls vorhanden
                    If Left(file, 1) = """" Then file = Mid(file, 2)
                    If Right(file, 1) = """" Then file = Left(file, Len(file) - 1)
                    
                    If file <> "" Then
                        If Not fso.FileExists(file) Then
                            fehlerListe = fehlerListe & "Fehler: " & file & " existiert nicht (Zeile " & d & ")" & vbCrLf
                        End If
                    End If
                Next
            Next
            
            If fehlerListe <> "" Then
                MsgBox "Fehler in Anhängen gefunden:" & vbCrLf & fehlerListe
                ActiveDocument.Undo
                GoTo Cleanup
            End If
            
            '******************************************************************************
            ' ** 10. Platzhalter im Dokument ersetzen (MODIFIKATIONSMÖGLICHKEIT: Namen der Platzhalter anpassen) **
            '******************************************************************************
            Dim fehlerMeldung As String
            Dim sentCount As Integer
            sentCount = 0
            
            For i = startRow To lastRow
                Dim strTo As String, strSubj As String, strBody As String
                Dim strAnrede As String, strVorname As String, strNachname As String
                Dim strAttach As String, strCC As String, strBCC As String
                Dim strUnternehmen As String
                Dim strSendezeitpunkt As Variant
                
                ' Daten aus Excel lesen (mit Fehlertoleranz)
                If SpalteTo <> "" Then
                    strTo = xlWS.Range(SpalteTo & i).Value
                Else
                    strTo = ""
                End If
                If SpalteSubj <> "" Then
                    strSubj = xlWS.Range(SpalteSubj & i).Value
                Else
                    strSubj = ""
                End If
                If SpalteAnrede <> "" Then
                    strAnrede = xlWS.Range(SpalteAnrede & i).Value
                Else
                    strAnrede = ""
                End If
                If SpalteVorname <> "" Then
                    strVorname = xlWS.Range(SpalteVorname & i).Value
                Else
                    strVorname = ""
                End If
                If SpalteNachname <> "" Then
                    strNachname = xlWS.Range(SpalteNachname & i).Value
                Else
                    strNachname = ""
                End If
                If SpalteUnternehmen <> "" Then
                    strUnternehmen = xlWS.Range(SpalteUnternehmen & i).Value
                Else
                    strUnternehmen = ""
                End If
                If SpalteAttach <> "" Then
                    strAttach = xlWS.Range(SpalteAttach & i).Value
                Else
                    strAttach = ""
                End If
                If SpalteCC <> "" Then
                    strCC = xlWS.Range(SpalteCC & i).Value
                Else
                    strCC = ""
                End If
                If SpalteBCC <> "" Then
                    strBCC = xlWS.Range(SpalteBCC & i).Value
                Else
                    strBCC = ""
                End If
                If SpalteSendezeitpunkt <> "" Then
                    strSendezeitpunkt = xlWS.Range(SpalteSendezeitpunkt & i).Value
                Else
                    strSendezeitpunkt = "" ' Oder vbEmpty
                End If

                
                ' Anrede-Behandlung
                If useCustomAnrede Then
                    Select Case strAnrede
                        Case "Frau": strAnrede = "Sehr geehrte Frau"
                        Case "Herr": strAnrede = "Sehr geehrter Herr"
                        Case Else: strAnrede = ""
                    End Select
                Else
                    ' Bei Nein: Zellinhalt unverändert nutzen.
                    ' strAnrede bleibt wie gelesen.
                End If
                
                strBody = htmlTemplate

                ' Ersetze Platzhalter
                If SpalteAnrede <> "" Then strBody = Replace(strBody, "%Anrede%", strAnrede)
                If SpalteTitel <> "" Then strBody = Replace(strBody, "%Titel%", xlWS.Range(SpalteTitel & i).Value)
                If SpalteVorname <> "" Then strBody = Replace(strBody, "%Vorname%", strVorname)
                If SpalteNachname <> "" Then strBody = Replace(strBody, "%Nachname%", strNachname)
                If SpalteUnternehmen <> "" Then strBody = Replace(strBody, "%Unternehmen%", xlWS.Range(SpalteUnternehmen & i).Value)
                
                '******************************************************************************
                ' ** 11. E-Mail-Versand (MODIFIKATIONSMÖGLICHKEIT: Vorgang pausieren) **
                '******************************************************************************
                On Error Resume Next
                Set objMail = objOutlook.CreateItem(0)  ' Erstelle neue Mail-Instanz für jeden Durchlauf
                
                With objMail
                    .To = strTo
                    .CC = strCC
                    .BCC = strBCC
                    .Subject = strSubj
                    .HTMLBody = strBody     ' Hier strBody statt htmlTemplate verwenden
                    .BodyFormat = 2
                    
                    ' Anhänge hinzufügen
                    If strAttach <> "" Then
                        Dim attachArray() As String
                        attachArray = Split(strAttach, ",")
                        Dim attFile As Variant
                        For Each attFile In attachArray
                            attFile = Trim(attFile)
                            ' Entferne führende und abschließende Anführungszeichen, falls vorhanden
                            If Left(attFile, 1) = """" Then attFile = Mid(attFile, 2)
                            If Right(attFile, 1) = """" Then attFile = Left(attFile, Len(attFile) - 1)
                            If attFile <> "" Then
                                .Attachments.Add attFile
                            End If
                        Next attFile
                    End If

                    ' Sendezeitpunkt setzen
                    If SpalteSendezeitpunkt <> "" And strSendezeitpunkt <> "" And IsDate(strSendezeitpunkt) Then ' Sicherstellen, dass strSendezeitpunkt nicht leer ist und ein Datum enthält
                        If CDate(strSendezeitpunkt) > Now Then ' Prüfen, ob der Zeitpunkt in der Zukunft liegt
                            .DeferredDeliveryTime = CDate(strSendezeitpunkt)
                        Else
                            ' Optional: Hinweis, wenn der Zeitpunkt in der Vergangenheit liegt
                            MsgBox "Der Sendezeitpunkt für die E-Mail an " & strTo & " (Zeile " & i & ") liegt in der Vergangenheit und wurde nicht gesetzt." & vbCrLf & "Die E-Mail wird beim Klicken auf 'Senden' im angezeigten Fenster normal behandelt.", vbInformation
                        End If
                    ElseIf SpalteSendezeitpunkt <> "" And strSendezeitpunkt <> "" And Not IsDate(strSendezeitpunkt) Then
                        ' Optional: Hinweis, wenn der Wert kein gültiges Datum ist
                        MsgBox "Der Wert '" & strSendezeitpunkt & "' in der Spalte Sendezeitpunkt für die E-Mail an " & strTo & " (Zeile " & i & ") ist kein gültiges Datum und wird ignoriert.", vbInformation
                    End If
                    
                    ' E-Mail senden oder anzeigen
                    If sendDirectly Then
                        .Send
                    Else
                        .Display
                    End If
                End With
                
                If Err.Number <> 0 Then
                    fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & ": " & Err.Description
                Else
                    sentCount = sentCount + 1
                End If
                
                Set objMail = Nothing
                Err.Clear
            Next
            
            '******************************************************************************
            ' ** 12. Abschluss und Bereinigung **
            '******************************************************************************
            ' Am Ende: Alle Änderungen rückgängig machen
            On Error Resume Next 
            If Not objUndo Is Nothing Then
                objUndo.EndCustomRecord
                doc.Undo
            End If
            On Error GoTo 0

            If fehlerMeldung <> "" Then
                MsgBox "Fehler aufgetreten:" & vbCrLf & fehlerMeldung
            Else
                MsgBox "Erfolgreich: " & sentCount & " E-Mails " & IIf(sendDirectly, "gesendet", "generiert")
            End If
            
Cleanup:
            Set fso = Nothing
            Set objMail = Nothing
            Set objOutlook = Nothing
            
            ' Prüfe ob Excel-Objekte existieren bevor sie geschlossen werden
            If Not xlWB Is Nothing Then
                xlWB.Close SaveChanges:=False
            End If
            If Not xlApp Is Nothing Then
                xlApp.Quit
            End If
            
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
        Else
            MsgBox "Keine Datei ausgewählt"
        End If
    End With
End Sub

Function ExportWordToHTML(doc As Document) As String
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\" & "temp_email_" & Format(Now, "yyyymmddhhmmss") & ".html"
    
    ' Erstelle ein neues temporäres Dokument und kopiere den Inhalt
    Dim tempDoc As Document
    Set tempDoc = Documents.Add
    tempDoc.Range.FormattedText = doc.Range.FormattedText
    
    ' Speichere das temporäre Dokument als HTML
    tempDoc.SaveAs2 FileName:=tempPath, FileFormat:=wdFormatFilteredHTML
    
    ' Lies den HTML-Inhalt
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tempPath For Input As #fileNum
    ExportWordToHTML = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
    
    ' Schließe das temporäre Dokument ohne zu speichern
    tempDoc.Close SaveChanges:=False
    
    ' Sicheres Löschen mit FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Wiederhole bis zu 5x bei gesperrter Datei
    Dim i As Integer
    For i = 1 To 5
        On Error Resume Next
        fso.DeleteFile tempPath, True
        If Err.Number = 0 Then Exit For
        If Err.Number = 70 Then
            Sleep 1000
        Else
            Exit For
        End If
        On Error GoTo 0
    Next i
    
    Set fso = Nothing
End Function
