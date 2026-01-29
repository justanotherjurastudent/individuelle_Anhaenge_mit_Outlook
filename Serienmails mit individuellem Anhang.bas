'******************************************************************************
' ** MODIFIKATIONS-HINWEISE (ZUR EINFÜGUNG AM ANFANG DES CODES) **
'******************************************************************************
' 1. **Sprachanpassung**: Alle Texte in deutschen Dialogen können an Ihre Sprache angepasst werden (z.B. "E-Mail auswählen" in englisch).
' 2. **Spaltenanpassung**: Die Suchbegriffe für Spaltenköpfe (z.B. "Anrede", "E-Mail") können erweitert oder geändert werden.
' 3. **Formatierung**: Die HTML-Codezeichen für Formate (z.B. <b> für Fettdruck) können durch andere Tags ersetzt werden.
' 4. **Anhang-Validierung**: Die Anhang-Prüfung kann erweitert werden, um Ordnerpfade oder Netzwerkpfade zu unterstützen.
' 5. **Vorlagenverwaltung**: Mehrere E-Mail-Vorlagen könnten über ein Auswahlmenü eingeführt werden.
' 6. **Automatisierung**: Die E-Mail-Versendung könnte über einen Timer oder Terminplaner gesteuert werden.
' 7. **Dateianhänge**: Implementiert intelligente Aufteilung von Dateipfaden mit verbesserter Komma-Behandlung für Dateinamen wie "Alexander, Schneider.xlsx".
'******************************************************************************

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub SendEmailsFromWordWithExcelWithAbfrage()
    
    '******************************************************************************'
    ' ** ZENTRALES DEBUG-PROTOKOLL **'
    '******************************************************************************'
    Const MacroVersion As String = "2026-01-29-7"
    Debug.Print "=== E-MAIL-SERIENERSTELLUNG GESTARTET ==="
    Debug.Print "Startzeit: " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    Debug.Print "Word-Version: " & Application.Version

    On Error GoTo FatalError
    
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
        Debug.Print "Outlook-Status: Läuft"
    Else
        Debug.Print "Outlook-Status: Nicht gefunden"
    End If
    
    If Not outlookRunning Then
        LogAbort "Outlook nicht geöffnet"
        MsgBox "Outlook ist nicht geöffnet. Bitte starten Sie Outlook und versuchen Sie es erneut.", vbExclamation, "Outlook erforderlich"
        GoTo Cleanup
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
            Debug.Print "Excel-Datei: " & Pfad
            Debug.Print "Excel-Version: " & xlApp.Version
            Debug.Print "Arbeitsblätter verfügbar: " & xlWB.Worksheets.Count
            
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
                    LogAbort "Arbeitsblatt-Auswahl abgebrochen"
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
            Debug.Print "Ausgewähltes Arbeitsblatt: " & xlWS.Name & " (Nummer " & selectedSheet & ")"
            
            ' Restliche Initialisierung
            Set doc = ActiveDocument
            Set objOutlook = CreateObject("Outlook.Application")

            ' Optional: Absenderkonto festlegen (leer lassen = Standardkonto verwenden)
            Dim preferredAccountSmtp As String
            Dim preferredAccountDisplayName As String
            preferredAccountSmtp = ""        ' z.B. "max.mustermann@firma.de"
            preferredAccountDisplayName = "" ' z.B. "Max Mustermann"
            Dim autoSelectAccount As Boolean
            autoSelectAccount = True ' Fallback: erstes Konto verwenden, falls kein Standardkonto greift
            Dim forceDisplayForSend As Boolean
            forceDisplayForSend = True ' Für stabilen Direktversand: Inspector kurz anzeigen

            On Error Resume Next
            Debug.Print "Outlook-Konten vorhanden: " & objOutlook.Session.Accounts.Count
            Dim debugAcc As Object
            Dim debugIdx As Integer
            debugIdx = 1
            For Each debugAcc In objOutlook.Session.Accounts
                Debug.Print "  Konto " & debugIdx & ": " & debugAcc.DisplayName & " (" & debugAcc.SmtpAddress & ")"
                debugIdx = debugIdx + 1
            Next debugAcc
            On Error GoTo 0
            
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
                SpalteAnrede = InputBox("Spalte für ""Anrede"" (z.B. A oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteAnrede = Chr(AnredeRange.Column + 64)
            End If
            
            ' Titel-Spalte finden
            Dim TitelRange As Excel.Range
            Set TitelRange = xlWS.Cells.Find("Titel", LookIn:=xlValues, LookAt:=xlWhole)
            If TitelRange Is Nothing Then
                SpalteTitel = InputBox("Spalte für ""Titel"" (z.B. B oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteTitel = Chr(TitelRange.Column + 64)
            End If
            
            ' Vorname-Spalte finden
            Dim VornameRange As Excel.Range
            Set VornameRange = xlWS.Cells.Find("Vorname", LookIn:=xlValues, LookAt:=xlWhole)
            If VornameRange Is Nothing Then
                SpalteVorname = InputBox("Spalte für ""Vorname"" (z.B. C oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteVorname = Chr(VornameRange.Column + 64)
            End If
            
            ' Nachname-Spalte finden
            Dim NachnameRange As Excel.Range
            Set NachnameRange = xlWS.Cells.Find("Nachname", LookIn:=xlValues, LookAt:=xlWhole)
            If NachnameRange Is Nothing Then
                SpalteNachname = InputBox("Spalte für ""Nachname"" (z.B. D oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteNachname = Chr(NachnameRange.Column + 64)
            End If

            ' Unternehmen-Spalte finden
            Dim SuchbegriffeUnternehmen As Variant
            Dim term As Variant
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
                SpalteUnternehmen = InputBox("Spalte für ""Unternehmen"" (z.B. X oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
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
                SpalteTo = InputBox("Spalte für ""E-Mail"" (z.B. E):", "Spalte finden")
                If SpalteTo = "" Then
                    LogAbort "E-Mail-Spalte nicht definiert"
                    MsgBox "Die Spalte mit den E-Mail-Adressen muss definiert sein. Vorgang wurde abgebrochen.", vbExclamation
                    GoTo Cleanup
                End If
            End If

            ' CC-Spalte finden
            Dim CCRange As Excel.Range
            Dim SpalteCC As String
            Set CCRange = xlWS.Cells.Find("CC", LookIn:=xlValues, LookAt:=xlWhole)
            If CCRange Is Nothing Then
                SpalteCC = InputBox("Spalte für ""CC"" (z.B. H oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteCC = Chr(CCRange.Column + 64)
            End If

            ' BCC-Spalte finden
            Dim BCCRange As Excel.Range
            Dim SpalteBCC As String
            Set BCCRange = xlWS.Cells.Find("BCC", LookIn:=xlValues, LookAt:=xlWhole)
            If BCCRange Is Nothing Then
                SpalteBCC = InputBox("Spalte für ""BCC"" (z.B. I oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            Else
                SpalteBCC = Chr(BCCRange.Column + 64)
            End If
            
            ' Betreff-Spalte finden
            Dim BetreffRange As Excel.Range
            Set BetreffRange = xlWS.Cells.Find("Betreff", LookIn:=xlValues, LookAt:=xlWhole)
            If BetreffRange Is Nothing Then
                SpalteSubj = InputBox("Spalte für ""Betreff"" (z.B. F oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
                If SpalteSubj = "" Then
                    LogAbort "Betreff-Spalte nicht definiert"
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
                SpalteAttach = InputBox("Spalte für Anhänge (z.B. G oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
            End If

            ' Sendezeitpunkt-Spalte vorab suchen (ohne sofortige Abfrage)
            Dim SendezeitpunktRange As Excel.Range
            Dim lastRow As Long
            Dim d As Long
            Set SendezeitpunktRange = xlWS.Cells.Find("Sendezeitpunkt", LookIn:=xlValues, LookAt:=xlWhole)
            If Not SendezeitpunktRange Is Nothing Then
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
                      "Anrede:                   " & SpalteAnrede & vbCrLf & _
                      "Titel:                        " & SpalteTitel & vbCrLf & _
                      "Vorname:                " & SpalteVorname & vbCrLf & _
                      "Nachname:             " & SpalteNachname & vbCrLf & _
                      "Unternehmen:        " & SpalteUnternehmen & vbCrLf & vbCrLf & _
                      "E-Mail:                     " & SpalteTo & vbCrLf & _
                      "CC:                           " & SpalteCC & vbCrLf & _
                      "BCC:                         " & SpalteBCC & vbCrLf & vbCrLf & _
                      "Betreff:                    " & SpalteSubj & vbCrLf & _
                      "Anhang:                  " & SpalteAttach & vbCrLf & _
                      "Sendezeitpunkt:    " & SpalteSendezeitpunkt & vbCrLf & vbCrLf & _
                      "Diese Spalten wurden zu den Kontaktinformationen gefunden. Sind Sie einverstanden?"
                
                confirmColumns = MsgBox(msg, vbQuestion + vbYesNoCancel, "Spaltenbestätigung")
                
                Select Case confirmColumns
                    Case vbYes
                        correctionLoop = False ' Beenden
                    Case vbCancel
                        LogAbort "Spaltenbestätigung abgebrochen"
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
                            LogAbort "Spaltenkorrektur abgebrochen"
                            GoTo Cleanup
                        End If
                        
                        Dim num As Integer
                        Dim neueSpalte As String
                        num = Val(selectedColumn)
                        
                        Select Case num
                            Case 1 To 11
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
            If useCustomStartRes = vbCancel Then
                LogAbort "Startzeilen-Abfrage abgebrochen"
                GoTo Cleanup
            End If

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
            ' ** 5a. Mindestdaten-Check (FRÜHE PRÜFUNG: vor allen anderen Fragen) **
            '******************************************************************************
            ' Finde die letzte Zeile über alle Spalten hinweg, um leere Zellen in Spalte A zu berücksichtigen
            On Error Resume Next
            Dim lastRowSearch As Object
            Set lastRowSearch = xlWS.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not lastRowSearch Is Nothing Then
                lastRow = lastRowSearch.Row
            End If
            If Err.Number <> 0 Or lastRow = 0 Then
                lastRow = xlWS.UsedRange.Rows.Count + xlWS.UsedRange.Row - 1
            End If
            On Error GoTo 0
            
            Debug.Print "Zu verarbeitende Zeilen: " & (lastRow - startRow + 1) & " (Zeilen " & startRow & " bis " & lastRow & ")"

            ' Mindestdaten-Check (E-Mail-Adresse + Betreff) als Sammelmeldung
            Dim missingRows As Object
            Set missingRows = CreateObject("Scripting.Dictionary")
            Dim missingSummary As String
            Dim r As Long
            For r = startRow To lastRow
                Dim tmpTo As String
                Dim tmpSubj As String
                If SpalteTo <> "" Then
                    tmpTo = Trim(CStr(xlWS.Range(SpalteTo & r).Value))
                Else
                    tmpTo = ""
                End If
                If SpalteSubj <> "" Then
                    tmpSubj = Trim(CStr(xlWS.Range(SpalteSubj & r).Value))
                Else
                    tmpSubj = ""
                End If
                If tmpTo = "" Or tmpSubj = "" Then
                    Dim missDetails As String
                    missDetails = IIf(tmpTo = "", "E-Mail-Adresse", "") & _
                                  IIf(tmpTo = "" And tmpSubj = "", ", ", "") & _
                                  IIf(tmpSubj = "", "Betreff", "")
                    missingRows(CStr(r)) = missDetails
                    missingSummary = missingSummary & "Zeile " & r & ": " & missDetails & vbCrLf
                End If
            Next r
            If missingRows.Count > 0 Then
                Dim missingPrompt As String
                missingPrompt = "Es fehlen Mindestdaten in folgenden Zeilen:" & vbCrLf & vbCrLf & _
                                missingSummary & vbCrLf & _
                                "Möchten Sie fortfahren?" & vbCrLf & _
                                "(Die betroffenen E-Mails werden übersprungen)"
                If MsgBox(missingPrompt, vbExclamation + vbYesNo, "Fehlende Mindestdaten") = vbNo Then
                    LogAbort "Mindestdaten-Check abgebrochen"
                    GoTo Cleanup
                End If
            End If
            
            '******************************************************************************
            ' ** 6. Anrede auswählen **
            '******************************************************************************
            Dim useCustomAnredeRes As VbMsgBoxResult
            useCustomAnredeRes = MsgBox("Möchten Sie die voreingestellte formelle Anrede übernehmen?" & vbCrLf & _
            "Diese erkennt automatisch verschiedene Angaben (z.B. Herr, m, Mann / Frau, w, weiblich) " & vbCrLf & _
            "und wandelt sie in 'Sehr geehrter Herr' bzw. 'Sehr geehrte Frau' um.", vbYesNoCancel + vbQuestion, "Formelle Anrede")
            If useCustomAnredeRes = vbCancel Then
                LogAbort "Anrede-Auswahl abgebrochen"
                GoTo Cleanup
            End If
            Dim useCustomAnrede As Boolean
            useCustomAnrede = (useCustomAnredeRes = vbYes)
            
            '******************************************************************************
            ' ** 6a. Sendezeitpunkt-Spalte bestätigen (falls noch nicht gesetzt) **
            '******************************************************************************
            If SpalteSendezeitpunkt = "" Then
                Set SendezeitpunktRange = xlWS.Cells.Find("Sendezeitpunkt", LookIn:=xlValues, LookAt:=xlWhole)
                If SendezeitpunktRange Is Nothing Then
                    SpalteSendezeitpunkt = InputBox("Spalte für Sendezeitpunkt (z.B. K oder leer lassen, wenn nicht vorhanden):", "Spalte finden")
                Else
                    SpalteSendezeitpunkt = Chr(SendezeitpunktRange.Column + 64)
                End If
            End If

            '******************************************************************************
            ' ** 6b. Outlook-Konto auswählen (falls mehrere vorhanden) **
            '******************************************************************************
            Dim accountCount As Integer
            accountCount = 0
            On Error Resume Next
            accountCount = objOutlook.Session.Accounts.Count
            On Error GoTo 0
            
            Dim selectedAccount As Object
            Set selectedAccount = Nothing
            
            If accountCount > 1 Then
                ' Mehrere Konten vorhanden - User muss eines wählen
                Dim accountSelectionConfirmed As Boolean
                accountSelectionConfirmed = False
                
                Do While Not accountSelectionConfirmed
                    Dim accountList As String
                    accountList = "Folgende Outlook-Konten sind verfügbar:" & vbCrLf & vbCrLf
                    
                    Dim acc As Object
                    Dim accIdx As Integer
                    accIdx = 1
                    For Each acc In objOutlook.Session.Accounts
                        On Error Resume Next
                        Dim accDisplay As String
                        accDisplay = acc.DisplayName
                        If accDisplay = "" Then accDisplay = acc.SmtpAddress
                        If accDisplay = "" Then accDisplay = "(Unbekanntes Konto)"
                        accountList = accountList & accIdx & " - " & accDisplay & vbCrLf
                        accIdx = accIdx + 1
                        On Error GoTo 0
                    Next acc
                    
                    accountList = accountList & vbCrLf & "Wählen Sie das Konto aus, das verwendet werden soll:"
                    
                    Dim selectedAccountNum As String
                    selectedAccountNum = InputBox(Prompt:=accountList, Title:="Outlook-Konto auswählen", Default:="1")
                    
                    If selectedAccountNum = "" Then
                        LogAbort "Konto-Auswahl abgebrochen"
                        GoTo Cleanup
                    End If
                    
                    ' Validiere die Eingabe
                    Dim selectedAccNum As Integer
                    On Error Resume Next
                    selectedAccNum = CInt(selectedAccountNum)
                    On Error GoTo 0
                    
                    If selectedAccNum < 1 Or selectedAccNum > accountCount Then
                        MsgBox "Ungültige Auswahl. Bitte eine Zahl zwischen 1 und " & accountCount & " eingeben.", vbExclamation
                        ' Schleife wiederholt sich
                    Else
                        ' Hole das ausgewählte Konto
                        On Error Resume Next
                        Set selectedAccount = objOutlook.Session.Accounts.Item(selectedAccNum)
                        On Error GoTo 0
                        
                        If selectedAccount Is Nothing Then
                            MsgBox "Das ausgewählte Konto konnte nicht geladen werden.", vbExclamation
                            ' Schleife wiederholt sich
                        Else
                            ' Bestätigung anzeigen
                            accDisplay = selectedAccount.DisplayName
                            If accDisplay = "" Then accDisplay = selectedAccount.SmtpAddress
                            If accDisplay = "" Then accDisplay = "(Unbekanntes Konto)"
                            
                            Dim confirmAccount As VbMsgBoxResult
                            confirmAccount = MsgBox("Absender-Konto: " & accDisplay & vbCrLf & vbCrLf & _
                                                   "Dieser Mail-Account wird im Folgenden verwendet." & vbCrLf & _
                                                   "Ist das korrekt?", vbYesNoCancel + vbQuestion, "Konto-Bestätigung")
                            
                            If confirmAccount = vbCancel Then
                                LogAbort "Konto-Bestätigung abgebrochen"
                                GoTo Cleanup
                            ElseIf confirmAccount = vbNo Then
                                ' Zurück zur Kontoauswahl - Schleife wiederholt sich
                            Else
                                ' vbYes - Bestätigung akzeptiert
                                accountSelectionConfirmed = True
                            End If
                        End If
                    End If
                Loop
                
                ' Übernahme in preferredAccountSmtp und preferredAccountDisplayName (wird später bei .SendUsingAccount verwendet)
                preferredAccountSmtp = selectedAccount.SmtpAddress
                preferredAccountDisplayName = selectedAccount.DisplayName
            ElseIf accountCount = 1 Then
                ' Nur ein Konto vorhanden - automatisch verwenden
                Set selectedAccount = objOutlook.Session.Accounts.Item(1)
                On Error Resume Next
                preferredAccountSmtp = selectedAccount.SmtpAddress
                preferredAccountDisplayName = selectedAccount.DisplayName
                On Error GoTo 0
            End If

            '******************************************************************************
            ' ** 7. E-Mail-Versandoption: Direkt versenden oder nur generieren lassen **
            '******************************************************************************
            Dim sendDirectlyRes As VbMsgBoxResult
            sendDirectlyRes = MsgBox("Möchten Sie die E-Mails direkt versenden? Wenn nein, dann werden die E-Mails nur generiert und Sie senden jede E-Mail einzeln ab.", vbYesNoCancel + vbQuestion, "Versandoption")
            If sendDirectlyRes = vbCancel Then
                LogAbort "Versandoption abgebrochen"
                GoTo Cleanup
            End If
            Dim sendDirectly As Boolean
            sendDirectly = (sendDirectlyRes = vbYes)

            If sendDirectly Then
                Dim confirmSend As VbMsgBoxResult
                confirmSend = MsgBox("Sind Sie sicher, dass alle E-Mails sofort nach ihrer Erstellung automatisch versendet werden sollen?", vbYesNoCancel + vbQuestion, "Bestätigung E-Mail-Versand")
                If confirmSend = vbCancel Then
                    LogAbort "Versand-Bestätigung abgebrochen"
                    GoTo Cleanup
                End If
                If confirmSend = vbNo Then
                    sendDirectly = False
                End If
            End If
            
            ' Debug: Parameter-Zusammenfassung
            Debug.Print "Parameter:"
            Debug.Print "  - Anrede: " & IIf(useCustomAnrede, "Formell (Sehr geehrter Herr/Frau)", "Aus Excel übernehmen")
            Debug.Print "  - Versand: " & IIf(sendDirectly, "Direkt versenden", "Nur generieren")
            Debug.Print "  - Outlook-Konto: " & IIf(selectedAccount Is Nothing, "Standard", selectedAccount.DisplayName & " (" & selectedAccount.SmtpAddress & ")")
            
            '******************************************************************************
            ' ** 8. Word-Inhalt vorbereiten für E-Mail-Body **
            '******************************************************************************
            ' Hinweis: Es werden keine Änderungen am Originaldokument vorgenommen.
            ' Daher wird bewusst kein UndoRecord verwendet (würde sonst User-Änderungen rückgängig machen).

            '******************************************************************************
            ' ** 9. Anhang-Validierung (MODIFIKATIONSMÖGLICHKEIT: erweiterter Umgang mit Netzwerkpfaden) **
            '******************************************************************************
            Dim fehlerListe As String
            
            If SpalteAttach <> "" Then
                For d = startRow To lastRow
                    Dim DateipfadCheck As String
                    ' Berücksichtige Hyperlinks in Excel-Zellen
                    DateipfadCheck = GetCellFilePathsWithHyperlinks(xlWS, SpalteAttach & d)
                    Dim arrFileNames() As String
                    arrFileNames = SplitFilePathsSmart(DateipfadCheck)
                    
                    For Each file In arrFileNames
                        Dim cleanFile As String
                        cleanFile = NormalizeAttachmentPath(CStr(file), xlWS.Parent.Path)
                        
                        If cleanFile <> "" Then
                            Dim fileExists As Boolean
                            fileExists = fso.FileExists(cleanFile)
                            
                            If Not fileExists Then
                                fehlerListe = fehlerListe & "Fehler: " & cleanFile & " existiert nicht (Zeile " & d & ")" & vbCrLf
                            End If
                        End If
                    Next
                Next
            End If
            
            If fehlerListe <> "" Then
                Debug.Print "FEHLER in Anhang-Validierung:"
                Debug.Print fehlerListe

                ' Zähle die Anzahl der fehlerhaften Einträge
                Dim errorLines() As String
                errorLines = Split(fehlerListe, vbCrLf)
                Dim errorCount As Integer
                errorCount = 0
                Dim line As Variant
                For Each line In errorLines
                    If Trim(line) <> "" And Left(Trim(line), 6) = "Fehler" Then
                        errorCount = errorCount + 1
                    End If
                Next

                ' Vorgang abbrechen und User anweisen, Pfade zu korrigieren
                MsgBox "ACHTUNG: " & errorCount & " Datei(en) konnten nicht gefunden werden!" & vbCrLf & vbCrLf & _
                       "Folgende Dateien sind betroffen:" & vbCrLf & vbCrLf & _
                       fehlerListe & vbCrLf & vbCrLf & _
                       "Bitte korrigieren Sie die Dateipfade in der Excel-Tabelle oder entfernen Sie die fehlerhaften Einträge und starten Sie den Vorgang erneut.", _
                       vbExclamation, "Dateipfad-Fehler"

                GoTo Cleanup
            End If
            
            '******************************************************************************
            ' ** 10. Platzhalter im Dokument ersetzen (MODIFIKATIONSMÖGLICHKEIT: Namen der Platzhalter anpassen) **
            '******************************************************************************
            Dim fehlerMeldung As String
            Dim sentCount As Integer
            sentCount = 0
            
            For i = startRow To lastRow
                Debug.Print "Verarbeite Zeile " & i & " von " & lastRow & "..."
                Dim strTo As String, strSubj As String, strBody As String
                Dim strAnrede As String, strVorname As String, strNachname As String
                Dim strAttach As String, strCC As String, strBCC As String
                Dim strUnternehmen As String
                Dim strSendezeitpunkt As Variant
                Dim stepInfo As String
                stepInfo = "Init"
                
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
                ' Leerzeichen in Anrede ignorieren
                strAnrede = Trim(strAnrede)
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
                    ' Berücksichtige Hyperlinks in Excel-Zellen
                    strAttach = GetCellFilePathsWithHyperlinks(xlWS, SpalteAttach & i)
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

                ' Mindestdaten aus Sammelprüfung überspringen
                If missingRows.Exists(CStr(i)) Then
                    fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & " (Zeile " & i & "): Mindestdaten fehlen (" & missingRows(CStr(i)) & ")."
                    Debug.Print "FEHLER bei Zeile " & i & " (" & strVorname & " " & strNachname & "): Mindestdaten fehlen (" & missingRows(CStr(i)) & ")."
                    GoTo NextIteration
                End If

                
                ' Anrede-Behandlung
                If useCustomAnrede Then
                    Select Case LCase(strAnrede)
                        Case "frau", "w", "f", "weiblich"
                            strAnrede = "Sehr geehrte Frau"
                        Case "herr", "m", "mann", "männlich"
                            strAnrede = "Sehr geehrter Herr"
                        Case Else
                            strAnrede = ""
                    End Select
                Else
                    ' Bei Nein: Zellinhalt unverändert nutzen.
                    ' strAnrede bleibt wie gelesen.
                End If
                
                ' Temporäre Kopie des Dokuments erstellen und Platzhalter ersetzen
                Dim tempDoc As Document
                Set tempDoc = Nothing ' Initialisierung für Error-Handling
                
                On Error GoTo ErrorHandler
                Application.ScreenUpdating = False ' Bildschirmaktualisierung ausschalten
                Set tempDoc = Documents.Add(Visible:=False) ' Unsichtbares Dokument erstellen
                tempDoc.Range.FormattedText = doc.Range.FormattedText
                
                ' Ersetze Platzhalter im temporären Dokument
                With tempDoc.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Execute FindText:="%Anrede%", ReplaceWith:=strAnrede, Replace:=wdReplaceAll
                    If SpalteTitel <> "" Then
                        Dim titelVal As String
                        titelVal = Trim(xlWS.Range(SpalteTitel & i).Value)
                        If titelVal <> "" Then
                            .Execute FindText:="%Titel%", ReplaceWith:=titelVal, Replace:=wdReplaceAll
                        Else
                            .Execute FindText:="%Titel%", ReplaceWith:="", Replace:=wdReplaceAll
                        End If
                    Else
                        .Execute FindText:="%Titel%", ReplaceWith:="", Replace:=wdReplaceAll
                    End If
                    If SpalteVorname <> "" Then .Execute FindText:="%Vorname%", ReplaceWith:=strVorname, Replace:=wdReplaceAll
                    If SpalteNachname <> "" Then .Execute FindText:="%Nachname%", ReplaceWith:=strNachname, Replace:=wdReplaceAll
                    If SpalteUnternehmen <> "" Then .Execute FindText:="%Unternehmen%", ReplaceWith:=xlWS.Range(SpalteUnternehmen & i).Value, Replace:=wdReplaceAll
                End With
                
                ' Bereinigung von Leerzeichen (Doppelte Leerzeichen und Leerzeichen vor Satzzeichen)
                With tempDoc.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    
                    ' 1. Doppelte Leerzeichen durch einfache ersetzen
                    Do While .Execute(FindText:="  ", ReplaceWith:=" ", Replace:=wdReplaceAll)
                    Loop
                    
                    ' 2. Leerzeichen vor Satzzeichen entfernen
                    Dim satzzeichen As Variant
                    Dim zeichen As Variant
                    satzzeichen = Array(",", ".", "!", "?", ":", ";")
                    
                    For Each zeichen In satzzeichen
                        .Execute FindText:=" " & zeichen, ReplaceWith:=zeichen, Replace:=wdReplaceAll
                    Next zeichen
                End With
                
                '******************************************************************************
                ' ** 11. E-Mail-Versand (MODIFIKATIONSMÖGLICHKEIT: Vorgang pausieren) **
                '******************************************************************************
                stepInfo = "CreateItem"
                Set objMail = objOutlook.CreateItem(0)  ' Erstelle neue Mail-Instanz für jeden Durchlauf
                
                ' Absenderkonto direkt nach Erstellung setzen (wichtig: muss VOR dem Inspector erfolgen)
                stepInfo = "SetSendAccount"
                If preferredAccountSmtp <> "" Or preferredAccountDisplayName <> "" Then
                    Dim sendAccount As Object
                    Set sendAccount = FindOutlookAccount(objOutlook, preferredAccountSmtp, preferredAccountDisplayName)
                    If Not sendAccount Is Nothing Then
                        On Error Resume Next
                        Set objMail.SendUsingAccount = sendAccount
                        Debug.Print "Konto für Zeile " & i & " gesetzt: " & sendAccount.DisplayName & " (" & sendAccount.SmtpAddress & ")"
                        On Error GoTo ErrorHandler
                    Else
                        Debug.Print "WARNUNG Zeile " & i & ": Absenderkonto nicht gefunden (" & _
                                   IIf(preferredAccountSmtp <> "", preferredAccountSmtp, preferredAccountDisplayName) & ")"
                    End If
                End If
                
                With objMail
                    .To = strTo
                    .CC = strCC
                    .BCC = strBCC
                    .Subject = strSubj
                    .BodyFormat = 2 ' HTML-Format
                    
                    ' Word-Inhalt direkt in E-Mail übernehmen (behält Bilder bei)
                    ' Wichtig: Kein Clipboard-Paste, da Outlook/WordEditor sonst teils den Stil "Hyperlink"
                    ' (blau/unterstrichen) auf den eingefügten Text anwenden kann.
                    ' Hinweis: `Range.FormattedText = ...` kann scheitern, weil Word (Makro) und Outlook
                    ' WordEditor intern unterschiedliche Word-Instanzen nutzen (COM-Interface-Mismatch).
                    ' Daher nutzen wir Copy/Paste, aber erzwingen "Originalformatierung".
                    Dim editorDoc As Object
                    Dim insp As Object
                    stepInfo = "GetInspector"
                    Set insp = .GetInspector
                    ' WordEditor ist nur verfügbar, wenn der Inspector initialisiert ist
                    If insp Is Nothing Or insp.WordEditor Is Nothing Then
                        .Display
                        DoEvents
                        Set insp = .GetInspector
                    End If
                    Set editorDoc = insp.WordEditor
                    Dim insertRange As Object
                    Set insertRange = editorDoc.Range(0, 0)

                    On Error Resume Next
                    insertRange.Style = editorDoc.Styles(wdStyleNormal)
                    On Error GoTo ErrorHandler

                    stepInfo = "PasteContent"
                    tempDoc.Range.Copy
                    On Error Resume Next
                    insertRange.PasteAndFormat wdFormatOriginalFormatting
                    If Err.Number <> 0 Then
                        Err.Clear
                        insertRange.Paste
                    End If
                    On Error GoTo ErrorHandler
                    
                    ' Anhänge hinzufügen
                    stepInfo = "AddAttachments"
                    If strAttach <> "" Then
                        Dim attachArray() As String
                        attachArray = SplitFilePathsSmart(strAttach)
                        Dim attFile As Variant
                        For Each attFile In attachArray
                            Dim cleanAttFile As String
                            cleanAttFile = NormalizeAttachmentPath(CStr(attFile), xlWS.Parent.Path)
                            If cleanAttFile <> "" Then
                                .Attachments.Add cleanAttFile
                            End If
                        Next attFile
                    End If

                    ' Sendezeitpunkt setzen
                    If SpalteSendezeitpunkt <> "" And strSendezeitpunkt <> "" And IsDate(strSendezeitpunkt) Then ' Sicherstellen, dass strSendezeitpunkt nicht leer ist und ein Datum enthält
                        If CDate(strSendezeitpunkt) > Now Then ' Prüfen, ob der Zeitpunkt in der Zukunft liegt
                            .DeferredDeliveryTime = CDate(strSendezeitpunkt)
                        Else
                            ' Optional: Hinweis, wenn der Zeitpunkt in der Vergangenheit liegt
                            MsgBox "HINWEIS: Sendezeitpunkt in der Vergangenheit" & vbCrLf & vbCrLf & _
                                   "Die E-Mail an " & strTo & " (Zeile " & i & ") hat einen Sendezeitpunkt in der Vergangenheit." & vbCrLf & _
                                   "Der verzögerte Versand wurde daher nicht aktiviert." & vbCrLf & vbCrLf & _
                                   "Die E-Mail wird beim Klick auf 'Senden' sofort versendet.", _
                                   vbInformation, "Sendezeitpunkt-Hinweis"
                        End If
                    ElseIf SpalteSendezeitpunkt <> "" And strSendezeitpunkt <> "" And Not IsDate(strSendezeitpunkt) Then
                        ' Optional: Hinweis, wenn der Wert kein gültiges Datum ist
                        MsgBox "WARNUNG: Ungültiger Sendezeitpunkt" & vbCrLf & vbCrLf & _
                               "Der Wert '" & strSendezeitpunkt & "' in der Sendezeitpunkt-Spalte (Zeile " & i & ")" & vbCrLf & _
                               "für die E-Mail an " & strTo & " ist kein gültiges Datum." & vbCrLf & vbCrLf & _
                               "Der verzögerte Versand wird für diese E-Mail nicht aktiviert." & vbCrLf & _
                               "Die E-Mail wird beim Klick auf 'Senden' sofort versendet.", _
                               vbExclamation, "Ungültiger Sendezeitpunkt"
                    End If
                    
                    ' E-Mail senden oder anzeigen
                    If sendDirectly Then
                        ' Schutz gegen leere/ungültige Empfänger (verhindert Laufzeitfehler 5 bei .Send)
                        If Trim(.To) = "" Then
                            fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & " (Zeile " & i & "): Empfänger fehlt (To ist leer)."
                            Debug.Print "FEHLER bei Zeile " & i & " (" & strVorname & " " & strNachname & "): Empfänger fehlt (To ist leer)."
                            GoTo NextIteration
                        End If
                        If Not .Recipients.ResolveAll Then
                            fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & " (Zeile " & i & "): Empfänger konnte nicht aufgelöst werden (" & .To & ")."
                            Debug.Print "FEHLER bei Zeile " & i & " (" & strVorname & " " & strNachname & "): Empfänger konnte nicht aufgelöst werden (" & .To & ")."
                            GoTo NextIteration
                        End If
                        ' Konto ist bereits nach CreateItem gesetzt - hier nur noch Fallback für Standard-Konto
                        If autoSelectAccount Then
                            On Error Resume Next
                            If .SendUsingAccount Is Nothing Then
                                If objOutlook.Session.Accounts.Count > 0 Then
                                    Set .SendUsingAccount = objOutlook.Session.Accounts.Item(1)
                                    Debug.Print "INFO Zeile " & i & ": SendUsingAccount automatisch auf Konto 1 (Fallback) gesetzt."
                                End If
                            End If
                            On Error GoTo ErrorHandler
                        End If
                        If forceDisplayForSend Then
                            stepInfo = "DisplayForSend"
                            .Display
                            DoEvents
                            Sleep 200
                        End If
                        stepInfo = "Send"
                        On Error Resume Next
                        .Send
                        Dim sendErr As Long
                        Dim sendDesc As String
                        sendErr = Err.Number
                        sendDesc = Err.Description
                        On Error GoTo ErrorHandler
                        If sendErr <> 0 Then
                            If sendDesc = "" Then sendDesc = "Unbekannter Fehler"
                            fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & " (Zeile " & i & "): Versand fehlgeschlagen: " & sendDesc & " (Nr. " & sendErr & ")"
                            Debug.Print "FEHLER bei Zeile " & i & " (" & strVorname & " " & strNachname & "): Versand fehlgeschlagen: " & sendDesc & " (Nr. " & sendErr & ")"
                            Err.Clear
                            GoTo NextIteration
                        End If
                    Else
                        stepInfo = "Display"
                        .Display
                    End If
                End With
                
                ' Erfolgreich: Temporäres Dokument sicher schließen
                On Error Resume Next
                If Not tempDoc Is Nothing Then
                    tempDoc.Close SaveChanges:=False
                    Set tempDoc = Nothing
                End If
                Application.ScreenUpdating = True ' Bildschirmaktualisierung wieder einschalten
                On Error GoTo 0
                
                sentCount = sentCount + 1
                Set objMail = Nothing
                GoTo NextIteration
                
ErrorHandler:
                ' Fehlerbehandlung: Temporäres Dokument sicher schließen
                Application.ScreenUpdating = True ' Bildschirmaktualisierung wieder einschalten
                On Error Resume Next
                If Not tempDoc Is Nothing Then
                    tempDoc.Close SaveChanges:=False
                    Set tempDoc = Nothing
                    Debug.Print "Temporäres Dokument nach Fehler geschlossen (Zeile " & i & ")"
                End If
                Set objMail = Nothing
                On Error GoTo 0
                
                Dim errDesc As String
                errDesc = Err.Description
                If errDesc = "" Then errDesc = "Unbekannter Fehler"
                fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler bei " & strVorname & " " & strNachname & " (Zeile " & i & "): " & errDesc & " (Nr. " & Err.Number & ", Schritt: " & stepInfo & ")"
                If strAttach <> "" Then
                    fehlerMeldung = fehlerMeldung & vbCrLf & "  Betroffene Datei(en): " & strAttach
                End If
                Debug.Print "FEHLER bei Zeile " & i & " (" & strVorname & " " & strNachname & "): " & errDesc & " (Nr. " & Err.Number & ", Schritt: " & stepInfo & ")"
                If strAttach <> "" Then
                    Debug.Print "  Betroffene Datei(en): " & strAttach
                End If
                Err.Clear
                
NextIteration:
            Next
            
            '******************************************************************************
            ' ** 12. Abschluss und Bereinigung **
            '******************************************************************************
            ' Keine Undo-Bereinigung nötig, da das Originaldokument nicht verändert wird.

            If fehlerMeldung <> "" Then
                Debug.Print "FEHLER bei der Verarbeitung:"
                Debug.Print fehlerMeldung
                MsgBox "ACHTUNG: Fehler bei der E-Mail-Verarbeitung!" & vbCrLf & vbCrLf & _
                       "Folgende Probleme sind aufgetreten:" & vbCrLf & vbCrLf & _
                       fehlerMeldung & vbCrLf & vbCrLf & _
                       "Die fehlerhaften E-Mails wurden übersprungen. " & sentCount & " E-Mails wurden erfolgreich verarbeitet.", _
                       vbExclamation, "Versand/Generierung trotz Fehler"
            Else
                Debug.Print "=== VERARBEITUNG ABGESCHLOSSEN ==="
                Debug.Print "Erfolgreich verarbeitet: " & sentCount & " E-Mails"
                Debug.Print "Versandmodus: " & IIf(sendDirectly, "Direkt versendet", "Nur generiert")
                Debug.Print "Endzeit: " & Format(Now, "dd.mm.yyyy hh:nn:ss")
                MsgBox "ERFOLG: E-Mail-Serienversand abgeschlossen!" & vbCrLf & vbCrLf & _
                       sentCount & " E-Mails wurden erfolgreich " & IIf(sendDirectly, "versendet", "generiert") & ".", _
                       vbInformation, "Vorgang erfolgreich abgeschlossen"
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
            LogAbort "Keine Datei ausgewählt"
            MsgBox "Keine Datei ausgewählt"
        End If
    End With
    Exit Sub

FatalError:
    Debug.Print "ABBRUCH: Unerwarteter Fehler: " & Err.Description & " (Nr. " & Err.Number & ")"
    Resume Cleanup
End Sub

Sub LogAbort(reason As String)
    Debug.Print "ABBRUCH: " & reason
End Sub

Function FindOutlookAccount(objOutlook As Object, Optional smtp As String = "", Optional displayName As String = "") As Object
    '******************************************************************************
    ' ** Outlook-Konto anhand SMTP-Adresse oder Anzeigename finden **
    '******************************************************************************
    If objOutlook Is Nothing Then Exit Function
    If smtp = "" And displayName = "" Then Exit Function

    Dim acc As Object
    For Each acc In objOutlook.Session.Accounts
        If smtp <> "" Then
            If LCase(acc.SmtpAddress) = LCase(smtp) Then
                Set FindOutlookAccount = acc
                Exit Function
            End If
        ElseIf displayName <> "" Then
            If LCase(acc.DisplayName) = LCase(displayName) Then
                Set FindOutlookAccount = acc
                Exit Function
            End If
        End If
    Next acc
End Function


Function SplitFilePathsSmart(filePaths As String) As String()
    '******************************************************************************
    ' ** Intelligente Aufteilung von Dateipfaden **
    ' ** Unterscheidet zwischen Kommas in Dateinamen und Pfad-Trennzeichen **
    '******************************************************************************
    Dim result() As String
    Dim resultCount As Integer
    resultCount = 0
    
    If Trim(filePaths) = "" Then
        ReDim result(0)
        result(0) = ""
        SplitFilePathsSmart = result
        Exit Function
    End If
    
    ' Erstelle ein vorläufiges Array mit ausreichend Platz
    ReDim result(100)
    
    Dim i As Integer
    Dim currentPath As String
    Dim inQuotes As Boolean
    Dim char As String
    currentPath = ""
    inQuotes = False
    
    ' Gehe durch jeden Charakter
    For i = 1 To Len(filePaths)
        char = Mid(filePaths, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
            currentPath = currentPath & char
        ElseIf (char = "," Or char = ";") And Not inQuotes Then
            ' Prüfe ob das Trennzeichen ein echter Pfad-Trennzeichen ist
            Dim restOfString As String
            restOfString = Mid(filePaths, i + 1)
            
            ' Entferne führende Leerzeichen für die Prüfung
            Dim trimmedRest As String
            trimmedRest = LTrim(restOfString)
            
            Dim isPathSeparator As Boolean
            isPathSeparator = False
            
            If Len(trimmedRest) > 0 Then
                ' Fall 1: Trenner vor Dateierweiterung (z.B. "file;.pdf")
                If Left(trimmedRest, 1) = "." Then
                    isPathSeparator = False
                ' Fall 2: Trenner vor neuem Pfad mit Laufwerksbuchstabe (z.B. "C:\" oder "C:/")
                ElseIf Len(trimmedRest) >= 3 And (Mid(trimmedRest, 2, 2) = ":\" Or Mid(trimmedRest, 2, 2) = ":/") Then
                    isPathSeparator = True
                ' Fall 3: Trenner vor UNC-Pfad (z.B. "\\server\" oder "//server/")
                ElseIf Len(trimmedRest) >= 2 And (Left(trimmedRest, 2) = "\\" Or Left(trimmedRest, 2) = "//") Then
                    isPathSeparator = True
                ' Fall 4: Trenner vor Anführungszeichen (neuer quoted Pfad)
                ElseIf Left(trimmedRest, 1) = """" Then
                    isPathSeparator = True
                ' Fall 5: Trenner vor absolutem Pfad (z.B. "\folder\" oder "/folder/")
                ElseIf Left(trimmedRest, 1) = "\" Or Left(trimmedRest, 1) = "/" Then
                    isPathSeparator = True
                Else
                    ' Fallback: Prüfe ob der aktuelle Pfad bereits vollständig aussieht
                    ' (hat eine Dateierweiterung am Ende)
                    Dim currentTrimmed As String
                    currentTrimmed = Trim(currentPath)
                    If HasFileExtension(currentTrimmed) Then
                        isPathSeparator = True
                    Else
                        isPathSeparator = False
                    End If
                End If
            End If
            
            If isPathSeparator Then
                ' Das Komma ist ein echter Trennzeichen
                If Trim(currentPath) <> "" Then
                    result(resultCount) = Trim(currentPath)
                    resultCount = resultCount + 1
                End If
                currentPath = ""
            Else
                ' Das Komma ist Teil des Dateinamens
                currentPath = currentPath & char
            End If
        Else
            currentPath = currentPath & char
        End If
    Next i
    
    ' Füge den letzten Pfad hinzu, falls vorhanden
    If Trim(currentPath) <> "" Then
        result(resultCount) = Trim(currentPath)
        resultCount = resultCount + 1
    End If
    
    ' Redimensioniere das Array auf die tatsächliche Größe
    ReDim Preserve result(IIf(resultCount = 0, 0, resultCount - 1))
    
    SplitFilePathsSmart = result
End Function

Function HasFileExtension(filePath As String) As Boolean
    '******************************************************************************
    ' ** Hilfsfunktion: Prüft ob ein Pfad eine Dateierweiterung hat **
    '******************************************************************************
    If Len(filePath) = 0 Then
        HasFileExtension = False
        Exit Function
    End If
    
    ' Entferne Anführungszeichen falls vorhanden
    Dim cleanPath As String
    cleanPath = filePath
    If Left(cleanPath, 1) = """" And Right(cleanPath, 1) = """" Then
        cleanPath = Mid(cleanPath, 2, Len(cleanPath) - 2)
    End If
    
    ' Suche nach dem letzten Punkt und Pfadtrennern
    Dim lastDotPos As Integer
    Dim lastSlashPos As Integer
    Dim lastBackslashPos As Integer
    lastDotPos = InStrRev(cleanPath, ".")
    lastSlashPos = InStrRev(cleanPath, "/")
    lastBackslashPos = InStrRev(cleanPath, "\")
    
    ' Ein Punkt muss vorhanden sein und nach dem letzten Pfadtrenner kommen
    If lastDotPos > 0 And lastDotPos > lastSlashPos And lastDotPos > lastBackslashPos Then
        ' Prüfe ob nach dem Punkt noch 1-4 Zeichen kommen (typische Erweiterung)
        Dim extensionLength As Integer
        extensionLength = Len(cleanPath) - lastDotPos
        If extensionLength >= 1 And extensionLength <= 4 Then
            HasFileExtension = True
        Else
            HasFileExtension = False
        End If
    Else
        HasFileExtension = False
    End If
End Function

Function GetCellFilePathsWithHyperlinks(ws As Excel.Worksheet, cellAddress As String) As String
    '******************************************************************************
    ' ** Holt Dateipfade aus Excel-Zelle und berücksichtigt Hyperlinks **
    '******************************************************************************
    Dim result As String
    Dim targetRange As Excel.Range
    Set targetRange = ws.Range(cellAddress)
    
    ' Prüfe ob die Zelle Hyperlinks enthält
    If targetRange.Hyperlinks.Count > 0 Then
        Dim hyperlink As Excel.Hyperlink
        Set hyperlink = targetRange.Hyperlinks(1)
        
        ' Debug: Nur bei Fehlern oder wichtigen Änderungen
        If hyperlink.Address <> hyperlink.TextToDisplay Then
            Debug.Print "Hyperlink in " & cellAddress & ": '" & hyperlink.Address & "'"
        End If
        
        ' Versuche verschiedene Quellen für den Dateipfad
        If hyperlink.Address <> "" And hyperlink.Address <> hyperlink.TextToDisplay Then
            ' Verwende Address, wenn es sich vom angezeigten Text unterscheidet
            result = hyperlink.Address
        ElseIf hyperlink.SubAddress <> "" Then
            ' Verwende SubAddress, falls Address leer ist
            result = hyperlink.SubAddress
        ElseIf InStr(targetRange.Formula, "file:") > 0 Then
            ' Versuche URL aus der Zellformel zu extrahieren
            result = ExtractFileUrlFromFormula(targetRange.Formula)
        Else
            ' Fallback: Verwende angezeigten Text
            result = hyperlink.TextToDisplay
        End If
        
    Else
        ' Kein Hyperlink - prüfe trotzdem die Formel
        If InStr(targetRange.Formula, "file:") > 0 Then
            result = ExtractFileUrlFromFormula(targetRange.Formula)
        Else
            ' Verwende normalen Zellwert
            result = targetRange.Value
        End If

        ' Leere Zellen dürfen NICHT verarbeitet werden
        If IsNoAttachmentToken(result) Then
            GetCellFilePathsWithHyperlinks = ""
            Exit Function
        End If
    End If
    
    GetCellFilePathsWithHyperlinks = result
End Function

Function ConvertRelativeToAbsolutePath(filePath As String, basePath As String) As String
    '******************************************************************************
    ' ** Wandelt relative Pfade in absolute Pfade um **
    '******************************************************************************
    Dim result As String
    result = Trim(filePath)

    ' WICHTIG: Leere Werte / Platzhalter bedeuten "kein Anhang"
    If IsNoAttachmentToken(result) Then
        ConvertRelativeToAbsolutePath = ""
        Exit Function
    End If
    
    ' Hilfsvariable für die Prüfung ohne Anführungszeichen
    Dim checkPath As String
    checkPath = result
    If Left(checkPath, 1) = """" Then checkPath = Mid(checkPath, 2)
    
    ' Wenn schon absoluter Pfad oder file:/// URL, keine Änderung nötig
    If Len(checkPath) >= 3 And (Mid(checkPath, 2, 2) = ":\" Or Mid(checkPath, 2, 2) = ":/") Then
        ' Schon absoluter Pfad (z.B. C:\... oder C:/...)
        ConvertRelativeToAbsolutePath = result
        Exit Function
    ElseIf LCase(Left(checkPath, 4)) = "file" Then
        ' Schon file:/// URL
        ConvertRelativeToAbsolutePath = result
        Exit Function
    ElseIf Left(checkPath, 2) = "\\" Or Left(checkPath, 2) = "//" Then
        ' UNC-Pfad (z.B. \\server\... oder //server/...)
        ConvertRelativeToAbsolutePath = result
        Exit Function
    End If
    
    ' Relativer Pfad - umwandeln
    ' Verwende Excel's eigene Funktion für absolute Pfade
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Kombiniere Basis-Pfad mit relativem Pfad
    Dim combinedPath As String
    If Right(basePath, 1) <> "\" Then
        combinedPath = basePath & "\" & result
    Else
        combinedPath = basePath & result
    End If
    
    ' Normalisiere den Pfad (löst ..\ auf)
    result = fso.GetAbsolutePathName(combinedPath)
    
    On Error GoTo 0
    ConvertRelativeToAbsolutePath = result
End Function

Function ExtractFileUrlFromFormula(formula As String) As String
    '******************************************************************************
    ' ** Extrahiert file:/// URLs aus Excel-Formeln **
    '******************************************************************************
    Dim result As String
    result = ""
    
    ' Suche nach file: URLs in der Formel
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = InStr(LCase(formula), "file:")
    If startPos > 0 Then
        ' Finde das Ende der URL (nächstes Anführungszeichen oder Komma)
        endPos = startPos
        Do While endPos <= Len(formula)
            Dim char As String
            char = Mid(formula, endPos, 1)
            If char = """" Or char = "," Or char = ")" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        
        ' Extrahiere die URL
        result = Mid(formula, startPos, endPos - startPos)
    End If
    
    ExtractFileUrlFromFormula = result
End Function

Function ConvertFileUrlToPath(fileUrl As String) As String
    '******************************************************************************
    ' ** Konvertiert file:/// URLs in normale Dateipfade **
    '******************************************************************************
    Dim result As String
    result = Trim(fileUrl)
    
    ' Entferne file:/// Präfix (case-insensitive)
    If LCase(Left(result, 8)) = "file:///" Then
        result = Mid(result, 9)
    ElseIf LCase(Left(result, 7)) = "file://" Then
        result = Mid(result, 8)
    ElseIf LCase(Left(result, 5)) = "file:" Then
        result = Mid(result, 6)
    End If
    
    ' Ersetze / durch \ für Windows-Pfade
    result = Replace(result, "/", "\")
    
    ' URL-Dekodierung für in Windows-Dateinamen erlaubte Zeichen
    result = Replace(result, "%20", " ")  ' Leerzeichen (häufig)
    result = Replace(result, "%27", "'")  ' Apostroph
    result = Replace(result, "%28", "(")  ' Klammer auf
    result = Replace(result, "%29", ")")  ' Klammer zu
    result = Replace(result, "%2B", "+")  ' Plus
    result = Replace(result, "%2C", ",")  ' Komma
    result = Replace(result, "%2D", "-")  ' Bindestrich
    result = Replace(result, "%2E", ".")  ' Punkt
    result = Replace(result, "%3D", "=")  ' Gleichzeichen
    result = Replace(result, "%40", "@")  ' At-Zeichen
    result = Replace(result, "%5B", "[")  ' Eckige Klammer auf
    result = Replace(result, "%5D", "]")  ' Eckige Klammer zu
    result = Replace(result, "%5F", "_")  ' Unterstrich
    result = Replace(result, "%60", "`")  ' Backtick
    result = Replace(result, "%7B", "{")  ' Geschweifte Klammer auf
    result = Replace(result, "%7D", "}")  ' Geschweifte Klammer zu
    result = Replace(result, "%7E", "~")  ' Tilde
    
    ' Deutsche Umlaute
    result = Replace(result, "%C3%A4", "ä")  ' ä
    result = Replace(result, "%C3%B6", "ö")  ' ö
    result = Replace(result, "%C3%BC", "ü")  ' ü
    result = Replace(result, "%C3%84", "Ä")  ' Ä
    result = Replace(result, "%C3%96", "Ö")  ' Ö
    result = Replace(result, "%C3%9C", "Ü")  ' Ü
    result = Replace(result, "%C3%9F", "ß")  ' ß
    
    ConvertFileUrlToPath = result
End Function

Function IsNoAttachmentToken(value As String) As Boolean
    '******************************************************************************
    ' ** Erlaubt E-Mails ohne Anhang: Leere Werte und Platzhalter werden ignoriert **
    '******************************************************************************
    Dim s As String
    s = LCase(Trim(value))
    
    Select Case s
        Case "", "-", "–", "—", "kein", "keine", "keiner", "ohne", "n/a", "na", "null"
            IsNoAttachmentToken = True
        Case Else
            IsNoAttachmentToken = False
    End Select
End Function

Function LooksLikeFileReference(value As String) As Boolean
    '******************************************************************************
    ' ** Heuristik: Nur plausible Dateireferenzen werden als Anhang interpretiert **
    '******************************************************************************
    Dim t As String
    t = Trim(value)
    
    If t = "" Then
        LooksLikeFileReference = False
        Exit Function
    End If
    
    If LCase(Left(t, 4)) = "file" Then
        LooksLikeFileReference = True
        Exit Function
    End If
    
    If Len(t) >= 3 And (Mid(t, 2, 2) = ":\" Or Mid(t, 2, 2) = ":/") Then
        LooksLikeFileReference = True
        Exit Function
    End If
    
    If Left(t, 2) = "\\" Or Left(t, 2) = "//" Then
        LooksLikeFileReference = True
        Exit Function
    End If
    
    If InStr(t, "\") > 0 Or InStr(t, "/") > 0 Then
        LooksLikeFileReference = True
        Exit Function
    End If
    
    If HasFileExtension(t) Then
        LooksLikeFileReference = True
        Exit Function
    End If
    
    LooksLikeFileReference = False
End Function

Function NormalizeAttachmentPath(rawValue As String, basePath As String) As String
    '******************************************************************************
    ' ** Normalisiert einen Anhangseintrag; gibt "" zurück wenn kein Anhang gemeint **
    '******************************************************************************
    Dim s As String
    s = Trim(rawValue)
    
    If IsNoAttachmentToken(s) Then
        NormalizeAttachmentPath = ""
        Exit Function
    End If
    
    ' Konvertiere file:/// URLs zu normalen Pfaden
    s = ConvertFileUrlToPath(s)
    
    ' Entferne führende und abschließende Anführungszeichen, falls vorhanden
    If Left(s, 1) = """" Then s = Mid(s, 2)
    If Right(s, 1) = """" Then s = Left(s, Len(s) - 1)
    s = Trim(s)
    
    If IsNoAttachmentToken(s) Then
        NormalizeAttachmentPath = ""
        Exit Function
    End If
    
    ' Wenn der Zellinhalt kein plausibler Dateipfad ist (z.B. Freitext), ignorieren.
    If Not LooksLikeFileReference(s) Then
        NormalizeAttachmentPath = ""
        Exit Function
    End If
    
    ' Relative Pfade erst jetzt umwandeln (pro Eintrag, nicht als Gesamtkette)
    s = ConvertRelativeToAbsolutePath(s, basePath)

    ' Ordner sind keine Anhänge (Outlook erwartet Dateien)
    If Right(s, 1) = "\" Or Right(s, 1) = "/" Then
        NormalizeAttachmentPath = ""
        Exit Function
    End If
    
    NormalizeAttachmentPath = Trim(s)
End Function


