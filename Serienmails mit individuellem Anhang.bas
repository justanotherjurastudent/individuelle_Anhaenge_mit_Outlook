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
            Dim SpalteUnternehmen As String  ' Neue Variable für die Unternehmen-Spalte
            
            ' Anrede-Spalte finden
            Dim AnredeRange As Excel.Range
            Set AnredeRange = xlWS.Range("A1:Z1").Find("Anrede", LookIn:=xlValues, LookAt:=xlWhole)
            If AnredeRange Is Nothing Then
                SpalteAnrede = InputBox("Spalte für Anrede (z.B. A oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteAnrede = Chr(AnredeRange.Column + 64)
            End If
            
            ' Titel-Spalte finden
            Dim TitelRange As Excel.Range
            Set TitelRange = xlWS.Cells.Find("Titel", LookIn:=xlValues, LookAt:=xlWhole)
            If TitelRange Is Nothing Then
                SpalteTitel = InputBox("Spalte für Titel (z.B. B oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteTitel = Chr(TitelRange.Column + 64)
            End If
            
            ' Vorname-Spalte finden
            Dim VornameRange As Excel.Range
            Set VornameRange = xlWS.Cells.Find("Vorname", LookIn:=xlValues, LookAt:=xlWhole)
            If VornameRange Is Nothing Then
                SpalteVorname = InputBox("Spalte für Vornamen (z.B. C oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteVorname = Chr(VornameRange.Column + 64)
            End If
            
            ' Nachname-Spalte finden
            Dim NachnameRange As Excel.Range
            Set NachnameRange = xlWS.Cells.Find("Nachname", LookIn:=xlValues, LookAt:=xlWhole)
            If NachnameRange Is Nothing Then
                SpalteNachname = InputBox("Spalte für Nachnamen (z.B. D oder leer lassen, wenn nicht vorhanden):")
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
                SpalteUnternehmen = InputBox("Spalte für Unternehmen (z.B. X oder leer lassen, wenn nicht vorhanden):")
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
                SpalteTo = InputBox("Spalte für E-Mail (z.B. E):")
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
                SpalteCC = InputBox("Spalte für CC (z.B. H oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteCC = Chr(CCRange.Column + 64)
            End If

            ' BCC-Spalte finden
            Dim BCCRange As Excel.Range
            Dim SpalteBCC As String
            Set BCCRange = xlWS.Cells.Find("BCC", LookIn:=xlValues, LookAt:=xlWhole)
            If BCCRange Is Nothing Then
                SpalteBCC = InputBox("Spalte für BCC (z.B. I oder leer lassen, wenn nicht vorhanden):")
            Else
                SpalteBCC = Chr(BCCRange.Column + 64)
            End If
            
            ' Betreff-Spalte finden
            Dim BetreffRange As Excel.Range
            Set BetreffRange = xlWS.Cells.Find("Betreff", LookIn:=xlValues, LookAt:=xlWhole)
            If BetreffRange Is Nothing Then
                SpalteSubj = InputBox("Spalte für Betreff (z.B. F oder leer lassen, wenn nicht vorhanden):")
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
                      "Anhang:           " & SpalteAttach & vbCrLf & vbCrLf & _
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
                                     "10 - Anhang" & vbCrLf & vbCrLf

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
            '******************************************************************************
            ' UndoRecord für Rückgängig-Block starten
            On Error Resume Next
            If Not objUndo Is Nothing Then
                objUndo.StartCustomRecord "Serienmail-Änderungen"
            End If
            On Error GoTo 0

            ' Umschreiben von Hyperlinks
            For Each HL In ActiveDocument.Hyperlinks
                HL.Range.Text = "<a href=""" & HL.Address & """>" & HL.Range.Text & "</a>"
            Next

            ' HTML-Tags für Formate (Fett, Kursiv, etc.)
            Dim rng As Range
            Set rng = doc.Content

            ' Erst Formatierungen verarbeiten
            Dim formatSettings() As Variant
            formatSettings = Array( _
                Array("Bold", True, "<b>", "</b>"), _
                Array("Italic", True, "<i>", "</i>"), _
                Array("Underline", True, "<u>", "</u>"), _
                Array("Superscript", True, "<sup>", "</sup>"), _
                Array("Subscript", True, "<sub>", "</sub>"), _
                Array("SmallCaps", True, "<small>", "</small>"), _
                Array("AllCaps", True, "<big>", "</big>") _
            )

            ' Weitere Formatierungen verarbeiten
            For Each setting In formatSettings
                With rng.Find
                    .ClearFormatting
                    .Forward = True
                    .Wrap = wdFindStop
                    Select Case setting(0)
                        Case "Bold": .Font.Bold = setting(1)
                        Case "Italic": .Font.Italic = setting(1)
                        Case "Underline": .Font.Underline = setting(1)
                        Case "Superscript": .Font.Superscript = setting(1)
                        Case "Subscript": .Font.Subscript = setting(1)
                        Case "SmallCaps": .Font.SmallCaps = setting(1)
                        Case "AllCaps": .Font.AllCaps = setting(1)
                    End Select
                    .Text = ""
                    
                    Do While .Execute
                        rng.InsertBefore setting(2)
                        rng.InsertAfter setting(3)
                        rng.Collapse wdCollapseEnd
                    Loop
                End With
            Next

            ' Farben und Hervorhebungen verarbeiten (Separater Range)
            Dim rngFontColor As Range
            Set rngFontColor = doc.Content
            Dim rngMarkColor As Range
            Set rngMarkColor = doc.Content

            ' Farben-Einstellungen
            Dim colorSettings As Variant
            colorSettings = Array( _
                Array(wdColorRed, "#FF0000", "color"), _
                Array(wdColorBlue, "#0000FF", "color"), _
                Array(wdColorGreen, "#008000", "color"), _
                Array(wdColorYellow, "#FFFF00", "color"), _
                Array(wdColorMagenta, "#FF00FF", "color"), _
                Array(wdColorCyan, "#00FFFF", "color") _
            )

            ' Hervorhebungsfarben-Zuordnung
            Dim highlightColors As Variant
            highlightColors = Array( _
                Array(wdYellow, "#FFFF00"), _
                Array(wdBrightGreen, "#00FF00"), _
                Array(wdTurquoise, "#00FFFF"), _
                Array(wdPink, "#FF00FF"), _
                Array(wdBlue, "#0000FF"), _
                Array(wdRed, "#FF0000"), _
                Array(wdDarkBlue, "#000080"), _
                Array(wdTeal, "#008080"), _
                Array(wdGreen, "#008000"), _
                Array(wdViolet, "#800080"), _
                Array(wdDarkRed, "#800000"), _
                Array(wdDarkYellow, "#808000"), _
                Array(wdGray50, "#808080"), _
                Array(wdGray25, "#C0C0C0"), _
                Array(wdBlack, "#000000") _
            )

            ' Zuerst Farben verarbeiten
            Dim cs As Variant
            For Each cs In colorSettings
                With rngFontColor.Find
                    .ClearFormatting
                    .Font.Color = cs(0)
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = True
                    .Text = ""
                    .Replacement.Text = ""
                    .Execute
                    Do While .Found
                        rngFontColor.InsertBefore "<span style=""color: " & cs(1) & ";"">"
                        rngFontColor.InsertAfter "</span>"
                        .Execute
                    Loop
                End With
            Next cs

            ' Dann Hervorhebungen verarbeiten
            With rngMarkColor.Find
                .ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .Replacement.ClearFormatting
                .Text = ""
                .Replacement.Text = ""
                .Highlight = True
                
                .Execute
                Do While .Found
                    Dim highlightColor As String
                    highlightColor = ""
                    
                    ' Finde die passende Farbe
                    Dim hc As Variant
                    For Each hc In highlightColors
                        If rngMarkColor.HighlightColorIndex = hc(0) Then
                            highlightColor = hc(1)
                            Exit For
                        End If
                    Next hc
                    
                    If highlightColor <> "" Then
                        rngMarkColor.InsertBefore "<span style='background-color: " & highlightColor & ";'>"
                        rngMarkColor.InsertAfter "</span>"
                    End If
                    
                    .Execute
                Loop
            End With


            ' Dann Listen verarbeiten
            Dim para As Paragraph
            Dim htmlContent As String
            Dim currentLevel As Integer
            Dim lastLevel As Integer
            Dim listLevels(9) As Integer  ' Array zum Speichern der Listentypen (0=none, 1=bullet, 2=number, 3=multilevel)
            htmlContent = ""
            
            currentLevel = 0
            lastLevel = 0
            
            For Each para In doc.Paragraphs
                ' Entferne nur die Absatzmarke am Ende
                Dim paraText As String
                paraText = para.Range.Text
                If Len(paraText) > 0 Then
                    If Right(paraText, 1) = vbCr Then
                        paraText = Left(paraText, Len(paraText) - 1)
                    End If
                End If
                
                ' Prüfe ob der Absatz Teil einer Liste ist
                If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                    currentLevel = para.Range.ListFormat.ListLevelNumber
                    
                    ' Bestimme den Listentyp dieser Ebene
                    Dim listType As Integer
                    Select Case True
                        Case para.Range.ListFormat.ListType = wdListBullet
                            listType = 1
                        Case para.Range.ListFormat.ListType = wdListSimpleNumbering
                            listType = 2
                        Case para.Range.ListFormat.ListType = wdListMultiLevel
                            ' Prüfe den tatsächlichen Typ dieser Ebene
                            If para.Range.ListFormat.ListString Like "*•*" Then
                                listType = 1
                            Else
                                listType = 2
                            End If
                        Case Else
                            listType = 0
                    End Select
                    
                    ' Behandle Änderungen in der Verschachtelungsebene
                    If currentLevel > lastLevel Then
                        ' Neue tiefere Ebene
                        listLevels(currentLevel) = listType
                        If listType = 1 Then
                            htmlContent = htmlContent & "<ul style='margin: 0; padding-left: 20px;'>"
                        Else
                            htmlContent = htmlContent & "<ol style='margin: 0; padding-left: 20px;'>"
                        End If
                    ElseIf currentLevel < lastLevel Then
                        ' Zurück zu höherer Ebene - schließe Zwischenebenen
                        For i = lastLevel To currentLevel + 1 Step -1
                            If listLevels(i) = 1 Then
                                htmlContent = htmlContent & "</ul>"
                            Else
                                htmlContent = htmlContent & "</ol>"
                            End If
                        Next i
                    ElseIf currentLevel > 0 And listLevels(currentLevel) <> listType Then
                        ' Gleiche Ebene aber anderer Listentyp
                        If listLevels(currentLevel) = 1 Then
                            htmlContent = htmlContent & "</ul><ol style='margin: 0; padding-left: 20px;'>"
                        Else
                            htmlContent = htmlContent & "</ol><ul style='margin: 0; padding-left: 20px;'>"
                        End If
                        listLevels(currentLevel) = listType
                    End If
                    
                    ' Listenelement hinzufügen
                    htmlContent = htmlContent & "<li style='margin-bottom: 6px;'>" & paraText & "</li>"
                    
                Else
                    ' Kein Listenelement - schließe alle offenen Listen
                    Dim j As Integer
                    For j = lastLevel To 1 Step -1
                        If listLevels(j) = 1 Then
                            htmlContent = htmlContent & "</ul>"
                        Else
                            htmlContent = htmlContent & "</ol>"
                        End If
                        listLevels(j) = 0
                    Next j
                    
                    ' Extra Zeilenumbruch nach Liste einfügen wenn vorher eine Liste war
                    If lastLevel > 0 Then
                        htmlContent = htmlContent & "<br>"
                    End If
                    
                    ' Normaler Text
                    If para.Next Is Nothing Then
                        htmlContent = htmlContent & paraText
                    ElseIf para.Range.Text = vbCr Then
                        htmlContent = htmlContent & "<br>"
                    Else
                        htmlContent = htmlContent & paraText & "<br><br>"
                    End If
                    
                    currentLevel = 0
                End If
                
                lastLevel = currentLevel
            Next para
            
            ' Schließe noch offene Listen
            Dim k As Integer
            For k = lastLevel To 1 Step -1
                If listLevels(k) = 1 Then
                    htmlContent = htmlContent & "</ul>"
                Else
                    htmlContent = htmlContent & "</ol>"
                End If
            Next k
            
            ' Ersetze den Dokumentinhalt mit dem HTML-formatierten Text
            doc.Content.Text = htmlContent

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
                
                ' Platzhalter ersetzen (nur wenn Spalte definiert)
                strBody = doc.Content.Text
                If SpalteAnrede <> "" Then strBody = Replace(strBody, "%Anrede%", strAnrede) ' <--- Hier können Sie die Platzhalter anpassen
                If SpalteTitel <> "" Then strBody = Replace(strBody, "%Titel%", xlWS.Range(SpalteTitel & i).Value)
                If SpalteVorname <> "" Then strBody = Replace(strBody, "%Vorname%", strVorname)
                If SpalteNachname <> "" Then strBody = Replace(strBody, "%Nachname%", strNachname)
                If SpalteUnternehmen <> "" Then strBody = Replace(strBody, "%Unternehmen%", xlWS.Range(SpalteUnternehmen & i).Value)
                
                ' Standard-Schriftart und -Schriftgröße aus dem gesamten Dokument ermitteln
                Dim fontName As String
                Dim fontSize As Single
                Dim isUniqueFont As Boolean

                ' Starte mit dem Haupttext
                fontName = ActiveDocument.Content.Font.Name
                fontSize = ActiveDocument.Content.Font.Size
                isUniqueFont = True

                ' Durchlaufe alle StoryRanges (Haupttext, Kopf-/Fußzeilen, etc.)
                Dim sr As Range
                For Each sr In ActiveDocument.StoryRanges
                    If sr.Font.Name <> fontName Then
                        isUniqueFont = False
                        Exit For
                    End If
                Next sr

                ' Falls unterschiedliche Schriftarten gefunden werden, verwende die Standardschriftart und -größe aus dem Normal-Format
                If Not isUniqueFont Then
                    fontName = ActiveDocument.Styles(wdStyleNormal).Font.Name
                    fontSize = ActiveDocument.Styles(wdStyleNormal).Font.Size
                End If

                Set objMail = objOutlook.CreateItem(0)

                ' Anschließend in den HTML-Code einbetten (Schriftgröße in pt)
                Dim htmlTemplate As String
                htmlTemplate = "<html>" & _
                               "<head>" & _
                               "<meta charset=""UTF-8"">" & _
                               "<style type=""text/css"">" & _
                               "body { font-family: " & fontName & "; font-size: " & fontSize & "pt; }" & _
                               "</style>" & _
                               "</head>" & _
                               "<body>" & strBody & "</body>" & _
                               "</html>"

                objMail.HTMLBody = htmlTemplate
                
                '******************************************************************************
                ' ** 11. E-Mail-Versand (MODIFIKATIONSMÖGLICHKEIT: Vorgang pausieren) **
                '******************************************************************************
                On Error Resume Next
                With objMail
                    .To = strTo
                    .CC = strCC
                    .BCC = strBCC
                    .Subject = strSubj
                    .HTMLBody = htmlTemplate
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
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
        Else
            MsgBox "Keine Datei ausgewählt"
        End If
    End With
End Sub
