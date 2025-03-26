Attribute VB_Name = "Serienmails"
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
                      "Anrede:           " & SpalteAnrede & vbCrLf & _
                      "Titel:                " & SpalteTitel & vbCrLf & _
                      "Vorname:        " & SpalteVorname & vbCrLf & _
                      "Nachname:     " & SpalteNachname & vbCrLf & _
                      "E-Mail:             " & SpalteTo & vbCrLf & _
                      "Betreff:            " & SpalteSubj & vbCrLf & _
                      "Anhang:         " & SpalteAttach & vbCrLf & vbCrLf & _
                      "Diese Spalten wurden zu den Kontaktinformationen gefunden. Sind Sie einverstanden?"
                
                confirmColumns = MsgBox(msg, vbQuestion + vbYesNoCancel, "Spaltenbestätigung")
                
                Select Case confirmColumns
                    Case vbYes
                        correctionLoop = False ' Beenden
                    Case vbCancel
                        GoTo Cleanup
                    Case vbNo
                        ' Korrekturschleife mit explizitem Abbruch über "Abbrechen"-Button
                        Dim columnList As String
                        columnList = "Wählen Sie die Spalte zur Korrektur:" & vbCrLf & _
                                    "1 - Anrede" & vbCrLf & _
                                    "2 - Titel" & vbCrLf & _
                                    "3 - Vorname" & vbCrLf & _
                                    "4 - Nachname" & vbCrLf & _
                                    "5 - E-Mail" & vbCrLf & _
                                    "6 - Betreff" & vbCrLf & _
                                    "7 - Anhang" & vbCrLf & vbCrLf
                        
                        Dim selectedColumn As String
                        selectedColumn = InputBox( _
                            Prompt:=columnList & vbCrLf & "Geben Sie die Nummer der Spalte ein:", _
                            Title:="Spalte korrigieren")
                        
                        If selectedColumn = "" Then ' Abbruch über 'Abbrechen'-Button
                            GoTo Cleanup
                        End If
                        
                        Dim num As Integer
                        Dim neueSpalte As String
                        num = Val(selectedColumn)
                        
                        Select Case num
                            Case 1 To 7 ' Nur gültige Ziffern
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
                                        neueSpalte = InputBox("Neue Spalte für E-Mail:", "E-Mail")
                                        If neueSpalte = "" Then
                                            SpalteTo = ""
                                        Else
                                            SpalteTo = neueSpalte
                                        End If
                                    Case 6
                                        neueSpalte = InputBox("Neue Spalte für Betreff:", "Betreff")
                                        If neueSpalte = "" Then
                                            SpalteSubj = ""
                                        Else
                                            SpalteSubj = neueSpalte
                                        End If
                                    Case 7
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
            objUndo.StartCustomRecord "VBA-Aktionen"
            
            ' Umschreiben von Hyperlinks
            For Each HL In ActiveDocument.Hyperlinks
                HL.Range.Text = "<a href=""" & HL.Address & """>" & HL.Range.Text & "</a>"
            Next
            
            ' HTML-Tags für Formate (Fett, Kursiv, etc.)
            Dim rng As Range
            Set rng = doc.Content
            With rng.Find
                .ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
            End With
            
            ' Alle Formate in einem Loop verarbeiten
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
            
            For Each setting In formatSettings
                With rng.Find
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
            
            ' Zeilenumbrüche ersetzen
            doc.Content.Text = Replace(doc.Content.Text, vbCr, "<br><br>")
            doc.Content.Text = Replace(doc.Content.Text, vbLf, "<br>")
            
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
                Dim strAttach As String
                
                ' Daten aus Excel lesen (mit Fehlertoleranz)
                strTo = IIf(SpalteTo <> "", xlWS.Range(SpalteTo & i).Value, "")
                strSubj = IIf(SpalteSubj <> "", xlWS.Range(SpalteSubj & i).Value, "")
                strAnrede = IIf(SpalteAnrede <> "", xlWS.Range(SpalteAnrede & i).Value, "")
                strVorname = IIf(SpalteVorname <> "", xlWS.Range(SpalteVorname & i).Value, "")
                strNachname = IIf(SpalteNachname <> "", xlWS.Range(SpalteNachname & i).Value, "")
                strAttach = IIf(SpalteAttach <> "", xlWS.Range(SpalteAttach & i).Value, "")
                
                ' Anrede-Behandlung
                If Not useCustomAnrede Then
                    Select Case strAnrede
                        Case "Frau": strAnrede = "Sehr geehrte Frau"
                        Case "Herr": strAnrede = "Sehr geehrter Herr"
                        Case Else: strAnrede = ""
                    End Select
                End If
                
                ' Platzhalter ersetzen (nur wenn Spalte definiert)
                strBody = doc.Content.Text
                If SpalteAnrede <> "" Then strBody = Replace(strBody, "%Anrede%", strAnrede) ' <--- Hier können Sie die Platzhalter anpassen
                If SpalteTitel <> "" Then strBody = Replace(strBody, "%Titel%", xlWS.Range(SpalteTitel & i).Value)
                If SpalteVorname <> "" Then strBody = Replace(strBody, "%Vorname%", strVorname)
                If SpalteNachname <> "" Then strBody = Replace(strBody, "%Nachname%", strNachname)
                
                '******************************************************************************
                ' ** 11. E-Mail-Versand (MODIFIKATIONSMÖGLICHKEIT: Vorgang pausieren) **
                '******************************************************************************
                On Error Resume Next
                Set objMail = objOutlook.CreateItem(0)
                With objMail
                    .To = strTo
                    .Subject = strSubj
                    .HTMLBody = strBody
                    .BodyFormat = 2
                    
                    ' Anhänge hinzufügen
                    If strAttach <> "" Then
                        Dim attachArray() As String
                        attachArray = Split(strAttach, ",")
                        For Each file In attachArray
                            file = Trim(file)
                            If file <> "" Then
                                .Attachments.Add file
                            End If
                        Next
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
            objUndo.EndCustomRecord
            ActiveDocument.Undo
            
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
