Sub SendEmailsFromWordWithExcelWithAbfrage()
Dim objOutlook As Object
Dim objMail As Object
Dim doc As Document
Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim xlWS As Excel.Worksheet
Dim Pfad As Variant
Dim fd As Office.FileDialog
Dim objUndo As UndoRecord
Set objUndo = Application.UndoRecord

' Wähle die Excel-Datei aus. Der Dateipfad wird dann für den weiteren Cod zwischengespeichert.
Set xlApp = New Excel.Application
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .Title = "Excel-Liste auswählen"
    .Filters.Clear
    .Filters.Add "Excel-Dateien", "*.xl*"
    .AllowMultiSelect = False
    .ButtonName = "Auswählen"
    If .Show = -1 Then
        Pfad = .SelectedItems(1)
        Set xlWB = xlApp.Workbooks.Open(Pfad)
        Set xlWS = xlWB.Sheets(1)
        ' Hier kommt der Code, der ausgeführt wird, wenn der Benutzer eine Datei ausgewählt hat.
        ' Outlook wird geöffnet
        Set objOutlook = CreateObject("Outlook.Application")

        ' Das aktuell geöffnete und aktive Word-Dokument wird herangezogen
        Set doc = ActiveDocument
        
        
        'Nun folgt die Abfrage, in welcher Spalte die einzelnen Daten stehen
        Dim AnredeRange As Excel.Range
        Set AnredeRange = xlWS.Range("A1:Z10").Find("Anrede", LookIn:=xlValues, LookAt:=xlWhole)

        If Not AnredeRange Is Nothing Then
        'Die Anrede-Spalte ist in der gefundenen Zelle
        SpalteAnrede = Chr(AnredeRange.Cells.Column + 64)
        Else
        'Die Anrede-Spalte wurde nicht gefunden
        SpalteAnrede = InputBox("Geben Sie den Spaltenbuchstaben für die Anrede ein (z. B. A):")
            If SpalteAnrede = "" Then
            If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub
            End If
            End If
        End If

        Dim TitelRange As Excel.Range
        Set TitelRange = xlWS.Cells.Find("Titel", LookIn:=xlValues, LookAt:=xlWhole)

        If Not TitelRange Is Nothing Then
        SpalteTitel = Chr(TitelRange.Cells.Column + 64)
        Else
        SpalteTitel = InputBox("Geben Sie den Spaltenbuchstaben für den Titel ein (z. B. B):")
            If SpalteTitel = "" Then
            If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub
            End If
            End If
        End If

        Dim VornameRange As Excel.Range
        Set VornameRange = xlWS.Cells.Find("Vorname", LookIn:=xlValues, LookAt:=xlWhole)

        If Not VornameRange Is Nothing Then
        SpalteVorname = Chr(VornameRange.Cells.Column + 64)
        Else
        SpalteVorname = InputBox("Geben Sie den Spaltenbuchstaben für den Vornamen ein (z. B. C):")
            If SpalteVorname = "" Then
            If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub
            End If
            End If
        End If

        Dim NachnameRange As Excel.Range
        Set NachnameRange = xlWS.Cells.Find("Nachname", LookIn:=xlValues, LookAt:=xlWhole)

        If Not NachnameRange Is Nothing Then
        SpalteNachname = Chr(NachnameRange.Cells.Column + 64)
        Else
        SpalteNachname = InputBox("Geben Sie den Spaltenbuchstaben für den Nachnamen ein:")
            If SpalteNachname = "" Then
            If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub
            End If
            End If
        End If

        
        ' Suchbegriffe definieren
        Suchbegriffe = Array("E-Mail", "email", "e-Mail", "e-mail")

        ' Nach jedem Suchbegriff suchen
        For s = LBound(Suchbegriffe) To UBound(Suchbegriffe)
            Dim ToRange As Excel.Range
            Set ToRange = xlWS.Cells.Find(Suchbegriffe(s), LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not ToRange Is Nothing Then
                SpalteTo = Chr(ToRange.Cells.Column + 64)
                Exit For
            End If
        Next s

        ' Wenn keine Übereinstimmung gefunden wurde, benutze InputBox
        If ToRange Is Nothing Then
            SpalteTo = InputBox("Geben Sie den Spaltenbuchstaben für die Empfängeradresse ein:")
            If SpalteTo = "" Then
                If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
                    ' Aufräumen
                    Set objMail = Nothing
                    Set objOutlook = Nothing
                    ' Close the Excel file
                    xlWB.Close SaveChanges:=False
                    xlApp.Quit
                    Set xlWS = Nothing
                    Set xlWB = Nothing
                    Set xlApp = Nothing
                    Set Pfad = Nothing
                    Exit Sub
                End If
            End If
        End If

        
        Dim BetreffRange As Excel.Range
        Set BetreffRange = xlWS.Cells.Find("Betreff", LookIn:=xlValues, LookAt:=xlWhole)

        If Not BetreffRange Is Nothing Then
        SpalteSubj = Chr(BetreffRange.Cells.Column + 64)
        Else
        SpalteSubj = InputBox("Geben Sie den Spaltenbuchstaben für den Betreff ein:")
            If SpalteSubj = "" Then
            If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub
            End If
            End If
        End If

        Dim Suchbegriffe2 As Variant
        Dim t As Integer
        ' Suchbegriffe definieren
        Suchbegriffe2 = Array("Anhang", "Anhänge")
        
        Dim AnhangGefunden As Boolean
        AnhangGefunden = False

        ' Nach jedem Suchbegriff suchen
        For t = LBound(Suchbegriffe2) To UBound(Suchbegriffe2)
            Dim AnhangRange As Excel.Range
            Set AnhangRange = xlWS.Cells.Find(Suchbegriffe2(t), LookIn:=xlValues, LookAt:=xlWhole)

            If Not AnhangRange Is Nothing Then
                SpalteAttach = Chr(AnhangRange.Cells.Column + 64)
                AnhangGefunden = True
                Exit For
            End If
        Next t

        If Not AnhangGefunden Then
            SpalteAttach = InputBox("Geben Sie den Spaltenbuchstaben für den Anhang ein:")
            If SpalteAttach = "" Then
                If MsgBox("Möchten Sie den Vorgang abbrechen?", vbYesNo) = vbYes Then
                    ' Aufräumen
                    Set objMail = Nothing
                    Set objOutlook = Nothing
                    ' Close the Excel file
                    xlWB.Close SaveChanges:=False
                    xlApp.Quit
                    Set xlWS = Nothing
                    Set xlWB = Nothing
                    Set xlApp = Nothing
                    Set Pfad = Nothing
                    Exit Sub
                End If
            End If
        End If

        
        Dim Einverstanden As VbMsgBoxResult
        Einverstanden = vbNo
        Do While Einverstanden = vbNo ' Schleife, um die Eingabe zu wiederholen, falls der Benutzer nicht einverstanden ist
        ' Zeige die MsgBox mit den Datengruppen und Spaltenzuordnungen an
        Dim msg As String
        msg = "Folgende Datengruppen wurden gefunden:" & vbNewLine & _
          "Anrede: " & SpalteAnrede & vbNewLine & _
          "Titel: " & SpalteTitel & vbNewLine & _
          "Vorname: " & SpalteVorname & vbNewLine & _
          "Nachname: " & SpalteNachname & vbNewLine & _
          "E-Mail: " & SpalteTo & vbNewLine & _
          "Betreff: " & SpalteSubj & vbNewLine & _
          "Anhang: " & SpalteAttach & vbNewLine & vbNewLine & _
          "Sind Sie mit den ausgewählten Spalten einverstanden?"
        Einverstanden = MsgBox(msg, vbQuestion + vbYesNoCancel, "Datengruppen und Spalten")
        ' Falls der Benutzer nicht einverstanden ist, frage nach der Spalte, die korrigiert werden soll
        If Einverstanden = vbNo Then
        Dim Datengruppe As String
        Datengruppe = InputBox("Bitte wählen Sie die Datengruppe aus, die nicht korrekt zugeordnet ist:" & vbNewLine & _
                              "Anrede, Titel, Vorname, Nachname, E-Mail, Betreff, Anhang")

        ' Weise der ausgewählten Datengruppe die korrekte Spalte zu
        Select Case Datengruppe
            Case "Anrede"
                SpalteAnrede = InputBox("Bitte geben Sie die korrekte Spalte für Anrede ein:")
            Case "Titel"
                SpalteTitel = InputBox("Bitte geben Sie die korrekte Spalte für Titel ein:")
            Case "Vorname"
                SpalteVorname = InputBox("Bitte geben Sie die korrekte Spalte für Vorname ein:")
            Case "Nachname"
                SpalteNachname = InputBox("Bitte geben Sie die korrekte Spalte für Nachname ein:")
            Case "E-Mail"
                SpalteTo = InputBox("Bitte geben Sie die korrekte Spalte für Empfängeradresse ein:")
            Case "Betreff"
                SpalteSubj = InputBox("Bitte geben Sie die korrekte Spalte für für Betreff ein:")
            Case "Anhang"
                SpalteAttach = InputBox("Bitte geben sie die korrekte Spalte für den Anhang ein:")
            End Select
        
        ElseIf Einverstanden = vbCancel Then
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
        Exit Sub 'Beende den Sub, wenn der Benutzer "Abbrechen" auswählt
        
        Else 'Benutzer ist einverstanden
            Einverstanden = False
        End If
        Loop


        ' Starte die Aktions-Aufzeichnung
        objUndo.StartCustomRecord ("VBA-Aktionen")
        
        ' Da der Text in HTML-Text umgewandelt werden muss, werden nun z. B. fettgedruckte Zeichen mit den jeweils
        ' passenden HTML-Tags umgeben, damit dies in der E-Mail korrekt dargestellt wird.
        Dim bRange As Range, iRange As Range, uRange As Range, supRange As Range, subRange As Range, smallRange As Range, bigRange As Range
        Set bRange = doc.Content
        Set iRange = doc.Content
        Set uRange = doc.Content
        Set supRange = doc.Content
        Set subRange = doc.Content
        Set smallRange = doc.Content
        Set bigRange = doc.Content

        ' Umschließe Hyperlinks mit den HTML-Tags
        Dim HL As hyperlink
        For Each HL In ActiveDocument.Hyperlinks
        HL.Range.Text = "<a href=""" & HL.Address & """>" & HL.Range.Text & "</a>"
        Next

        With bRange.Find
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With iRange.Find
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With uRange.Find
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With supRange.Find
        .ClearFormatting
        .Font.Superscript = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With subRange.Find
        .ClearFormatting
        .Font.Subscript = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With smallRange.Find
        .ClearFormatting
        .Font.SmallCaps = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        With bigRange.Find
        .ClearFormatting
        .Font.AllCaps = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        End With

        Do While bRange.Find.Execute() Or iRange.Find.Execute() Or uRange.Find.Execute() Or supRange.Find.Execute() Or subRange.Find.Execute() Or smallRange.Find.Execute() Or bigRange.Find.Execute() 'Loop through all matches
        If bRange.Find.Found Then
            bRange.InsertBefore "<b>" 'Insert <b> before bold word
            bRange.InsertAfter "</b>" 'Insert </b> after bold word
            bRange.Collapse wdCollapseEnd 'Move to next match
        End If
       
        If iRange.Find.Found Then
            iRange.InsertBefore "<i>" 'Insert <i> before italicized word
            iRange.InsertAfter "</i>" 'Insert </i> after italicized word
            iRange.Collapse wdCollapseEnd 'Move to next match
        End If
       
        If uRange.Find.Found Then
            uRange.InsertBefore "<u>" 'Insert <u> before underlined word
            uRange.InsertAfter "</u>" 'Insert </u> after underlined word
            uRange.Collapse wdCollapseEnd
        End If
       
        If supRange.Find.Found Then
            supRange.InsertBefore "<sup>" 'Insert <sup> before superscript word
            supRange.InsertAfter "</sup>" 'Insert </sup> after superscript word
            supRange.Collapse wdCollapseEnd
        End If
       
        If subRange.Find.Found Then
            subRange.InsertBefore "<sub>" 'Insert <sub> before subscript word
            subRange.InsertAfter "</sub>" 'Insert </sub> after subscript word
            subRange.Collapse wdCollapseEnd
        End If
       
        If smallRange.Find.Found Then
            smallRange.InsertBefore "<small>" 'Insert <small> before small caps word
            smallRange.InsertAfter "</small>" 'Insert </small> after small caps word
            smallRange.Collapse wdCollapseEnd
        End If
       
        If bigRange.Find.Found Then
            bigRange.InsertBefore "<big>" 'Insert <big> before all caps word
            bigRange.InsertAfter "</big>" 'Insert </big> after all caps word
            bigRange.Collapse wdCollapseEnd
        End If
        Loop
           
        ' Ersetze Zeilenumbrüche mit dem passenden HTML-Tag
        doc.Content = Replace(doc.Content, vbCr, "<br><br>")
        doc.Content = Replace(doc.Content, vbLf, "<br>")

        Dim rngTo As Excel.Range, rngSubj As Excel.Range, rngAttach As Excel.Range
        Dim strTo As String, strSubj As String, strAttach As String, strBody As String
        Dim strAnrede As String, strVorname As String, strNachname As String, DateipfadCheck As String
        Dim lastRow As Long
        
        ' Nun folgt der Teil, in dem der erstellte E-Mail-Text mit den Daten aus der Excel-Tabelle ersetzt wird, bis keine ausgefüllt Zeile mehr vorhanden ist.
        lastRow = xlWS.Cells(Rows.Count, 1).End(xlUp).Row
        
        For d = 2 To lastRow
        strVorname = xlWS.Range(SpalteVorname & d).Value ' Wert in der Spalte mit dem Vornamen, aktuelle Zeile
        strNachname = xlWS.Range(SpalteNachname & d).Value ' Wert in der Spalte mit dem Nachnamen, aktuelle Zeile
        DateipfadCheck = xlWS.Range(SpalteAttach & d).Value ' Wert in der Spalte mit den Anhängen, aktuelle Zeile
        arrFileNames = Split(DateipfadCheck, ",") 'Trenne den Zellinhalt anhand eines Kommas
        
        ' Überprüfe, ob die Dokumentenpfade gültig sind
        Dim a As Long
        Dim fehlerListe As String
        For a = LBound(arrFileNames) To UBound(arrFileNames)
        If Dir(arrFileNames(a)) = "" Then
        fehlerMeldungAnhang = "Der Pfad zu " & arrFileNames(a) & " ist ungültig."
        fehlerListe = fehlerListe & "Datensatz " & strNachname & ", " & strVorname & ": " & fehlerMeldungAnhang & vbCrLf & "Bitte korrigieren Sie den Pfad bzw. Pfade und starten den Vorgang erneut. Es wurde keine E-Mail versendet."
        End If
        Next a
        Next d
        If fehlerListe <> "" Then
            MsgBox "Folgende Fehler wurden gefunden:" & vbCrLf & fehlerListe
            ActiveDocument.Undo
            ' Aufräumen
            Set objMail = Nothing
            Set objOutlook = Nothing
            ' Close the Excel file
            xlWB.Close SaveChanges:=False
            xlApp.Quit
            Set xlWS = Nothing
            Set xlWB = Nothing
            Set xlApp = Nothing
            Set Pfad = Nothing
            Exit Sub ' beende den Sub, falls Fehler gefunden wurden
        End If
        ' Wiederhole die Datenersetzung für jeden Datensatz
        For i = 2 To lastRow ' Schleife von Zeile 2 bis zur letzten Zeile
        strTo = xlWS.Range(SpalteTo & i).Value ' Wert in der Spalte mit den Empfängeradressen, aktuelle Zeile
        strSubj = xlWS.Range(SpalteSubj & i).Value ' Wert in der Spalte mit den Betreffzeilen, aktuelle Zeile
        strAnrede = xlWS.Range(SpalteAnrede & i).Value ' Wert in der Spalte mit der Anrede, aktuelle Zeile
        strTitel = xlWS.Range(SpalteTitel & i).Value ' Wert in der Spalte mit dem Titel, aktuelle Zeile
        strVorname = xlWS.Range(SpalteVorname & i).Value ' Wert in der Spalte mit dem Vornamen, aktuelle Zeile
        strNachname = xlWS.Range(SpalteNachname & i).Value ' Wert in der Spalte mit dem Nachnamen, aktuelle Zeile
        strAttach = xlWS.Range(SpalteAttach & i).Value ' Wert in der Spalte mit den Anhängen, aktuelle Zeile
        
       
        ' Ersetze die Platzhalter mit den Daten aus der Excel-Liste
       
        strBody = Replace(doc.Content.Text, "%Anrede%", strAnrede)
        strBody = Replace(strBody, "%Titel%", strTitel)
        strBody = Replace(strBody, "%Vorname%", strVorname)
        strBody = Replace(strBody, "%Nachname%", strNachname)
       
     
        ' Öffne eine neue E-Mail
        On Error Resume Next
        Set objMail = objOutlook.CreateItem(0)
       
        ' Stelle die E-Mail Einstellungen richtig ein
        With objMail
            .To = strTo
            .Subject = strSubj
            .BodyFormat = 2 'olFormatHTML
            .HTMLBody = strBody
             If strAttach <> "" Then
            .Attachments.Add (strAttach)
             End If
            .Send ' Send the email
        End With


        Dim fehlerMeldung As String
        Dim sentCount As Integer
        ' Überprüfen, ob der Sendevorgang erfolgreich war oder ob es einen Fehler gab
        If Err.Number <> 0 Then
        ' Fehlerbehandlungsroutine
        fehlerMeldung = fehlerMeldung & vbCrLf & "Fehler beim Senden an " & strVorname & " " & strNachname & ": " & Err.Description
        Else
        sentCount = sentCount + 1 ' erhöhe die Anzahl der gesendeten E-Mails
        End If

        Set objMail = Nothing
        Err.Clear
        On Error GoTo 0
        Next i
        
        ' Entferne alle HTML-Tags aus dem Dokument
        objUndo.EndCustomRecord
        ActiveDocument.Undo

        ' Wenn Fehler aufgetreten sind, zeige eine Fehlermeldung
        If fehlerMeldung <> "" Then
        MsgBox "Folgende Datensätze konnten nicht gesendet werden:" & vbCrLf & fehlerMeldung
        Else ' Wenn keine Fehler aufgetreten sind, zeige eine Meldung, dass alle E-Mails erfolgreich gesendet wurden
        MsgBox "Alle E-Mails wurden erfolgreich gesendet (" & sentCount & " gesendet)."
        End If

        ' Aufräumen
        Set objMail = Nothing
        Set objOutlook = Nothing
        ' Close the Excel file
        xlWB.Close SaveChanges:=False
        xlApp.Quit
        Set xlWS = Nothing
        Set xlWB = Nothing
        Set xlApp = Nothing
        Set Pfad = Nothing
   
    Else
        MsgBox "Der Sendungsvorgang wurde nicht gestartet, da keine Datei ausgewählt wurde."
    End If
End With

End Sub