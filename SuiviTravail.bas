Attribute VB_Name = "SuiviTravail"
Option Explicit

' Version finale : bouton d'alerte, dashboard et envoi Outlook.
' Enregistrez le classeur en .xlsm avant d'exécuter ces macros.

Public Sub InstallerClasseur()
    Dim ws As Worksheet
    Dim dash As Worksheet
    Dim shp As Shape

    Set ws = GetOrCreateSheet("Suivi")
    Set dash = GetOrCreateSheet("Dashboard")

    Application.ScreenUpdating = False

    If Trim$(ws.Cells(1, 1).Value) = "" Then
        ws.Range("A1:M1").Value = Array("Projet", "Usine", "N° article", "Statut", "Avancement (%)", _
            "Date de début", "Date de fin", "Jours restants", "Date d'envoi email", _
            "Email envoyé", "Alerte", "Remarques", "Email destinataire")
    End If

    FormatSuivi ws
    ConstruireDashboard dash
    ActualiserAlertes

    On Error Resume Next
    dash.Shapes("btnAlertes").Delete
    On Error GoTo 0
    Set shp = dash.Shapes.AddShape(msoShapeRoundedRectangle, 260, 55, 180, 35)
    With shp
        .Name = "btnAlertes"
        .TextFrame2.TextRange.Characters.Text = "Actualiser les alertes"
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .OnAction = "ActualiserAlertes"
    End With

    Application.ScreenUpdating = True
    MsgBox "Installation terminée. Le bouton est disponible dans Dashboard.", vbInformation
End Sub

Public Sub ActualiserAlertes()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim status As String, dueDate As Variant, sent As String, progress As Double
    Set ws = GetOrCreateSheet("Suivi")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        status = LCase$(Trim$(CStr(ws.Cells(r, 4).Value)))
        dueDate = ws.Cells(r, 7).Value
        sent = LCase$(Trim$(CStr(ws.Cells(r, 10).Value)))
        progress = Val(ws.Cells(r, 5).Value)
        If progress > 1 Then progress = progress / 100

        ws.Cells(r, 8).FormulaLocal = "=SI(D" & r & "=""Terminé"";0;G" & r & "-AUJOURDHUI())"
        If IsDate(dueDate) And status <> "terminé" Then
            If CDate(dueDate) < Date Then
                ws.Cells(r, 11).Value = "RETARD"
            ElseIf progress < 0.5 And CDate(dueDate) - Date <= 3 Then
                ws.Cells(r, 11).Value = "URGENT"
            ElseIf IsDate(ws.Cells(r, 9).Value) And ws.Cells(r, 9).Value <= Date And sent <> "oui" Then
                ws.Cells(r, 11).Value = "EMAIL"
            Else
                ws.Cells(r, 11).ClearContents
            End If
        Else
            ws.Cells(r, 11).ClearContents
        End If
    Next r

    ColorerLignes ws, lastRow
    MettreAJourDashboard
End Sub

Public Sub EnvoyerAlertesOutlook()
    Dim ws As Worksheet, lastRow As Long, r As Long, countSent As Long
    Dim outlook As Object, mail As Object
    Set ws = GetOrCreateSheet("Suivi")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set outlook = CreateObject("Outlook.Application")

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 10).Value))) <> "OUI" _
           And Len(Trim$(CStr(ws.Cells(r, 13).Value))) > 0 _
           And (UCase$(CStr(ws.Cells(r, 11).Value)) = "RETARD" Or UCase$(CStr(ws.Cells(r, 11).Value)) = "URGENT" Or UCase$(CStr(ws.Cells(r, 11).Value)) = "EMAIL") Then

            Set mail = outlook.CreateItem(0)
            With mail
                .To = ws.Cells(r, 13).Value
                .Subject = "Alerte suivi - " & ws.Cells(r, 1).Value & " - " & ws.Cells(r, 3).Value
                .Body = "Bonjour," & vbCrLf & vbCrLf & _
                    "Projet : " & ws.Cells(r, 1).Value & vbCrLf & _
                    "Usine : " & ws.Cells(r, 2).Value & vbCrLf & _
                    "Article : " & ws.Cells(r, 3).Value & vbCrLf & _
                    "Statut : " & ws.Cells(r, 4).Value & vbCrLf & _
                    "Avancement : " & Format(ws.Cells(r, 5).Value, "0%") & vbCrLf & _
                    "Date de fin : " & ws.Cells(r, 7).Text & vbCrLf & _
                    "Alerte : " & ws.Cells(r, 11).Value & vbCrLf & _
                    "Remarques : " & ws.Cells(r, 12).Value & vbCrLf & vbCrLf & _
                    "Merci de mettre à jour le suivi."
                .Display ' Remplacez .Display par .Send après un test validé.
            End With
            ws.Cells(r, 10).Value = "Oui"
            countSent = countSent + 1
        End If
    Next r

    ActualiserAlertes
    MsgBox countSent & " message(s) préparé(s) dans Outlook.", vbInformation
End Sub

Private Sub ConstruireDashboard(ByVal dash As Worksheet)
    With dash
        .Cells.Clear
        .Range("A1:B1").Merge
        .Range("A1").Value = "TABLEAU DE BORD - SUIVI DE TRAVAIL"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = vbWhite
        .Range("A1:B1").Interior.Color = RGB(31, 78, 121)
        .Range("A3:B3").Value = Array("Indicateur", "Valeur")
        .Range("A4:A10").Value = Application.Transpose(Array("Total travaux", "En cours", "Terminé", "En retard", "Avancement moyen", "Alertes à traiter", "Emails à envoyer"))
        .Range("B4").Formula = "=COUNTA(Suivi!A2:A1000)"
        .Range("B5").Formula = "=COUNTIF(Suivi!D2:D1000,""En cours"")"
        .Range("B6").Formula = "=COUNTIF(Suivi!D2:D1000,""Terminé"")"
        .Range("B7").Formula = "=COUNTIF(Suivi!D2:D1000,""En retard"")"
        .Range("B8").Formula = "=IFERROR(AVERAGE(Suivi!E2:E1000),0)"
        .Range("B9").Formula = "=COUNTIF(Suivi!K2:K1000,""RETARD"")+COUNTIF(Suivi!K2:K1000,""URGENT"")"
        .Range("B10").Formula = "=COUNTIF(Suivi!K2:K1000,""EMAIL"")"
        .Range("A3:B3").Font.Bold = True
        .Range("A3:B3").Interior.Color = RGB(31, 78, 121)
        .Range("A3:B3").Font.Color = vbWhite
        .Range("B4:B10").Font.Bold = True
        .Range("B8").NumberFormat = "0%"
        .Columns("A:B").AutoFit
    End With
End Sub

Private Sub MettreAJourDashboard()
    Dim dash As Worksheet
    Set dash = GetOrCreateSheet("Dashboard")
    dash.Range("B4").Formula = "=COUNTA(Suivi!A2:A1000)"
    dash.Range("B5").Formula = "=COUNTIF(Suivi!D2:D1000,""En cours"")"
    dash.Range("B6").Formula = "=COUNTIF(Suivi!D2:D1000,""Terminé"")"
    dash.Range("B7").Formula = "=COUNTIF(Suivi!D2:D1000,""En retard"")"
    dash.Range("B8").Formula = "=IFERROR(AVERAGE(Suivi!E2:E1000),0)"
    dash.Range("B9").Formula = "=COUNTIF(Suivi!K2:K1000,""RETARD"")+COUNTIF(Suivi!K2:K1000,""URGENT"")"
    dash.Range("B10").Formula = "=COUNTIF(Suivi!K2:K1000,""EMAIL"")"
End Sub

Private Sub FormatSuivi(ByVal ws As Worksheet)
    With ws.Range("A1:M1")
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("E2:E1000").NumberFormat = "0%"
    ws.Range("F2:G1000,I2:I1000").NumberFormat = "dd/mm/yyyy"
    ws.Range("A1:M1000").Borders.LineStyle = xlContinuous
    ws.Range("A1:M1").AutoFilter
    ws.Columns("A:M").AutoFit
    ws.Columns("L").ColumnWidth = 35
    ws.Columns("M").ColumnWidth = 28
End Sub

Private Sub ColorerLignes(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long, value As String
    For r = 2 To lastRow
        value = UCase$(Trim$(CStr(ws.Cells(r, 11).Value)))
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Pattern = xlNone
        Select Case value
            Case "RETARD": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(255, 199, 206)
            Case "URGENT": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(255, 235, 156)
            Case "EMAIL": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(255, 255, 204)
            Case Else
                Select Case LCase$(Trim$(CStr(ws.Cells(r, 4).Value)))
                    Case "terminé": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(198, 239, 206)
                    Case "en cours": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(221, 235, 247)
                    Case "en attente": ws.Range(ws.Cells(r, 1), ws.Cells(r, 13)).Interior.Color = RGB(231, 230, 230)
                End Select
        End Select
    Next r
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function
