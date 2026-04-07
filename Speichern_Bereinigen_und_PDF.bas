Option Explicit

Public Sub Speichern_Bereinigen_und_PDF()
    Dim wbSrc As Workbook, wbNew As Workbook
    Dim pfadZiel As String, basisName As String, neuerName As String
    Dim zielXlsx As String, zielPdf As String, sep As String
    Dim nmZelle As Variant
    Dim pfadAusZelle As String
    Dim ws As Worksheet

    ' Performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo EndeMitFehler

    Set wbSrc = ThisWorkbook

    '-------------------------------
    ' 1) Zielpfad ermitteln
    '    - Standard: Speicherort der Datei
    '    - Override: Pfad aus "Steuerung & ppt"!B17 (falls gefüllt)
    '-------------------------------

    ' Standardpfad = Speicherort der aktuellen Datei
    pfadZiel = wbSrc.Path

    ' Prüfen, ob das Blatt existiert (brauchen wir sowieso für B16)
    If Not BlattExistiert(wbSrc, "Steuerung & ppt") Then
        MsgBox "Blatt 'Steuerung & ppt' nicht gefunden.", vbCritical
        GoTo SauberRaus
    End If

    ' Pfad aus B17 lesen
    pfadAusZelle = Trim$(CStr(wbSrc.Worksheets("Steuerung & ppt").Range("B17").Value))
    If Len(pfadAusZelle) > 0 Then
        pfadZiel = pfadAusZelle
    End If

    ' Wenn immer noch kein Pfad vorhanden ist -> Abbruch
    If Len(Trim$(pfadZiel)) = 0 Then
        MsgBox "Kein Zielpfad gefunden." & vbCrLf & _
               "Bitte entweder:" & vbCrLf & _
               "1) die Arbeitsmappe speichern oder" & vbCrLf & _
               "2) in 'Steuerung & ppt'!B17 einen gültigen Ordnerpfad eintragen.", _
               vbExclamation
        GoTo SauberRaus
    End If

    ' Optional: Existenz des Ordners prüfen (nicht für http/SharePoint)
    If LCase$(Left$(pfadZiel, 4)) <> "http" Then
        If Right$(pfadZiel, 1) = "\" Or Right$(pfadZiel, 1) = "/" Then
            ' ok, schon mit Separator am Ende
        Else
            ' Ordner existiert?
            If Dir(pfadZiel, vbDirectory) = "" Then
                MsgBox "Der in 'Steuerung & ppt'!B17 angegebene Ordner existiert nicht:" & vbCrLf & _
                       pfadZiel, vbCritical
                GoTo SauberRaus
            End If
        End If
    End If

    '-------------------------------
    ' 2) Dateinamen aus B16 holen
    '-------------------------------
    nmZelle = wbSrc.Worksheets("Steuerung & ppt").Range("B16").Value
    basisName = OhneUngueltigeZeichen(CStr(nmZelle))
    If Len(Trim$(basisName)) = 0 Then
        basisName = LinkeOhneEndung(wbSrc.Name) & "_Export_" & Format(Now, "yyyymmdd_HHmm")
    End If
    neuerName = basisName

    '-------------------------------
    ' 3) Pfad + Dateinamen zusammenbauen
    '    - http/SharePoint -> "/"
    '    - sonst -> "\"
    '    - trailing slash beachten
    '-------------------------------

    ' Wenn Pfad bereits mit \ oder / endet, keinen zusätzlichen sep anhängen
    If Right$(pfadZiel, 1) = "\" Or Right$(pfadZiel, 1) = "/" Then
        sep = ""
    Else
        sep = IIf(LCase$(Left$(pfadZiel, 4)) = "http", "/", "\")
    End If

    zielXlsx = pfadZiel & sep & neuerName & ".xlsx"
    zielPdf = pfadZiel & sep & neuerName & ".pdf"

    ' 4) Neue leere Arbeitsmappe erzeugen
    Set wbNew = Application.Workbooks.Add(xlWBATWorksheet) ' 1 leeres Blatt

    ' 5) Alle Blätter aus Quellmappe in die neue Mappe kopieren
    For Each ws In wbSrc.Worksheets
        ws.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    Next ws
    ' Das automatisch erstellte leere erste Blatt löschen
    If wbNew.Worksheets.Count > 0 Then wbNew.Worksheets(1).Delete

    ' 6) In der neuen Mappe Werte einsetzen / Blätter löschen
    WerteAbZeileOderTabelle wbNew, "EO-Statistiken", 5
    WerteAbZeileOderTabelle wbNew, "apoFinanzBerater_Übersicht_vert", 5
    WerteNurDaten_Vertrag wbNew, "Vertrag_allgemein_DAEV"  ' nur Daten, Tabelle/Filter erhalten
    WerteAbZeileOderTabelle wbNew, "PDV_Mails", 5          ' PDV_Mails: Formeln -> Werte

    BlattLoeschenFallsDa wbNew, "Steuerung & ppt"
    BlattLoeschenFallsDa wbNew, "AVM_apoFinanzBerater vertraglic"
    BlattLoeschenFallsDa wbNew, "Repräsentanten aktiv"
    BlattLoeschenFallsDa wbNew, "Vertrag_allgemein_DAEV"

    ' 7) AutoRecover aus (UI-"AutoSpeichern" nicht per VBA steuerbar)
    Application.AutoRecover.Enabled = False

    ' 8) Als .xlsx speichern
    wbNew.SaveAs Filename:=zielXlsx, FileFormat:=xlOpenXMLWorkbook

    ' 9) Ausgewählte Blätter als PDF (1 Seite) exportieren
    ExportPdfAusBlaettern wbNew, _
        Array("Übersicht", "EO_Report", "Abgänge", "Zugänge", "Abgänge apoFinanz", "apoFinanzBerater_Übersicht_vert"), _
        zielPdf

    ' 10) Speichern & schließen
    wbNew.Close SaveChanges:=True

    MsgBox "Fertig!" & vbCrLf & "XLSX: " & zielXlsx & vbCrLf & "PDF : " & zielPdf, vbInformation
    GoTo SauberRaus

EndeMitFehler:
    MsgBox "Fehler: " & Err.Number & vbCrLf & Err.Description, vbCritical

SauberRaus:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'==================== Hilfsroutinen ====================

Private Function BlattExistiert(wb As Workbook, ByVal Blattname As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(Blattname)
    BlattExistiert = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Sub BlattLoeschenFallsDa(wb As Workbook, ByVal Blattname As String)
    If BlattExistiert(wb, Blattname) Then wb.Worksheets(Blattname).Delete
End Sub

' Nur den Datenbereich (ab Startzeile ODER Tabellen-DataBodyRange) auf Werte setzen – Header/Filter bleiben intakt
Private Sub WerteAbZeileOderTabelle(wb As Workbook, ByVal Blattname As String, Optional ByVal startZeile As Long = 5)
    If Not BlattExistiert(wb, Blattname) Then Exit Sub
    Dim ws As Worksheet: Set ws = wb.Worksheets(Blattname)

    If ws.ListObjects.Count > 0 Then
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            If Not lo.DataBodyRange Is Nothing Then
                lo.DataBodyRange.Value = lo.DataBodyRange.Value
            End If
        Next lo
    Else
        Dim lastRow As Long
        lastRow = LetzteBenutzteZeile(ws)
        If lastRow >= startZeile Then
            With ws
                .Range(.Rows(startZeile), .Rows(lastRow)).Value = .Range(.Rows(startZeile), .Rows(lastRow)).Value
            End With
        End If
    End If
End Sub

' Für "Vertrag_allgemein_DAEV" – Tabelle behalten, Filter in Zeile 1 sichern
Private Sub WerteNurDaten_Vertrag(wb As Workbook, ByVal Blattname As String)
    If Not BlattExistiert(wb, Blattname) Then Exit Sub
    Dim ws As Worksheet: Set ws = wb.Worksheets(Blattname)

    If ws.ListObjects.Count > 0 Then
        ' Es gibt eine oder mehrere Tabellen -> nur DataBodyRange auf Werte
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            If Not lo.DataBodyRange Is Nothing Then
                lo.DataBodyRange.Value = lo.DataBodyRange.Value
            End If
            ' Filterzeile bleibt bei ListObjects automatisch erhalten
        Next lo
    Else
        ' Keine Tabelle -> ab Zeile 2 bis letzte Zeile Werte setzen, Autofilter in Zeile 1 sicherstellen
        Dim lastRow As Long
        lastRow = LetzteBenutzteZeile(ws)
        If lastRow >= 2 Then
            With ws
                .Range(.Rows(2), .Rows(lastRow)).Value = .Range(.Rows(2), .Rows(lastRow)).Value
            End With
        End If
        ' AutoFilter auf Zeile 1 sicherstellen
        If Not ws.AutoFilterMode Then ws.Rows(1).AutoFilter
    End If
End Sub

Private Sub GanzeTabelleAufWerte(wb As Workbook, ByVal Blattname As String)
    If Not BlattExistiert(wb, Blattname) Then Exit Sub
    Dim ws As Worksheet: Set ws = wb.Worksheets(Blattname)
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub
    ws.UsedRange.Value = ws.UsedRange.Value
End Sub

Private Function LetzteBenutzteZeile(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        LetzteBenutzteZeile = 0
    Else
        LetzteBenutzteZeile = f.Row
    End If
End Function

Private Function OhneUngueltigeZeichen(ByVal s As String) As String
    Dim i As Long, verbot As Variant
    verbot = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    OhneUngueltigeZeichen = Trim$(s)
    For i = LBound(verbot) To UBound(verbot)
        OhneUngueltigeZeichen = Replace(OhneUngueltigeZeichen, CStr(verbot(i)), "_")
    Next i
End Function

Private Function LinkeOhneEndung(ByVal dateiname As String) As String
    Dim p As Long
    p = InStrRev(dateiname, ".")
    If p > 0 Then
        LinkeOhneEndung = Left$(dateiname, p - 1)
    Else
        LinkeOhneEndung = dateiname
    End If
End Function

Private Sub ExportPdfAusBlaettern(wb As Workbook, blattListe As Variant, zielPdf As String)
    Dim existierende As Collection: Set existierende = New Collection
    Dim i As Long, nm As String, ws As Worksheet

    For i = LBound(blattListe) To UBound(blattListe)
        nm = CStr(blattListe(i))
        If BlattExistiert(wb, nm) Then existierende.Add nm
    Next i
    If existierende.Count = 0 Then Exit Sub

    For i = 1 To existierende.Count
        Set ws = wb.Worksheets(existierende(i))
        If ws.AutoFilterMode Or ws.FilterMode Then
            On Error Resume Next
            ws.ShowAllData
            On Error GoTo 0
        End If
        With ws.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    Next i

    Dim arr() As String
    ReDim arr(1 To existierende.Count)
    For i = 1 To existierende.Count
        arr(i) = existierende(i)
    Next i

    wb.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=zielPdf, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    wb.Worksheets(arr(UBound(arr))).Select
End Sub
