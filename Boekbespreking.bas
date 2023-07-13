Attribute VB_Name = "Boekbespreking"

Sub BladenVerwijderen()

    Dim strNaam As String
    Dim i As Long
    Dim lngAantalBladen As Long
    Dim wsStart As Worksheet
    Dim strAntwoord As String
    
    Set wsStart = ActiveWorkbook.Sheets("Start")
    
    strAntwoord = MsgBox(Prompt:="Weet je zeker dat je alles wilt verwijderen?", Buttons:=vbYesNo, Title:="VERWIJDERING TABBLADEN EN GEGEVENS")
    
    If strAntwoord = vbNo Then
        Exit Sub
    Else
        lngAantalBladen = ActiveWorkbook.Sheets.Count
        wsStart.Range("B4:E48").ClearContents
    For i = lngAantalBladen To 1 Step -1
        strNaam = Sheets(i).Name
        If Not (strNaam = "Start" Or strNaam = "Basisblad") Then
            Application.DisplayAlerts = False
                Sheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i
    End If

End Sub

Sub snelkoppelingen()

Application.CommandBars("Workbook tabs").ShowPopup

End Sub


Sub BladenMaken()

Application.ScreenUpdating = False

Dim i As Integer
Dim lastrow As Integer
Dim wsStart As Worksheet
Dim wsBasis As Worksheet
Dim wsNieuw As Worksheet

lastrow = Worksheets("Start").Cells(2, 2).End(xlDown).Row
i = 4

Set wsBasis = ThisWorkbook.Worksheets("Basisblad")
Set wsStart = ThisWorkbook.Worksheets("Start")

wsBasis.Visible = xlSheetVisible

For i = 4 To lastrow
wsBasis.Cells.Copy
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = Worksheets("Start").Cells(i, 2).Value
    End With
Set wsNieuw = ActiveSheet
wsNieuw.Paste
wsNieuw.Cells(3, 4) = "= Start!" & wsStart.Cells(i, 3).Address
wsNieuw.Cells(4, 4) = wsStart.Cells(i, 2).Value
wsNieuw.Cells(5, 4) = "= Start!" & wsStart.Cells(i, 4).Address
wsStart.Cells(i, 5) = "='" & wsStart.Cells(i, 2) & "'!D6"
wsNieuw.Cells(23, 3).Select

ActiveWindow.DisplayGridlines = False

With wsNieuw.PageSetup
 .Zoom = False
 .FitToPagesTall = 1
 .FitToPagesWide = 1
End With

Next i

wsStart.Select
wsBasis.Visible = xlSheetVeryHidden
Application.CutCopyMode = False

End Sub

