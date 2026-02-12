Attribute VB_Name = "Module2"
Option Private Module
Option Explicit


Sub OdbezpieczWszystkieArkusze()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="admin"
    Next ws
    On Error GoTo 0
End Sub

Sub ZabezpieczWszystkieArkusze()
    Dim ws As Worksheet
    Dim h As String
    h = "admin"
    
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        
        Select Case Trim(ws.Name)

            Case "Dashboard"
                ws.Protect Password:=h, _
                           DrawingObjects:=True, _
                           Contents:=True, _
                           Scenarios:=True, _
                           UserInterfaceOnly:=True, _
                           AllowFiltering:=True, _
                           AllowUsingPivotTables:=True, _
                           AllowSorting:=True
                ws.EnableSelection = xlUnlockedCells
                

            Case "Manual_matching"
                ws.Protect Password:=h, _
                           DrawingObjects:=True, _
                           Contents:=True, _
                           Scenarios:=True, _
                           UserInterfaceOnly:=True
                ws.EnableSelection = xlUnlockedCells
                

            Case "Ewidencja_Unmatched", "Ewidencja_Ksiegowa_Full", "Wyciag_Bankowy_Full", "Wyciag_Unmatched", "Matched"
                ws.Protect Password:=h, _
                           DrawingObjects:=True, _
                           Contents:=True, _
                           Scenarios:=True, _
                           UserInterfaceOnly:=True, _
                           AllowFiltering:=True, _
                           AllowSorting:=True
                ws.EnableSelection = xlNoRestrictions
                
            Case Else

                ws.Protect Password:=h, Contents:=True
                ws.EnableSelection = xlNoRestrictions
        End Select
    Next ws
    On Error GoTo 0
End Sub

Sub ZatwierdzIPreniesMatch()
    Dim wsInput As Worksheet, tblArchive As ListObject
    Dim newRow As ListRow
    
    Set wsInput = ThisWorkbook.Worksheets("Manual_matching")
    Set tblArchive = ThisWorkbook.Worksheets("Archiwum_Manual_Matching").ListObjects("Tbl_Reczne_Archiwum")
    
    If wsInput.Range("B6").Value = "" Or wsInput.Range("G6").Value = "" Then
        MsgBox "Uzupe³nij oba numery ID!", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Call OdbezpieczWszystkieArkusze
    
    Set newRow = tblArchive.ListRows.Add
    newRow.Range.Value = wsInput.Range("B6:J6").Value
    
    wsInput.Range("B6,G6").ClearContents
    
    DoEvents
    ThisWorkbook.RefreshAll
    DoEvents
    
    Call ZabezpieczWszystkieArkusze
    Application.ScreenUpdating = True
    MsgBox "Dodano.", vbInformation
End Sub

