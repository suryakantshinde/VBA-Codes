'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Select the Specific "Data" in Autofilter and delete Full Row
Sub Delete()
With ActiveSheet
    .AutoFilterMode = False
    With Range("B1", Range("B" & Rows.Count).End(xlUp))
        .AutoFilter 1, "*====END OF FILE====*"
        On Error Resume Next
        .Offset(1).SpecialCells(12).EntireRow.Delete
    End With
    .AutoFilterMode = False
'----------------------------------------------------------------------
 With ActiveSheet
    .AutoFilterMode = False
    With Range("B1", Range("B" & Rows.Count).End(xlUp))
        .AutoFilter 1, "Merchant No."
        On Error Resume Next
        .Offset(1).SpecialCells(12).EntireRow.Delete
    End With
    .AutoFilterMode = False
'----------------------------------------------------------------------
 End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Check if Filter contains data
Range("A1:G" & LastRow).Select
Selection.AutoFilter
Range("A1").Select
ActiveSheet.Range("A1:G" & LastRow).AutoFilter Field:=6, Criteria1:="<>EMI*"
lr = Cells(Rows.Count, 1).End(xlUp).Row
 If lr > 1 Then
Range("F2:F" & lr).ClearContents
End If
Selection.AutoFilter
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Filter n Replace data in Cell
Range("A1").Select: Selection.AutoFilter
ActiveSheet.Range("A:AG").AutoFilter Field:=2, Criteria1:="RRQ"
lr_fil = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count
If lr_fil > 1 Then
Range("B2", "B" & lr_IPG).Select
Selection.SpecialCells(xlCellTypeVisible).Value = "Retrieval"
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************ Select the Specific column and enter data in Adjacent cell
Sub City_AP() 'AP --> Rest of AP Column 'J=11, B=2
 RemoveFilter
Sheets("BD Closed Call Report1").Select
Range("A1:AF1").Select
Selection.AutoFilter
    ActiveSheet.Range("A1:AF1").AutoFilter Field:=10, Criteria1:="#N/A"
    ActiveSheet.Range("A1:AF1").AutoFilter Field:=2, Criteria1:="AP"

     Set rng = Range("A2", Range("A655326").End(xlUp)).SpecialCells(xlCellTypeVisible)
    For Each cell In rng
        Range("J" & cell.Row).Value = "Rest of AP"
    Next cell
    Application.CutCopyMode = False
  RemoveFilter
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Coyp_FilteredData_To_NewSheet
Sub CoypFilteredData_Citic()
Dim wsData      As Worksheet
Dim wsDest      As Worksheet
Dim lr          As Long
Application.ScreenUpdating = False
Sheets("CITIC Merch Outstd Fund Rept-2").Activate
Set wsData = Worksheets("CITIC Merch Outstd Fund Rept-2")
Set wsDest = Worksheets("Sheet2")
lr = wsData.Cells(Rows.Count, "H").End(xlUp).Row
If wsData.FilterMode Then wsData.ShowAllData
With wsData.Rows(1)
    .AutoFilter Field:=8, Criteria1:="CITIC"
    If wsData.Range("H1:H" & lr).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
        wsData.Range("G2:G" & lr).SpecialCells(xlCellTypeVisible).Copy wsDest.Range("A" & Rows.Count).End(3)(2)
        wsDest.UsedRange.Borders.ColorIndex = xlNone
        wsDest.Select
        Range("A1").Activate
    End If
    .AutoFilter Field:=8
End With
Application.ScreenUpdating = True
End Sub
