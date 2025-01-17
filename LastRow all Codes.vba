'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************LastRow Code
LastRow = Range("A" & Rows.Count).End(xlUp).Row
Range("G2").AutoFill Destination:=Range("G2:G" & LastRow)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************PasteSpecial
ActiveCell.Copy
ActiveCell.Offset(0, -1).Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************PasteSpecial
Sub PasteSpecial_All_Columns()
'selecting range to last completely blank row
Sheets("RawData").Select
Range("A1:A" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select
Application.CutCopyMode = False
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = True
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
''************************************************************************Select Record Till Last Used
Sub Select_Record_Till_Last_Used()
Dim lr As Long
lr = Cells(Rows.Count, 1).End(xlUp).Row
Range("D2:D" & lr).Select
End Sub
