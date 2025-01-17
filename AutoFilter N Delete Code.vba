'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************    RemoveFilter
Function RemoveFilter()
If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
End Function

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
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Coyp_FilteredData_To_NewSheet
Sub CoypFilteredData_Citic()
Dim wsData      As Worksheet
Dim wsDest      As Worksheet
Dim lr          As Long
Application.ScreenUpdating = False
Sheets("OutGoing_Data").Activate
Set wsData = Worksheets("OutGoing_Data")
Set wsDest = Worksheets("OutGoing_Filter_Data")
lr = wsData.Cells(Rows.Count, "A").End(xlUp).Row
If wsData.FilterMode Then wsData.ShowAllData
With wsData.Rows(1)
    .AutoFilter Field:=1, Criteria1:="TAB"
 If wsData.Range("A1:A" & lr).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
            wsData.Range("A2:X" & lr).SpecialCells(xlCellTypeVisible).Copy wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Offset(1, 0)
        Range("A1").Activate
    End If
End With
Application.ScreenUpdating = True
End Sub
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Filterstuff()
    ' Select & Filter data
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = Worksheets("Main")
    Set rng = ws.Range("A2:AU" & ws.Range("A2").End(xlDown).Row)

    rng.AutoFilter

    ' Filter for things
    rng.AutoFilter Field:=39, Criteria1:="words"
    rng.AutoFilter Field:=43, Criteria1:="<>*wordswords*"

    ' Find the first unfiltered cell
    If rng.SpecialCells(xlCellTypeVisible).Count > rng.Columns.Count Then
    
        'Autofilter has returned at least one row of data
    Sheets("Output_File").Range("A1:W1000", Sheets("Output_File").Range("A1:W1000").End(xlDown)).Copy _
    Sheets("Destination").Range("A1")
          
    Else
        MsgBox "No data results from Autofilter"
        Exit Sub
    End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Filter criteria based on Sheet1 column A
'Search for the columns in Sheet2
'Copy Filtered rows(data) with header and paste it in new sheet
'rename the active sheet with MID in column A in Sheet1

Sub Create_Merchant_Name_Sheets_Start()
Sheets("Final_Macro_Sheet").Select
Range("A1").Select: Selection.AutoFilter
Columns("A:A").Select
Selection.Copy
Sheets("Hlp").Select
Range("D1").PasteSpecial
Columns("D:D").Select
ActiveSheet.Range("D:D").RemoveDuplicates Columns:=1, Header:=xlYes
lr_Hlp = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row

For i = 2 To lr_Hlp
Sheets("Hlp").Select
V_MID = Range("D" & i).Value
Sheets("Output_File").Select
ActiveSheet.Range("A:AZ").AutoFilter Field:=1, Criteria1:=V_MID

Sheets("Output_File").Range("A1:W1000", Sheets("Output_File").Range("A1:W1000").End(xlDown)).Copy _
  Sheets("Destination").Range("A1")

Dim name As String
merchantname = Sheets("Destination").Range("S2")
Sheets("Destination").name = merchantname
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'BELOW CODE SAVE THE ACTIVE SHEET WITH "MID NUMBER"
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'Gets the name of the currently visible worksheet
FileName = ActiveSheet.name
'Puts the worksheet into its own workbook
ThisWorkbook.ActiveSheet.Copy
'Saves the workbook - uses the name of the worksheet as the name of the new workbook
ActiveWorkbook.SaveAs "C:\CHBK_FA_Macro_2023\RB\Excel_Files\" & FileName & ".xlsx"
'Closes the newly created workbook so you are still looking at the original workbook
ActiveWorkbook.Close
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
Sheets.Add.name = "Destination"
Next
End Sub
