Sub Search_110()
    Dim LastRow As Long
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("Search 110").Cells.Delete
    LastRow = ThisWorkbook.Sheets("RawData").Range("A" & Rows.Count).End(xlUp).Row
    ThisWorkbook.Sheets("RawData").Activate
    ThisWorkbook.Sheets("RawData").Range("A1:A" & LastRow).Select
    For Each cel In Selection
        cel.Activate
        If cel.Value Like "*VSS-110*" Then
            If cel.Offset(2, 0).Value Like "*REPORTING FOR:      1000526662*" Then
                i = 1
                Do
                    ThisWorkbook.Sheets("Search 110").Range("A" & i).Value = ActiveCell.Value
                    ActiveCell.Offset(1, 0).Activate
                    i = i + 1
                Loop Until ActiveCell.Value Like "*END OF VSS-110 REPORT*"
            End If
        End If
    Next cel
    Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
Sub Search_120()
    Dim LastRow As Long
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("Search 120").Cells.Delete
    LastRow = ThisWorkbook.Sheets("RawData").Range("A" & Rows.Count).End(xlUp).Row
    ThisWorkbook.Sheets("RawData").Activate
    ThisWorkbook.Sheets("RawData").Range("A1:A" & LastRow).Select
    For Each cel In Selection
        cel.Activate
        If cel.Value Like "*VSS-120*" Then
            If cel.Offset(2, 0).Value Like "*REPORTING FOR:      1000526662*" Then
                i = 1
                Do
                    ThisWorkbook.Sheets("Search 120").Range("A" & i).Value = ActiveCell.Value
                    ActiveCell.Offset(1, 0).Activate
                    i = i + 1
                Loop Until ActiveCell.Value Like "*END OF VSS-120 REPORT*"
                
            End If
        End If
    Next cel
    Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
Sub Search_130()
    Dim LastRow As Long
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("Search 130").Cells.Delete
    LastRow = ThisWorkbook.Sheets("RawData").Range("A" & Rows.Count).End(xlUp).Row
    ThisWorkbook.Sheets("RawData").Activate
    ThisWorkbook.Sheets("RawData").Range("A1:A" & LastRow).Select
    For Each cel In Selection
        cel.Activate
        If cel.Value Like "*VSS-130*" Then
            If cel.Offset(2, 0).Value Like "*REPORTING FOR:      1000526662*" Then
                i = 1
                Do
                    ThisWorkbook.Sheets("Search 130").Range("A" & i).Value = ActiveCell.Value
                    ActiveCell.Offset(1, 0).Activate
                    i = i + 1
                Loop Until ActiveCell.Value Like "*END OF VSS-130 REPORT*"
            End If
        End If
    Next cel
    Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
Sub Search_N_Insert_Data_from_Sheets()
'On Error GoTo Surya
    Dim FindString As String
    Dim rng As Range
    Sheets("IMS_Weekly_APAC_OPS_KPI_Trend").Activate
    FindString = Sheets("IMS_Weekly_APAC_OPS_KPI_Trend").Range("A1").Value
    If Trim(FindString) <> "" Then
        With Sheets("IMS_Weekly_APAC_OPS_KPI_Trend").Range("C1:BW1")
            Set rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            Lookat:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not rng Is Nothing Then
       rng.Select
'--------------------------------------------------------------------------------
'''---------------------Boarding ------------------------
'''Total Active Merchant Base
Sheets("APAC_Data_Sheet").Range("B2:B2").Copy
Sheets("IMS_Weekly_APAC_OPS_KPI_Trend").Activate
ActiveCell.Offset(1, 0).Select
ActiveCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
rng.Select

'''Boarding HC (APAC Boarding)
Sheets("APAC_Data_Sheet").Range("B3:B3").Copy
Sheets("IMS_Weekly_APAC_OPS_KPI_Trend").Activate
ActiveCell.Offset(2, 0).Select
ActiveCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
rng.Select
End If
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
Sub Search_Date()
'On Error GoTo Surya
Sheets("IDFC_Weekly_APAC_OPS_KPI_Trend").Activate
    Range("AA4:AG4").Select
    Selection.Find(What:="4-5 Days", After:=ActiveCell, LookIn:=xlFormulas, _
        Lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(, 0).Resize(1, 1).Select
ActiveCell.Offset(1, 0).Resize(30, 1).Select
Selection.Copy
Range("F5").Activate
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=18
Application.CutCopyMode = False
'Surya:
'Exit Sub
End Sub

