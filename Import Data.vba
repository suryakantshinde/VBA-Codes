Sub Import_Data()
On Error GoTo Errormask
    Sheets("RawData").Activate
    Range("A1").Select
            Dim wbCSV   As Workbook
            Dim wsMstr  As Worksheet:   Set wsMstr = ThisWorkbook.Sheets("RawData")
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------------------------------------------------------------
                                         'Change Path as per the need
            Dim fPath   As String:      fPath = "C:\Instant Settlement Recon - Automation\Input_Files\"
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            Dim fCSV    As String
            Application.ScreenUpdating = False '........'Speed up macro
            fCSV = Dir(fPath & "Transaction_Dump.xlsx") '.........' Select file
                    Set wbCSV = Workbooks.Open(fPath & fCSV)
                    ActiveSheet.UsedRange.Copy wsMstr.Range("A" & Rows.Count).End(xlUp).Offset(1)
                    wbCSV.Close False
    Application.DisplayAlerts = False
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
ThisWorkbook.ActiveSheet.Cells.ClearFormats
ActiveSheet.UsedRange.EntireColumn.AutoFit
Worksheets("RawData").Cells.NumberFormat = "General"
Exit Sub
Errormask:
MsgBox "File Name - SUBMISSION_SEARCH_All is incorrect", vbInformation
Exit Sub
End Sub
'=============================================================================================================================================================
'=============================================================================================================================================================
'Import_Data_from_ input folder - 'DMPH Cases Dump
Sub Import_Date_Dynamic_Path()
On Error GoTo Errormask
    Sheets("DMPH Cases Dump").Activate
    Range("A1").Select
            Dim wbCSV   As Workbook
            Dim wsMstr  As Worksheet:   Set wsMstr = ThisWorkbook.Sheets("DMPH Cases Dump")
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------------------------------------------------------------
                                         'Change Path as per the need
            Dim fPath   As String:      fPath = ThisWorkbook.Path & "/"
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------------------------------------------------------------------
            Dim fCSV    As String
            Application.ScreenUpdating = False '........'Speed up macro
            fCSV = Dir(fPath & "RawData.xlsx") '.........' Select file
                    Set wbCSV = Workbooks.Open(fPath & fCSV)
                    ActiveSheet.UsedRange.Copy wsMstr.Range("A" & Rows.Count).End(xlUp).Offset(1)
                    wbCSV.Close False
    Application.DisplayAlerts = False
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
ThisWorkbook.ActiveSheet.Cells.ClearFormats
ActiveSheet.UsedRange.EntireColumn.AutoFit
Exit Sub
Errormask:
MsgBox "File Not Found", vbInformation
Exit Sub
End Sub
