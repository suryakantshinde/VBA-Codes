'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Save specific sheets and add date and time
' Save specific sheets and add date and time

Sub SaveWorkbook()
Dim YesterD As Date
Dim Fname As String, ws As Worksheet
Fname = Sheets("Monthly_Data").Range("A1").Value
'add sheet name in bracket
Sheets(Array("Day Wise Summary", "City Wise Summary", "Installation Closed Calls Table", "Installation Closed Calls Chat", "Monthly_Data")).Copy
For Each ws In ActiveWorkbook.Worksheets
Next ws
With ActiveWorkbook
    .SaveAs FileName:="C:\All_In_One_Macro\Output_Files\DR20\Installation_Close_Call" & "_" & Format(DateAdd("d", -1, Date), "DD_MMM_YYYY") & ".xlsx"
    .Close
End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Save Filename as Cell Value with Current Date
Sub SaveAs()
    Dim SavePath    As String
    Dim SaveAs      As String
    Dim FileName    As String
    Dim sDate       As String

    '// Save it Path
    SavePath = "Z:\Regional Weekly Report\11 May\Forecast Template\"

    '// File Name
    FileName = "Fiscal 2015 Weekly Projections "

    '// Format the on "B4" to YYYY-MM-DD
    sDate = Format(Sheets("Sheet1").Range("B4"), "YYYY-MM-DD")

    '// Save with File Name & Date & .pdf
    SaveAs = FileName & sDate & ".pdf"
        Application.DisplayAlerts = True

        '// Export Active Sheet as pdf
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
            SavePath & SaveAs

        Application.DisplayAlerts = True
    End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************  Rearrange Sheets
Sub RearrangeSheets()
    Dim ws As Worksheet
    Dim orderList() As Variant
    Dim i As Long
    
    ' Define the custom order list
    orderList = Array("D-Sheet", "B-Sheet", "F-Sheet", "A-Sheet", "C-Sheet")
    
    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Loop through each sheet name in the custom order list
    For i = LBound(orderList) To UBound(orderList)
        ' Find the sheet with the corresponding name
       ' On Error Resume Next
        Set ws = Worksheets(orderList(i))
        On Error GoTo 0
        
        ' Move the sheet to the desired position
        If Not ws Is Nothing Then
            ws.Move Before:=Worksheets(1)
        End If
    Next i
    
    ' Enable screen updating again
    Application.ScreenUpdating = True
End Sub
