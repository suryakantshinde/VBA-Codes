'=========================================================================================================================
'**************************************************************************Macro - Negative sign in front selective numbers
Sub Negative()
    Dim lr As Long, i As Long
    lr = Range("A" & Rows.Count).End(xlUp).Row
    For i = 1 To lr
        If Range("A" & i) = "Surya" Then
            Range("B" & i) = Range("B" & i) * -1
        End If
    Next i
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

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Connect to SQL
'Add reference for Microsoft Activex Data Objects Library

Sub sbADO()
Dim sSQLQry As String
Dim ReturnArray

Dim Conn As New ADODB.Connection
Dim mrs As New ADODB.Recordset

Dim DBPath As String, sconnect As String

DBPath = ThisWorkbook.FullName

'You can provide the full path of your external file as shown below
'DBPath ="C:\InputData.xlsx"

sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';"

Conn.Open sconnect
    sSQLSting = "SELECT * From [DataSheet$]" ' Your SQL Statemnt (Table Name= Sheet Name=[DataSheet$])
    
    mrs.Open sSQLSting, Conn
        '=>Load the Data into an array
        'ReturnArray = mrs.GetRows
                ''OR''
        '=>Paste the data into a sheet
        ActiveSheet.Range("A2").CopyFromRecordset mrs
    'Close Recordset
    mrs.Close

'Close Connection
Conn.Close

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Connect to SQL
'Add reference for Microsoft Activex Data Objects Library

Sub sbADO2()
Dim sSQLQry As String
Dim ReturnArray

Dim Conn As New ADODB.Connection
Dim mrs As New ADODB.Recordset

Dim DBPath As String, sconnect As String

DBPath = ThisWorkbook.FullName

'You can provide the full path of your external file as shown below
'DBPath ="C:\InputData.xlsx"

'Using MSDASQL Provider
'sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & DBPath & ";HDR=Yes';"

'Using Microsoft.Jet.OLEDB Provider - If you get an issue with Jet OLEDN Provider try MSDASQL Provider (above statement)
sconnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath _
    & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
    
Conn.Open sconnect
    sSQLSting = "SELECT * From [DataSheet$]" ' Your SQL Statemnt (Table Name= Sheet Name=[DataSheet$])
    
    mrs.Open sSQLSting, Conn
        '=>Load the Data into an array
        'ReturnArray = mrs.GetRows
                ''OR''
        '=>Paste the data into a sheet
        ActiveSheet.Range("A2").CopyFromRecordset mrs
    'Close Recordset
    mrs.Close

'Close Connection
Conn.Close

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************ Message Box Example
Sub MessageBoxExample()
    Dim iRet As Integer
    Dim strPrompt As String
    Dim strTitle As String
 
    ' Promt
    strPrompt = "Have you coppied current month MIS again in Sheet Current_Month_MIS?"
 
    ' Dialog's Title
    strTitle = "My Tite"
 
    'Display MessageBox
    iRet = MsgBox(strPrompt, vbYesNo, strTitle)
 
    ' Check pressed button
    If iRet = vbNo Then
        MsgBox "Please Copy the current Month MIS in Sheet Current_Month_MIS then continue !"
    Else
        MsgBox "Yes!"
    End If
   
End Sub

'----------------------------------------------------------------------'----------------------------------------------------------------------
'Macro Message Box -  msgbox multiple lines
'----------------------------------------------------------------------'----------------------------------------------------------------------
        MsgBox "Cell AA1 does not contain CB Date." & vbCrLf & "Please change the data in right format." & vbCrLf & "Thank you", vbCritical
'----------------------------------------------------------------------'----------------------------------------------------------------------
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************ Date Filter
Sub Data_Date_Filter()
     
    Dim sDate As Variant, eDate As Variant
    
    Sheets("Sheet1").Select
    Columns("A:A").Select
    Selection.NumberFormat = "mm/dd/yyyy"
     
    sDate = Application.InputBox("Enter the starting date as mm/dd/yyyy", Type:=1 + 2)
    eDate = Application.InputBox("Enter the Ending date as mm/dd/yyyy", Type:=1 + 2)
     
    
    Application.ScreenUpdating = False
     
  Sheets("RawData_MAESTRO_PG").Cells.ClearContents
     
    With Sheets("SOURCE")
        .AutoFilterMode = False
        .Range("AA1").CurrentRegion.AutoFilter Field:=27, Criteria1:=">=" & sDate, Operator:=xlAnd, Criteria2:="<=" & eDate
        .Range("AA1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy Sheets("RawData_MAESTRO_PG").Range("A1")
    End With
     
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
     
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
"Select range at last row in column
Range(Selection, Selection.End(xlToRight)).Select
Range("A" & Rows.Count).End(xlUp).Select
Range("A2:A" & Lastrow).Select
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Last blank Row in column = Worksheets("Sheet1").Range("A1").End(xlDown).Row + 1
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Select the last blank row 
Sub Select_Last_Row()
'Step 1:  Declare Your Variables.
    Dim LastBlankRow As Long
'Step 2:  Capture the last used row number.
    LastBlankRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
'Step 3:  Select the next row down
    Cells(LastBlankRow, 1).Select
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ActiveCell.copy
ActiveCell.Offset(0, -1).Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~      
lastRow = Range("A" & Rows.Count).End(xlUp).Row
Range("G2").AutoFill Destination:=Range("G2:G" & lastRow)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Select entire Column till last row VBA
Dim LastRow As Long
With Worksheets("ScrapData")
  LastRow = .Cells(Rows.Count, "E").End(xlUp).row
  .range("E2:E" & LastRow).Select
End With
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PasteSpecial_All_Columns
Lastrow = Range("C" & Rows.Count).End(xlUp).Row '...DONT CHAGE "C"
Range("G2:G" & Lastrow).Select
Sub PasteSpecial_All_Columns()
Application.CutCopyMode = False
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = True
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Rearrange Sheets
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
