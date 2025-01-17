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
