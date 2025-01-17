Sub Delete_Sheets()
    'Step 1: Tell Excel what to do if Error
        On Error GoTo MyError
    'Step 2:  Add a sheet and name it

Application.DisplayAlerts = False
    Sheets("SheetName").Delete
    Sheets("SheetName").Delete
Application.DisplayAlerts = True
Exit Sub
    'Step 3: If here, an error happened; tell the user
MyError:
        MsgBox "Sheet Not Found.Click OK to Continue", vbInformation, "Focus Audit Macro"
End Sub
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
Sub Insert_Sheets()
                    
    'Step 1: Tell Excel what to do if Error
        On Error GoTo MyError
                   
    'Step 2:  Add a sheet and name it
With Sheets
    .Add().name = "Sheet1"
    .Add().name = "Sheet2"
End With
        Exit Sub
    'Step 3: If here, an error happened; tell the user
MyError:
        MsgBox "There is already a sheet called that."
                    
    End Sub
'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
Sub Delete_Sheets()
    'Step 1: Tell Excel what to do if Error
On Error Resume Next
    'Step 2:  Add a sheet and name it
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    Application.DisplayAlerts = False
For Each ws In ThisWorkbook.Worksheets
    If ws.CodeName <> "INSTRUCTION" Then ws.Delete
Next
Application.DisplayAlerts = True
Exit Sub
    'Step 3: If here, an error happened; tell the user
'MyError:
'        MsgBox "Sheet Not Found.Click OK to Continue", vbInformation, "Focus Audit Macro"
End Sub
