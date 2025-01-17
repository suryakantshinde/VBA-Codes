Sub SaveTimeStampedBackup()
'Create variable to hold the new file path
Dim saveAsName As String
'Set the file path
saveAsName = ActiveWorkbook.Path & "\" & _
Format(Now, "yymmdd-hhmmss") & " " & ActiveWorkbook.name
'Save the workbook
ActiveWorkbook.SaveCopyAs FileName:=saveAsName
End Sub
