''************************************************************************ Calculate Run Time Minutes ************************************************************************
Sub CalculateRunTime_Minutes()
'PURPOSE: Determine how many minutes it took for code to completely run

Dim StartTime As Double
Dim MinutesElapsed As String

'Remember time when macro starts
  StartTime = Timer

'*****************************
'Insert Your Code Here...
Data_Cleaning.Data_Cleaning_Start

'*****************************
If Application.Wait(Now + TimeValue("0:00:10")) Then
 'MsgBox "Time expired"
End If
'Determine how many seconds code took to run
  MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
  MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub
