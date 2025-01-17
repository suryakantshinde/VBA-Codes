'------------------------------------------------------------------------------------------------------------------------
'******************************************************************** Saving All Attachments to a Folder

Sub Saving_All_Attachments_to_Folder()
'Step 1:  Declare your variables
    Dim ns As Namespace
    Dim MyInbox As MAPIFolder
    Dim MItem As MailItem
    Dim Atmt As Attachment
    Dim FileName As String
'Step 2:  Set a reference to your inbox
    Set ns = GetNamespace("MAPI")
    Set MyInbox = ns.GetDefaultFolder(olFolderInbox)
'Step 3:  Check for messages in your inbox; exit if none
    If MyInbox.Items.Count = 0 Then
    MsgBox "No messages in folder."
    Exit Sub
    End If
'Step 4:  Create directory to hold attachments
    On Error Resume Next
    MkDir "C:\OffTheGrid\MyAttachments\"
'Step 5:  Start to loop through each mail item
    For Each MItem In MyInbox.Items
'Step 6:  Save each attachement then go to the next attachment
    For Each Atmt In MItem.Attachments
    FileName = "C:\OffTheGrid\MyAttachments\" & Atmt.FileName
    Atmt.SaveAsFile FileName
    Next Atmt
'Step 7:  Move to the next mail item
    Next MItem
'Step 8:  Memory cleanup
    Set ns = Nothing
    Set MyInbox = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------------------
'********************************************************************  Mailing the Active Workbook as Attachment
    Sub Mailing_the_Active_Workbook_as_Attachment()
'Step 1:  Declare your variables
    Dim OLApp As Outlook.Application
    Dim OLMail As Object
'Step 2:  Open Outlook start a new mail item
    Set OLApp = New Outlook.Application
    Set OLMail = OLApp.CreateItem(0)
    OLApp.Session.Logon
'Step 3:  Build your mail item and send
    With OLMail
    .To = "admin@datapigtechnologies.com; mike@datapigtechnologies.com"
    .CC = ""
    .BCC = ""
    .Subject = "This is the Subject line"
    .Body = "Hi there"
    .Attachments.Add ActiveWorkbook.FullName
    .Display  'Change to .Send to send without reviewing
    End With
'Step 4:  Memory cleanup
    Set OLMail = Nothing
    Set OLApp = Nothing
End Sub
