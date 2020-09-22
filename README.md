<div align="center">

## Save Outlook Mail Attachment


</div>

### Description

This code snippet is actually a Macro for Outlook 97, 98 or 2000 but can be easily instituted into VB by creating your Outlook.Application object to completely automate the mod from VB. It could also be used with MAPI mail as well.
 
### More Info
 
None...

Basic understanding of VBA and VB as well as Office Automation through VB.

None... Could be set as a True or False Method for total automation

None so far...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Kesler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-kesler.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-kesler-save-outlook-mail-attachment__1-27911/archive/master.zip)





### Source Code

```
Sub AutomateMe()
 Dim oApp As Application
 Dim oNS As NameSpace
 Dim oMsg As Object
 Dim oAttachments As Outlook.Attachments
 Dim strControl
 Set oApp = New Outlook.Application
 Set oNS = oApp.GetNamespace("MAPI")
 'Set folder to check the INBOX
 Set oFolder = oNS.GetDefaultFolder(olFolderInbox)
 strControl = 0
 For Each oMsg In oFolder.Items
 With oMsg
'Check for new mail (unread=true)
 If .UnRead Then
'This could use the .Subject as well to search for text in the subject line.
 If InStr(1, .Body, "Body Text to look for") > 0 Then
  oMsg.Attachments.Item(1).SaveAsFile "Your Drive:\Your Path\" _
  & oMsg.Attachments.Item(1).DisplayName
'Set mailItem to read by setting UnRead flag to false.
  .UnRead = False
  Exit Sub
 End If
 End If
 End With
 Next
End Sub
```

