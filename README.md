<h2>How to Batch Send Multiple Draft Emails with Outlook VBA</h2>
```
Sub SendAllDraftEmails()
    Dim objDrafts As Outlook.Items
    Dim objDraft As Object
    Dim strPrompt As String
    Dim nResponse As Integer
    Dim i As Long
 
    Set objDrafts = Outlook.Application.Session.GetDefaultFolder(olFolderDrafts).Items
 
    If objDrafts.Count > o Then
       strPrompt = "Are you sure to send out all the drafts?"
       nResponse = MsgBox(strPrompt, vbQuestion + vbYesNo, "Confirm Sending")
 
       If nResponse = vbYes Then
          For i = objDrafts.Count To 1 Step -1
              objDrafts.Item(i).Send
          Next
       End If
    Else
       MsgBox ("No Drafts!")
    End If
End Sub
```

1. At first, launch Outlook application and press “Alt + F11” shortcuts. 
2. Then you will open the VBA editor window, in which you should open a new module.
3. Subsequently, copy and paste the following VBA codes into it. 
4. After that, you can exit the VBA editor and proceed to add the VBA project to Quick Access Toolbar or ribbon. Here we will take Quick Access Toolbar as an example.
5. Firstly, go to “File” > “Options” > “Quick Access Toolbar” tab. 
6. Then follow the steps shown in the picture below to add the new macro to Quick Access Toolbar.
7. Finally you can back to main Outlook window. You will see the new button in Quick Access Toolbar. 
8. If there is no item in the Drafts folder, when you click the button, you will receive a message like the following screenshot.
