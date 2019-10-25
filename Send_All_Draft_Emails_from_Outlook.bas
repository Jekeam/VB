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
