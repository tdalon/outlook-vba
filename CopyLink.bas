Attribute VB_Name = "CopyLink"
Sub CopyLink_Callback()
    ' Calls: CopyToClipboardHTML - requires fclip
    Dim objMail As Outlook.MailItem
    Dim sLink As String
    Dim sHtml As String
    
    'One and ONLY one message must be selected
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox ("Select one and ONLY one message.")
        Exit Sub
    End If
    
    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    
    ' Using fclip
    
    sLink = "outlook:" & objMail.EntryID
    sText = objMail.Subject & " (" + objMail.SenderName & ")"
    sHtml = "<a href=" & sLink & ">" & sText & "</a>"
    Call CopyToClipboardHTML(sHtml, sLink)
    MsgBox ("Outlook link was copied to the clipboard.")
    
End Sub

