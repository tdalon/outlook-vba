Attribute VB_Name = "Item2TeamsChat"
' http://github.conti.de/gist/uid41890/7a9df0e7ea1fb7b12f35d7c7547f6186
' Requires Utils module:  http://github.conti.de/gist/uid41890/6d98ec3624920b4f7a2ab2a8674c1263

' ###### Item to Chat

Sub Item2Chat_Callback()
    Call Item2Chat
End Sub

Sub Item2Chat(Optional ByVal onlyToSender As Boolean = False, Optional ByVal ansBody As String = "")
    ' Open Teams Chat from Outlook Item (Email or Appointment)
    ' ansBody - option to copy body to the clipboard or not
    '         empty: user will be asked
    '         vbYes/vbNo: item body will be copied/or not to the chat message
    
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    
    sLink = Item2ChatLink(oItem, onlyToSender)
    
    
    ' Copy Email Body to Clipboard
    'ansBody = vbYes
    If TypeOf oItem Is Outlook.MeetingItem Then
        ansBody = vbNo
    End If
    If ansBody = "" Then
        ansBody = MsgBox("Do you want to copy body to the Clipboard?", vbYesNo + vbQuestion + vbDefaultButton1, "Copy Body?")
    End If
    If (ansBody = vbYes) Then
        CopyToClipboardHTML (oItem.HTMLBody)
    End If
    
    'Call RunLink(sLink)
    sLink = Replace(sLink, "https://teams.microsoft.com", "msteams:")
    openUrl (sLink)
    'Debug.Print sLink
End Sub

Function Item2ChatLink(oItem As Object, Optional ByVal onlyToSender As Boolean = False, Optional ByVal ansCopySubjectToMessage As String = "") As String
    ' Item2Chat(Optional ByVal onlyToSender As Boolean = False)
    ' Item2Chat(True) - only open Chat window with Sender
    ' Calls Utils: GetCurrentItem, RunLink,CopyToClipboardHTML*
    
    
    Dim sTopicName As String
    Select Case True
        Case TypeOf oItem Is Outlook.MailItem
            'Dim oMailItem As Outlook.MailItem
            'Set oMailItem = oItem
            sTopicName = "[Email] " & oItem.Subject
        Case TypeOf oItem Is Outlook.MeetingItem
            sTopicName = "[Meeting] " & oItem.Subject
        Case Else
            
    End Select
    
    ' Input box for TopicName
    sTopicName = InputBox("Enter the Chat Group Name", "Outlook to Chat", sTopicName)
    
    Dim sEmails As String
    Dim sLink As String
    
    Dim oRecip As Outlook.Recipient
    
    sLink = "https://teams.microsoft.com/l/chat/0/0?users=" ' Deep Link to Chat
    
    
    ' Do you want to add From Email?
    ansFrom = vbYes
    ansTo = vbYes
    
    
    ' Do you want to add CC Emails?
    If onlyToSender Then
        ansCc = vbNo
    Else
        MsgBoxCustom_Set vbYes, "All"
        MsgBoxCustom_Set vbNo, "No Cc"
        MsgBoxCustom ans, "Which Recipients do you wan to include in the Group chat?", vbYesNo + vbQuestion, "Outlook To Chat: Recipients"
        If (ans = vbYes) Then
            ansCcc = vbYes
        ElseIf (ans = vbNo) Then
            ansCc = vbNo
        End If
    End If
    
    curEmail = GetEmailAddress()
    
    ' Add From Email
    'Debug.Print GetSenderEmail(oItem)
    
    
    For Each oRecip In oItem.Recipients
        sNewEmail = Recip2Email(oRecip)
        If (sNewEmail = curEmail) Then
            GoTo NextRecip
        End If
        If TypeOf oItem Is Outlook.MailItem Then
            If (ansTo = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olTo) Then
                sEmails = sEmails & sNewEmail & ","
            ElseIf (ansCc = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olCC) Then
                sEmails = sEmails & sNewEmail & ","
            ElseIf (ansFrom = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olOriginator) Then
                sEmails = sEmails & sNewEmail & ","
            End If
        ElseIf TypeOf oItem Is Outlook.MeetingItem Then
            If (ansTo = vbYes) And (oRecip.Type = Outlook.OlMeetingRecipientType.olRequired) Then
                sEmails = sEmails & sNewEmail & ","
            ElseIf (ansCc = vbYes) And (oRecip.Type = Outlook.OlMeetingRecipientType.olOptional) Then
                sEmails = sEmails & sNewEmail & ","
            ElseIf (ansFrom = vbYes) And (oRecip.Type = Outlook.OlMeetingRecipientType.olOrganizer) Then
                sEmails = sEmails & sNewEmail & ","
            End If
            
        End If
NextRecip:
    Next oRecip
    
    
    
    ' Remove trailing ,
    sEmails = Left(sEmails, Len(sEmails) - 1)
    
    sLink = sLink & sEmails
    
    ' Name Group Chat - does not work for special characters needs to encode
    sLink = sLink & "&topicName=" & Replace(sTopicName, ":", "") ': are not supported
    
    ' Add TopicName - won't be set if group already existed
    ' sLink = sLink & "&topicName=RE:" & oMailItem.Subject
    
    ' Add message
    
    If InStr(sEmails, ",") Then ' Group Chat more than 1-1 can be named => do not copy subject to message because already in Chat Subject
    ansCopySubjectToMessage = vbNo
End If
If ansCopySubjectToMessage = "" Then
    ansCopySubjectToMessage = MsgBox("Do you want to copy Subject to the Chat Message?", vbYesNo + vbQuestion, "Copy Subject?")
End If

If (ansCopySubjectToMessage = vbYes) Then
    sLink = sLink & "&message=" & sTopicName   '& vbCrl & oMailItem.Body 'Markdown formatting does not work e.g. ## or * *
End If


Item2ChatLink = sLink

End Function

' ###########################
' Item to Teams Meeting
Sub CopyItem2TeamsMeetingLink()
    ' Calls Item2TeamsMeetingLink
    
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    Dim sText, sHtml As String
    
    sLink = Item2TeamsMeetingLink(oItem)
    
    sText = oItem.Subject
    sText = InputBox("Enter Link Display Text", "Link Text", sText)
    sHtml = "<a href=""" & sLink & """>" & sText & "</a>"
    Call CopyToClipboardHTML(sHtml, sText)
    'Debug.Print sLink
End Sub

Function Item2TeamsMeetingLink(oItem As Object, Optional ByVal onlyToSender As Boolean = False, Optional ByVal ansCopySubjectToMessage As String = "") As String
    ' Item2Chat(Optional ByVal onlyToSender As Boolean = False)
    ' Item2Chat(True) - only open Chat window with Sender
    ' Calls Utils: GetCurrentItem, RunLink,CopyToClipboardHTML*
    
    
    
    Dim sEmails As String
    Dim sLink As String
    
    Dim oRecip As Outlook.Recipient
    
    sLink = "https://teams.microsoft.com/l/meeting/new" ' Deep Link to New Meeting
    
    
    ' Do you want to add From Email?
    'ansFrom = vbYes
    If onlyToSender Then
        ansFrom = vbNo
    Else
        ansFrom = MsgBox("Do you want to include Sender (From)?", vbYesNo + vbQuestion, "Include Sender (From)?")
    End If
    
    ' Do you want to add To Emails?
    If onlyToSender Then
        ansTo = vbYes
    Else
        ansTo = MsgBox("Do you want to include all recipients in To?", vbYesNo + vbQuestion, "Include Recipients in To?")
    End If
    
    ' Do you want to add CC Emails?
    If onlyToSender Then
        ansCc = vbNo
    Else
        ansCc = MsgBox("Do you want to include all recipients in CC?", vbYesNo + vbQuestion + vbDefaultButton2, "Include Recipients in CC?")
    End If
    
    
    For Each oRecip In oItem.Recipients
        
        If ((ansTo = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olTo)) Or ((ansCc = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olCC)) Then
            
            sEmails = sEmails & Recip2Email(oRecip) & ","
            
        End If
        
    Next oRecip
    
    ' Add From Email
    'Debug.Print GetSenderEmail(oItem)
    If (ansFrom = vbYes) Then
        sEmails = sEmails & GetSenderEmail(oItem) & ","
    End If
    
    
    ' Remove trailing ,
    sEmails = Left(sEmails, Len(sEmails) - 1)
    
    
    sLink = sLink & "?attendees=" & sEmails
    
    
    ' Add TopicName - won't be set if group already existed
    ' sLink = sLink & "&topicName=RE:" & oMailItem.Subject
    
    ' Copy Subject
    sLink = sLink & "&subject=" & oItem.Subject  '& vbCrl & oMailItem.Body
    
    ' Copy Message
    sLink = sLink & "&content=" & oItem.Body
    
    Item2TeamsMeetingLink = sLink
    
    
End Function

Sub Item2ChatCopyLink_Callback()
    Call CopyItem2ChatLink(True, vbNo)
End Sub

Sub CopyItem2ChatLink(Optional ByVal onlyToSender As Boolean = True, Optional ByVal copySubjectToMessage As String = vbNo)
    ' Calls Item2ChatLink
    
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    Dim sText, sHtml As String
    
    sLink = Item2ChatLink(oItem, onlyToSender, copySubjectToMessage)
    
    sText = oItem.Subject
    sText = InputBox("Enter Link Display Text", "Link Text", sText)
    sHtml = "<a href=""" & sLink & """>" & sText & "</a>"
    Call CopyToClipboardHTML(sHtml, sText)
    'Debug.Print sLink
End Sub

' UTILS
Function Recipients2StringEmail(oItem As Object, Optional ByVal ansTo As String = vbYes, Optional ByVal ansCc As String = "") As String
    ' Calls Recip2Email
    
    Dim sEmails As String
    
    If ansTo = "" Then
        ansTo = MsgBox("Do you want to include all recipients in To?", vbYesNo + vbQuestion, "Include Recipients in To?")
    End If
    If ansCc = "" Then
        ansCc = MsgBox("Do you want to include all recipients in CC?", vbYesNo + vbQuestion + vbDefaultButton2, "Include Recipients in CC?")
    End If
    
    Dim oRecip As Recipient
    For Each oRecip In oItem.Recipients
        
        If ((ansTo = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olTo)) Or ((ansCc = vbYes) And (oRecip.Type = Outlook.OlMailRecipientType.olCC)) Then
            
            sEmails = sEmails & Recip2Email(oRecip) & ","
            
        End If
        
    Next oRecip
    
    
    ' Remove trailing ,
    sEmails = Left(sEmails, Len(sEmails) - 1)
    
    Recipients2StringEmail = sEmails
    
    
    
End Function

Function Recip2Email(oRecip As Outlook.Recipient) As String
    ' https://stackoverflow.com/a/51939384/2043349
    ' takes a Display Name (i.e. "James Smith") and turns it into an email address (james.smith@myco.com)
    ' necessary because the Outlook address is a long, convoluted string when the email is going to someone in the organization.
    ' source:  https://stackoverflow.com/questions/31161726/creating-a-check-names-button-in-excel
    
    Dim oEU As Object                            'Outlook.ExchangeUser
    If oRecip.Resolved Then
        Select Case oRecip.AddressEntry.AddressEntryUserType
            Case 0, 5                                'olExchangeUserAddressEntry & olExchangeRemoteUserAddressEntry
                Set oEU = oRecip.AddressEntry.GetExchangeUser
                If Not (oEU Is Nothing) Then
                    Recip2Email = oEU.PrimarySmtpAddress
                End If
            Case 10, 30                              'olOutlookContactAddressEntry & 'olSmtpAddressEntry
                Recip2Email = oRecip.AddressEntry.Address
        End Select
    End If
End Function

Function GetSenderEmail(oM As Variant)
    ' https://stackoverflow.com/a/52150247/2043349
    
    Dim oPA As Outlook.PropertyAccessor
    Set oPA = oM.PropertyAccessor
    
    GetSenderEmail = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")
    
End Function

