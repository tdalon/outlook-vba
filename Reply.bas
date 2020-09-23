Attribute VB_Name = "Reply"
' Module to ReplyToAll by a macro
' Clean-up Skype or Teams Meeting invitation from Body
Sub ReplyAll_Callback()
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    If TypeName(oItem) = "Nothing" Then
        MsgBox "No Item selected!"
        Exit Sub
    End If
    
    Call ReplyAll(oItem)
End Sub

Function ReplyAll(oItem As Object)
    Dim oReply As MailItem
    
    If TypeOf oItem Is MailItem Then
        Dim oMailItem As MailItem
        Set oMailItem = oItem
        Set oReply = oMailItem.ReplyAll
        Exit Function
    End If
    
    If TypeOf oItem Is AppointmentItem Then
        Dim oAppt As AppointmentItem
        Set oAppt = oItem
        '
        'Set oReply = oItem.ReplyAll ' Does not exist. only for MeetingItem
        
        ' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/e2fb4e2d-9e27-452e-a50a-7e5c0dac1af5/cant-get-vba-code-to-use-replyall-on-existing-meetings?forum=outlookdev
        
        ' Simulate click to ReplyAll
        oAppt.Display
        oAppt.GetInspector.CommandBars.ExecuteMso ("ReplyAll")
        oAppt.GetInspector.Close olDiscard
        
        Set oReply = Application.ActiveInspector.CurrentItem
        Do While oReply.Subject <> "RE: " & oAppt.Subject
            Set oReply = Application.ActiveInspector.CurrentItem
        Loop
        
        oReply.Display
        
        Dim objDoc As Word.Document
        Set objDoc = oReply.GetInspector.WordEditor ' Strange Error
        Dim oRng As Word.Range
        Set oRng = objDoc.Range
        'objDoc.Activate
        
        If InStr(oAppt.Location, "Skype ") Then
            sDelim = ".{137}"
        ElseIf InStr(oAppt.Location, "Microsoft Teams") Then
            sDelim = "_{80}"
        Else
            Exit Function
        End If
        
        
        
        With oRng.Find
            .MatchWildcards = True
            .Replacement.Text = ""
            .Text = sDelim & "*" & sDelim
            .Execute Replace:=wdReplaceAll
        End With
        
        
    End If
End Function

