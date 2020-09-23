Attribute VB_Name = "Copy"
' Module to Copy an Item by a macro
' Open directly for edit, check reminder
Sub Copy_Callback()
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    If TypeName(oItem) = "Nothing" Then
        MsgBox "No Item selected!"
        Exit Sub
    End If
    
    Call Copy(oItem)
End Sub

Function Copy(oItem As Object)
    'Dim oCopy As Object
    
    If TypeOf oItem Is MailItem Then
        'Dim oCopy As MailItem
      
        oItem.Copy.Display
   
    
    ElseIf TypeOf oItem Is AppointmentItem Then
        Dim oCopy As AppointmentItem
       
        'Set oCopy = oItem.Copy
        Set oCopy = Outlook.CreateItem(olAppointmentItem)
        With oCopy
            .Subject = oItem.Subject & " - Follow up from " & Format(oItem.Start, "yyyy-mm-dd")
            
            .Location = oItem.Location
            '.Body = oItem.Body ' TODO Copy Body keeping HTML format via Word editor
            .Categories = oItem.Categories
            .AllDayEvent = oItem.AllDayEvent
        End With
        
        oCopy.MeetingStatus = olMeeting          ' Convert to Meeting
    
 
        For Each Recip In oItem.Recipients
            'To copy the attendees who have accepted the meeting request
            'If obj.MeetingResponseStatus = olResponseAccepted Then
            'To copy who declined - "If olAttendee.MeetingResponseStatus = olResponseDeclined Then"
            'To copy who haven't respond - "If olAttendee.MeetingResponseStatus = olResponseNone" Then
            '  strAddrs = strAddrs & ";" & obj.Address
            'End If
            ' Skip organizer Automatically done at ResolveAll
            Set CopyRecip = oCopy.Recipients.Add(Recip.Address)
            CopyRecip.Type = Recip.Type
        Next
    
        oCopy.Recipients.ResolveAll
    
        ' Set Reminder if Meeting
        If (oCopy.MeetingStatus = olNonMeeting) And (oCopy.ReminderSet = False) Then
            oCopy.ReminderSet = True
            oCopy.ReminderMinutesBeforeStart = 15 ' Enter your default time
        End If
       
        
        
        ' TODO Copy Body keeping HTML format via Word editor
        oItem.Display
      
        Dim objDoc As Word.Document
        Dim oRng As Word.Range
        Set objDoc = oItem.GetInspector.WordEditor ' Needs to be displayed
        objDoc.Range.Copy
        
        oCopy.Display
               
        Set objDoc = oCopy.GetInspector.WordEditor
        objDoc.Range.Paste
      
        oItem.GetInspector.Close olPromptForSave 'olDiscard
       
      
       
        ' Clean-up Skype/Teams invitation
        
        Set oRng = objDoc.Range
        'objDoc.Activate
        
        If InStr(oCopy.Location, "Skype") Then   ' can be Skype-Besprechung
            sDelim = ".{137}"
        ElseIf InStr(oCopy.Location, "Microsoft Teams") Then
            sDelim = "_{80}"
        Else
            Exit Function
        End If
        
        With objDoc.Range.Find
            .MatchWildcards = True
            .Replacement.Text = ""
            .Text = sDelim & "*" & sDelim
            .Execute Replace:=wdReplaceAll
        End With
       
       
    End If
End Function

