Attribute VB_Name = "Appointment"

'Const Const_DefSkypeDialIn As String = "C:\Users\uid41890\Documents\GitHub\outlook-vba\Skype_DialIn.docx"
Const Const_DefSkypeDialIn As String = "C:\Users\uid41890\Documents\GitHub\outlook-vba" ' Directory without ending filesep
Const Const_DefTeamsDialIn As String = "C:\Users\uid41890\Documents\GitHub\outlook-vba\Teams_DialIn.docx"

Sub ToggleEnableForwarding()
    ' https://www.extendoffice.com/documents/outlook/4607-outlook-prevent-forwarding-meeting-invitation.html
    
    Dim xCurrentItem As Object
    
    Set xCurrentItem = Outlook.ActiveInspector.CurrentItem
    xCurrentItem.Actions("Forward").Enabled = Not (xCurrentItem.Actions("Forward").Enabled)
    If xCurrentItem.Actions("Forward").Enabled = False Then
        MsgBox "Forwarding  current meeting has been disabled. Any meeting attendee is prevented from forwarding this meeting."
    Else
        MsgBox "Forwarding  current meeting has been enabled."
    End If
    
End Sub

Public Sub MarkAsMIT()
' Mark selected Items in Calendar As Free and assign category "MIT"
Dim objItem As Object

Set coll = GetCurrentItems
If coll.Count = 0 Then
    Exit Sub
End If

For Each objItem In coll
    If TypeOf objItem Is Outlook.AppointmentItem Then
        objItem.Categories = "MIT"
        objItem.BusyStatus = olFree          ' Mark As Free
        objItem.Save
    End If
Next


End Sub

' https://gist.github.com/tdalon/60a746cfda75ad191e426ee421324386
Sub CheckTodayReminders()
    ' https://www.datanumen.com/blogs/quickly-send-todays-appointments-someone-via-outlook-vba/
    Dim objAppointments As Outlook.Items
    Dim objTodayAppointments As Outlook.Items
    Dim strFilter As String
    Dim objAppointment As Outlook.AppointmentItem ' Object
    
    Set objAppointments = Application.Session.GetDefaultFolder(olFolderCalendar).Items
    objAppointments.IncludeRecurrences = True
    objAppointments.Sort "[Start]", False        ' Bug: use False/descending see https://social.msdn.microsoft.com/Forums/office/en-US/919e1aee-ae67-488f-9adc-2c8518854b2a/how-to-get-recurring-appointment-current-date?forum=outlookdev
    
    
    'Find your today's appointments
    strFilter = Format(Now, "ddddd")
    'strFilter = "2019-03-07"
    strFilter = "[Start] > '" & strFilter & " 00:00 AM' AND [Start] <= '" & strFilter & " 11:59 PM'"
    Set objTodayAppointments = objAppointments.Restrict(strFilter)
    
    For Each objAppointment In objTodayAppointments
        Call CheckReminder(objAppointment)
    Next
    ' MsgBox "Meeting reminders were checked!"
    
End Sub

Sub CheckCurrentDayReminders()
    ' Check Reminder for selected Day in Calendar View
    
    
    ' https://www.datanumen.com/blogs/quickly-send-todays-appointments-someone-via-outlook-vba/
    Dim objAppointments As Outlook.Items
    Dim objTodayAppointments As Outlook.Items
    Dim strFilter As String
    Dim objAppointment As Outlook.AppointmentItem ' Object
    
    Set objAppointments = Application.Session.GetDefaultFolder(olFolderCalendar).Items
    objAppointments.IncludeRecurrences = True
    objAppointments.Sort "[Start]", False        ' Bug: use False/descending see https://social.msdn.microsoft.com/Forums/office/en-US/919e1aee-ae67-488f-9adc-2c8518854b2a/how-to-get-recurring-appointment-current-date?forum=outlookdev
    
    
    Dim objCurAppointment As Object              ' Object
    Set objCurAppointment = GetCurrentItem()
    
    If (objCurAppointment Is Nothing) Then
        strFilter = Format(Now, "ddddd")
    ElseIf Not TypeOf objCurAppointment Is Outlook.AppointmentItem Then
        strFilter = Format(Now, "ddddd")
    Else
        strFilter = Format(objCurAppointment.Start, "ddddd")
    End If
    
    
    'Find your today's appointments
    
    'strFilter = "2019-03-07"
    strFilter = "[Start] > '" & strFilter & " 00:00 AM' AND [Start] <= '" & strFilter & " 11:59 PM'"
    Set objTodayAppointments = objAppointments.Restrict(strFilter)
    
    For Each objAppointment In objTodayAppointments
        Call CheckReminder(objAppointment)
    Next
    ' MsgBox "Meeting reminders were checked!"
    
End Sub

Sub CheckReminder(objAppointment As Outlook.AppointmentItem)
    
    
    Debug.Print "Check Reminder for '" & objAppointment.Subject & "'..."
    
    ' OUTLOOK BUG - set reminder on the serie if serie has some exceptions does not work
    'If objAppointment.IsRecurring Then
    '    Set objAppointment = objAppointment.Parent
    'End If
    
    
    If objAppointment.ReminderSet = False Then
        ' Exclude Meetings mark as Free
        If Not (objAppointment.MeetingStatus = olNonMeeting) And (objAppointment.BusyStatus = olFree) Then
            Exit Sub
        End If
        objAppointment.ReminderSet = True
        objAppointment.ReminderMinutesBeforeStart = 15 ' Enter your default time
        objAppointment.Save
        Debug.Print "Reminder set for '" & objAppointment.Subject & "'."
    End If
    
    
End Sub

Sub CheckReminders()
    Dim objItem As Object
    
    Set coll = GetCurrentItems
    If coll.Count = 0 Then
        Exit Sub
    End If
    
    For Each objItem In coll
        If TypeOf objItem Is Outlook.AppointmentItem Then
            Call CheckReminder(objItem)
        End If
    Next
End Sub

Public Sub SetDefaultReminder()
Dim objItem As Object

Set coll = GetCurrentItems
If coll.Count = 0 Then
    Exit Sub
End If

For Each objItem In coll
    If TypeOf objItem Is Outlook.AppointmentItem Then
        If objItem.ReminderSet = False Then
            objItem.ReminderSet = True
            objItem.ReminderMinutesBeforeStart = 15 ' Enter your default time
            objItem.Save
        End If
    End If
Next


End Sub

Sub ForwardMeetingInvitation()
    ' https://www.datanumen.com/blogs/3-methods-forward-meeting-invitation-without-notifying-organizer/
    Dim olSel As Selection
    Dim olMeeting As AppointmentItem
    Dim olFYIMeeting As MeetingItem
    Dim olMail As MailItem
    
    Dim Recip As String
    
    Set olSel = Outlook.Application.ActiveExplorer.Selection
    Set olMeeting = olSel.Item(1)
    'Replace with your own desired recipient's email address
    Recip = "johnsmith@datanumen.com"
    
    Set olMail = olMeeting.Reply
    With olFYIMeeting
        .Recipients.Add (Recip)
        .Attachments.Add olMeeting
        .Display
    End With
    
    Set olSel = Nothing
    Set olMeeting = Nothing
    Set olFYIMeeting = Nothing
End Sub

Sub Duplicate()
    Dim Item As Object
    
    Set Item = GetCurrentItem()
    If Item Is Nothing Then
        MsgBox "No Item selected"
        Exit Sub
    End If
    If Not TypeOf Item Is Outlook.AppointmentItem Then
        Exit Sub
    End If
    
    
    Dim myCopiedItem As Outlook.AppointmentItem
    Set myCopiedItem = Item.Copy
    
    ' TODO Does not work - not method to set property RecurrencePattern
    If Item.IsRecurring Then
        Dim RecPat As RecurrencePattern
        Set RecPat = myCopiedItem.GetRecurrencePattern
        Set srcRecPat = Item.GetRecurrencePattern
        Set RecPat = srcRecPat
    End If
    
    myCopiedItem.Display
    ' TODO if user delete the item, it closes the window but does not delete it
    
End Sub


Sub FYIAppointment()
    ' Send copy of meeting with free availability/ no reminder
    ' Body of appointment is copied from Reply by Email body
    Dim olAppointment As Object
    Dim olMail As MailItem
    
    Set olAppointment = GetCurrentItem()
    If olAppointment Is Nothing Then
        MsgBox "No Item selected"
        Exit Sub
    End If
    If Not TypeOf olAppointment Is Outlook.AppointmentItem Then
        Exit Sub
    End If
    
    ' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/e2fb4e2d-9e27-452e-a50a-7e5c0dac1af5/cant-get-vba-code-to-use-replyall-on-existing-meetings?forum=outlookdev
    olAppointment.Display
    olAppointment.GetInspector.CommandBars.ExecuteMso ("Reply")
    olAppointment.Close (False)
    ' Get Created Email
    Set olMail = Application.ActiveInspector.CurrentItem
    
    
    'Dim olMeeting As MeetingItem
    Dim olFYIMeeting As AppointmentItem
    
    Set olFYIMeeting = Application.CreateItem(olAppointmentItem)
    
    'Replace with your own desired recipient's email address
    Dim Recip As String
    Recip = "johnsmith@datanumen.com"
    
    
    For i = olFYIMeeting.Recipients.Count To 1 Step -1
        olFYIMeeting.Recipients.Remove (i)
    Next i
    
    With olFYIMeeting
        
        '.Recipients.Add (Recip)
        '.Attachments.Add olMeeting
        .Body = olMail.Body
        .Subject = "FYI: " & olAppointment.Subject
        .BusyStatus = olFree
        .ReminderSet = False
        .Start = olAppointment.Start
        .Duration = olAppointment.Duration
        .Location = olAppointment.Location
        .Save
    End With
    olMail.Close (olDiscard)
    olFYIMeeting.Display
    olFYIMeeting.GetInspector.CommandBars.ExecuteMso ("InviteAttendees")
    
End Sub

Function AcceptMeeting(oItem As Object)
    ' Accept meeting and Delete meeting request
    ' Check if meeting collision based on https://www.slipstick.com/outlook/calendar/autoaccept-a-meeting-request-using-rules/
    
    Dim myAppt As Outlook.AppointmentItem
    Dim myMtg As Outlook.MeetingItem
    
    
    If TypeName(myMtgReq) = "Nothing" Then
        MsgBox "No meeting request selected!"
        Exit Function
    End If
    
    ' From Calendar view
    If TypeOf oItem Is AppointmentItem Then
        Set myAppt = oItem
    ElseIf TypeOf oItem Is MeetingItem Then
        ' From Inbox View
        Set myAppt = oItem.GetAssociatedAppointment(True) ' Add to Calendar
    Else
        MsgBox "No meeting request selected!"
        Exit Function
    End If
    
    '    ' Check if Meeting Conflict
    '    If Not IsFree(myAppt) Then
    '         answer = MsgBox("This meeting conflicts with another one." & vbCrLf & "Are you sure you want to accept this meeting?", vbYesNo + vbQuestion, "Meeting Collision")
    '         If answer = vbNo Then
    '            Exit Sub
    '        End If
    '    End If
    
    Set myMtg = myAppt.Respond(olMeetingAccepted, True)
    myMtg.Send
    
    ' Inbox View
    If TypeOf oItem Is MeetingItem Then
        ' Delete Email Request
        oItem.Delete
    End If
    
End Function


Public Sub MarkFree_Callback()
' Answer selected Meeting Items in the Inbox with Tentative, Delete Request
Dim objItem As Object
Dim objAppt As AppointmentItem

Set coll = GetCurrentItems()

For Each objItem In coll
    
    ' From Calendar view
    If TypeOf objItem Is AppointmentItem Then
        Call MarkFree(objItem)
    ElseIf TypeOf objItem Is MeetingItem Then
        ' From Inbox View
        Set objAppt = objItem.GetAssociatedAppointment(True) ' Add to Calendar
        Call MarkFree(objAppt)
        
        Set myRsp = objAppt.Respond(olMeetingAccepted, True)
        objItem.Delete
    End If
    
Next


End Sub
Function MarkFree(ApptItem As AppointmentItem)
    
    ' Remove Reminder
    ApptItem.ReminderSet = False
    ' Set Free
    ApptItem.BusyStatus = olFree
    ApptItem.Save
    
End Function

Function TentativeNoResponse(oItem As MeetingItem)
Dim myAppt As Outlook.AppointmentItem
Dim myMtgReq, myMtg As Outlook.MeetingItem
    
Set myAppt = oItem.GetAssociatedAppointment(True)
    
    
    Set myMtg = myAppt.Respond(olMeetingTentative, True)
    'myMtg.Send
    
    oItem.Delete
End Function


Public Sub TententativeNoResponse_Callback()
' Answer selected Meeting Items in the Inbox with Tentative, Delete Request
Dim objItem As Object

For Each objItem In coll
    
    ' From Calendar view
    If TypeOf objItem Is AppointmentItem Then
        Set myMtg = objItem.Respond(olMeetingTentative, True)
        
    ElseIf TypeOf objItem Is MeetingItem Then
        ' From Inbox View
        Set objAppt = objItem.GetAssociatedAppointment(True) ' Add to Calendar
        Set myMtg = objAppt.Respond(olMeetingTentative, True)
        objItem.Delete
    End If
    
Next

End Sub



' https://www.slipstick.com/outlook/calendar/autoaccept-a-meeting-request-using-rules/
Function IsFree(oAppt As AppointmentItem)
    
    
    Dim myAcct As Outlook.Recipient
    Dim myFB As String
    
    Set myAcct = Session.CreateRecipient("Thierry.Dalon@continental-corporation.com")
    
    ' Check for working time
    meetingtime = Format(oAppt.Start, "h:mm:ss AM/PM")
    
    If meetingtime < #8:00:00 AM# Or meetingtime > #4:00:00 PM# Then
        IsFree = False
        Debug.Print "Out of working time."
        Exit Function
    End If
    
    
    Debug.Print oAppt.Duration
    Debug.Print "Time: " & TimeValue(oAppt.Start)
    Debug.Print "periods before appt: " & TimeValue(oAppt.Start) * 288
    Debug.Print oAppt.Start
    
    myFB = myAcct.FreeBusy(oAppt.Start, 5, True)
    
    Debug.Print myFB
    
    Dim oResponse
    Dim i As Long
    Dim Test As String
    
    i = (TimeValue(oAppt.Start) * 288)
    Test = Mid(myFB, i + 1, (oAppt.Duration / 5)) ' Allow adjacent meetings
    
    Debug.Print "String to check: " & Test
    
    ' O : Free or Working Elsewhere, 1: Tentative, 2:Busy, 3: OOO
    If InStr(1, Test, "2") Or InStr(1, Test, "3") Then
        ' Consider Tentative as Free
        IsFree = False
    Else
        IsFree = True
    End If
End Function

Sub DeclineWithCopy()
    ' Decline Meetings and Save a Copy
    '
    ' http://www.vbaexpress.com/forum/showthread.php?40289-Solved-Keeping-a-declined-meeting-in-calendar
    Dim cAppt, oAppt As AppointmentItem
    Dim oResponse As Outlook.MeetingItem
    Set oAppt = GetCurrentItem
    Set cAppt = oAppt.Copy
    cAppt.Subject = "DECLINED: " & oAppt.Subject
    cAppt.BusyStatus = olFree
    cAppt.ReminderSet = False                    ' Remove Reminder
    'cAppt.Categories = "Declined"
    cAppt.Save
    
    Set oResponse = oAppt.Respond(olMeetingDeclined, False, False) ' Decline with Prompt for explication/ user send the response
    Set cAppt = Nothing
    Set oAppt = Nothing
    
End Sub

Sub DeclineAndSave()
    ' Decline Meetings but keep a ghost meeting as free in calendar
    ' Calls: Utils/GetCurrentItem
    '
    
    'https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff867189(v=office.14)
    Dim myAppt As Outlook.AppointmentItem
    Dim myMtgReq, myMtg As Outlook.MeetingItem
    Dim curItem As Object
    
    Set curItem = GetCurrentItem
    
    ' From Event
    If TypeOf curItem Is AppointmentItem Then
        Set myAppt = curItem
        ' From Inbox
    ElseIf TypeOf curItem Is MeetingItem Then
        Set myMtgReq = curItem
        Set myAppt = myMtgReq.GetAssociatedAppointment(True) ' Add to Calendar
    Else
        'If TypeName(myMtgReq) = "Nothing" Then
        MsgBox "Selection Error!"
        Exit Sub
    End If
    
    Set myMtg = myAppt.Respond(olMeetingDeclined, False, False) ' Decline with Prompt for explication/ user send the response
    'myMtg.Send
    
    ' Send Tentative Response
    Set myMtg = myAppt.Respond(olMeetingTentative, True) ' Send Tentative Response without asking
    
    
    ' Rename Title with prefixed Declined:
    myAppt.Subject = "DECLINED: " & myAppt.Subject
    'myAppt.Categories = "Declined"
    ' Set Availability to Free
    myAppt.BusyStatus = olFree
    myAppt.Save
    
    ' Delete original request
    On Error Resume Next                         ' From Calendar view
    myMtgReq.Delete
    
    
End Sub

Sub Check_Category()
    ' Check if category 'declined' exists. If not, create it and add a suitable color
    Dim objNamespace
    Dim objCategory
    
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objCategory = objNamespace.Categories.Item("Declined")
    If objCategory Is Nothing Then
        Set objCategory = objNamespace.Categories.Add("Declined") ', OlCategoryColor.olCategoryColorTeal)
    End If
    objCategory = Nothing
    objNamespace = Nothing
End Sub

Sub MakeDefAppointment(oAppt As AppointmentItem)
    
    ' Only for Default Calendar
    If Not (oAppt.Parent.FolderPath = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).FolderPath) Then
        Exit Sub
    End If
    
    ' olkApt.MeetingStatus = olMeeting ' Convert to meeting
    
    ' Set Skype Meeting using QAT and SendKey
    SendKeys "%", True                           ' Alt+1 for QAT first position
    SendKeys "1", True
    
    
    
    'TODO Update Dial-in numbers Part below does not work
    Exit Sub
    
    ' Wait for Skype Meeting being created
    Dim TimeInterval As Single
    TimeInterval = Timer
    Delay = 5
    Do Until Timer >= TimeInterval + Delay       'SET DELAY TO AMOUNT OF SECONDS TO TRY
        DoEvents
        If oAppt.Location <> "" Then Exit Do
    Loop
    
    If Timer >= TimeInterval + Delay Then
        'Timer exceeded
        Exit Sub
    End If
    
    ' Need to refresh after Skype/Teams Meeting creation
    Set oAppt = GetCurrentItem
    
    ' Customize Skype Invitation
    EditSkype oAppt
    
    
End Sub

' See
Sub EditPhone_Callback()
    Dim oAppt As AppointmentItem
    Set oAppt = GetCurrentItem
    
    If InStr(oAppt.Location, "Skype Meeting") Then
        Call EditSkype(oAppt)
    ElseIf InStr(oAppt.Location, "Microsoft Teams") Then
        Call EditTeams(oAppt)
    Else
        Exit Sub
    End If
    
    
End Sub

Sub EditSkype(oAppt As AppointmentItem, Optional wordFile As String = "")
    ' Replace numbers
    ' https://www.slipstick.com/outlook/customize-skype-business-invitation/
    
    
    ' Input from Word document - Edit path to your location
    If wordFile = "" Then
        wordFile = Const_DefSkypeDialIn
    End If
    If Not IsFile(wordFile) Then
        
        wordFile = FileOpen(wordFile & "\Skype_*.docx", "Dial-In Templates", "*.docx")
        
        If wordFile = "" Then
            MsgBox "No file selected.", vbExclamation, "Exit"
            Exit Sub
        End If
        
    End If
    
    
    ' Open file and copy content
    Dim srcDoc As Word.Document
    Set srcDoc = Word.Documents.Open(Filename:=wordFile, Visible:=False)
    srcDoc.Activate
    srcDoc.Range(0, 0).Select
    Word.Selection.WholeStory
    Word.Selection.Copy
    
    
    Dim objDoc As Word.Document
    Set objDoc = oAppt.GetInspector.WordEditor   ' Strange Error
    Dim oRng As Word.Range
    Set oRng = objDoc.Range
    'objDoc.Activate
    
    
    With oRng.Find
        .Text = "Join by phone*.{137}"
        .MatchWildcards = True
        .Execute
        If .Found Then
            oRng.Paste
        End If
    End With
    
    ' Close opened document without saving
    srcDoc.Close
    
End Sub

Sub EditTeams(oAppt As AppointmentItem, Optional wordFile As String = "")
    ' Replace numbers
    ' https://www.slipstick.com/outlook/customize-skype-business-invitation/
    
    
    ' Input from Word document - Edit path to your location
    If wordFile = "" Then
        wordFile = Const_DefTeamsDialIn
    End If
    If Not IsFile(wordFile) Then
        
        wordFile = FileOpen(wordFile & "\Teams_*.docx", "Dial-In Templates", "*.docx")
        
        
        If wordFile = "" Then
            MsgBox "No file selected.", vbExclamation, "Exit"
            Exit Sub
        End If
        
    End If
    
    
    Dim srcDoc As Word.Document
    Set srcDoc = Word.Documents.Open(Filename:=wordFile, Visible:=False)
    srcDoc.Activate
    srcDoc.Range(0, 0).Select
    Word.Selection.WholeStory
    Word.Selection.Copy
    
    
    Dim objDoc As Word.Document
    Set objDoc = oAppt.GetInspector.WordEditor   ' Strange Error
    Dim oRng As Word.Range
    Set oRng = objDoc.Range
    'objDoc.Activate
    
    
    With oRng.Find
        .Text = "+*_{80}"
        .MatchWildcards = True
        .Execute
        If .Found Then
            oRng.Paste
        End If
    End With
    
    ' Close opened document without saving
    srcDoc.Close
    
End Sub




' https://www.datanumen.com/blogs/2-methods-export-specific-date-range-outlook-calendar-icalendar-ics-file/
Sub ExportCalendarToIcs()
    Dim objCalendarFolder As Outlook.Folder
    Dim objCalendarExporter As Outlook.CalendarSharing
    Dim dStartDate As Date
    Dim dEndDate As Date
    Dim objShell As Object
    Dim objSavingFolder As Object
    Dim strSavingFolder As String
    Dim striCalendarFile As String
    
    'Get the current Calendar folder
    Set objCalendarFolder = Outlook.Application.ActiveExplorer.CurrentFolder
    
    
    If objCalendarFolder Is Nothing Or objCalendarFolder.DefaultItemType <> olAppointmentItem Then
        Dim Ns As Outlook.NameSpace
        Set Ns = Application.GetNamespace("MAPI")
        
        'use the default folder
        Set objCalendarFolder = Ns.GetDefaultFolder(olFolderCalendar)
    End If
    
    Set objCalendarExporter = objCalendarFolder.GetCalendarExporter
    
    'Enter the specific start date and end date
    dStartDate = InputBox("Enter the start date:", "Specify Start Date")
    dEndDate = InputBox("Enter the end date:", "Specify End Date")
    
    If dStartDate <> #1/1/4501# And dEndDate <> #1/1/4501# Then
        'Select a Windows folder for saving the exported iCalendar file
        Set objShell = CreateObject("Shell.Application")
        Set objSavingFolder = objShell.BrowseForFolder(0, "Select a folder:", 0, "")
        strSavingFolder = objSavingFolder.self.Path
        
        If strSavingFolder <> "" Then
            strCalendarFile = strSavingFolder & "\" & "Calendar from " & Format(dStartDate, "YYYY-MM-DD") & " to " & Format(dEndDate, "YYYY-MM-DD") & ".ics"
            
            'Export the calendar in specific date range
            With objCalendarExporter
                .IncludeWholeCalendar = False
                .StartDate = dStartDate
                .EndDate = dEndDate
                .CalendarDetail = olFullDetails
                .IncludeAttachments = False
                .IncludePrivateDetails = False
                .RestrictToWorkingHours = False
                .SaveAsICal strCalendarFile
            End With
            
            MsgBox "Calendar Exported Successfully to " & strCalendarFile & "!", vbInformation
        End If
        
    End If
    'Else
    '       MsgBox "Open a calendar folder, please!", vbExclamation + vbOKOnly
    'End If
End Sub



