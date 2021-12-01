Attribute VB_Name = "CheckReminder"
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
' Check Reminder for selected Items (GetCurrentItems)
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

