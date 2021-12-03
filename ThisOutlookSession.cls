VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1

Private Sub Application_Startup()
  Dim Ns As Outlook.NameSpace

  Set Ns = Application.GetNamespace("MAPI")
  Set Items = Ns.GetDefaultFolder(olFolderCalendar).Items
End Sub

' Ask if Remove Reminders on AllDay Events
' If no Reminder, Ask if Set default for meetings
Private Sub Items_ItemAdd(ByVal Item As Object)
'On Error Resume Next


If TypeOf Item Is Outlook.AppointmentItem Then

    'Checks to see if all day and if it has a reminder set to true
    If Item.AllDayEvent = True And Item.ReminderSet = True Then
        If MsgBox("Do you want to remove the reminder of the AllDayEvent?", vbYesNo) = vbNo Then
          Exit Sub
        End If
        Item.ReminderSet = False
        Item.Save

    End If ' AllDayEvent
    
    If Item.ReminderSet = False And Item.MeetingStatus = olMeeting And DateDiff("n", Now, Item.Start) > 0 Then
        If MsgBox("No Reminder for future meeting:Do you want to set a reminder (15min)?", vbYesNo) = vbNo Then
          Exit Sub
        End If
        With Item
            .ReminderSet = True
            .ReminderMinutesBeforeStart = 15
            .Save
        End With
    End If ' Meeting without reminder in the future

    
End If
End Sub





