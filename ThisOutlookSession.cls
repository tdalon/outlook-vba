VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents calItems As Outlook.Items
Attribute calItems.VB_VarHelpID = -1
Private WithEvents objIncomingItems As Outlook.Items
Attribute objIncomingItems.VB_VarHelpID = -1



Private Sub Application_Startup()
  Set calItems = Session.GetDefaultFolder(olFolderCalendar).Items
  Set objIncomingItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
  Call Cat_InitDic
End Sub

' Categorize Sent Emails
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    If TypeOf Item Is Outlook.MailItem And Len(Item.Categories) = 0 Then
        Call Cat_CheckRecip(Item)
    End If
End Sub


Private Sub calItems_ItemAdd(ByVal Item As Object)
'On Error Resume Next

' Ask if Remove Reminders on AllDay Events
' If no Reminder, Ask if Set default for meetings
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

' Categorize based on Recipients Domain
If TypeOf Item Is Outlook.AppointmentItem Then
    Call Cat_CheckRecip(Item)
End If

End Sub



Private Sub objIncomingItems_ItemAdd(ByVal oItem As Object)
    
Call Cat_CheckItem(oItem)
    
End Sub
