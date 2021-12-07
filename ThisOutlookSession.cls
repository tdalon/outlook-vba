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
Public WithEvents objIncomingItems As Outlook.Items
Attribute objIncomingItems.VB_VarHelpID = -1

Dim oDomain2CategoryDic


Private Sub Application_Startup()
  Set calItems = Session.GetDefaultFolder(olFolderCalendar).Items
  Set objIncomingItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
  
    ' Init Variable
    Set oDomain2CategoryDic = CreateObject("Scripting.Dictionary")
    oDomain2CategoryDic.Add "@customer1.com", "4Customer1"     'Add some keys and items
    oDomain2CategoryDic.Add "@customer2.com", "4Customer2"
End Sub

' Categorize Sent Emails
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    If TypeOf Item Is Outlook.MailItem And Len(Item.Categories) = 0 Then
        ' Check recipients
        Dim recs As Outlook.Recipients
        Dim rec As Outlook.Recipient
        
        Set recs = oItem.Recipients
    
       For i = recs.Count To 1 Step -1
           For Each varKey In oDomain2CategoryDic.Keys()
               Set rec = recs.Item(i)
               If (InStr(rec.Address, varKey) > 0) Then
                   oItem.Categories = oDomain2CategoryDic(varKey)
                   oItem.Save
                   Exit Sub
               End If
           Next
       Next
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

    ' Check recipients
     Dim recs As Outlook.Recipients
     Dim rec As Outlook.Recipient
     
     Set recs = Item.Recipients
 
    For i = recs.Count To 1 Step -1
        For Each varKey In oDomain2CategoryDic.Keys()
            Set rec = recs.Item(i)
            ' Debug.Print rec.Address
            
            If (InStr(rec.Address, varKey) > 0) Then
                Item.Categories = oDomain2CategoryDic(varKey)
                Item.Save
                Exit Sub
            End If
        Next
    Next

End If

End Sub





Private Sub objIncomingItems_ItemAdd(ByVal oItem As Object)
    'Dim objMeetingRequest As Outlook.MeetingItem
    'Dim objMeeting As Outlook.AppointmentItem
    
 
    If TypeOf oItem Is MeetingItem Then
       Set objMeetingRequest = oItem
       Set oItem = objMeetingRequest.GetAssociatedAppointment(True)
    End If ' MeetingRequest
  
    ' Check FromEmailAddress
    Dim SenderEmailAddress As String
    SenderEmailAddress = GetFromEmail(oItem)
    For Each varKey In oDomain2CategoryDic.Keys()
     
      If (InStr(LCase(SenderEmailAddress), varKey) > 0) Then
          oItem.Categories = oDomain2CategoryDic(varKey)
              oItem.Save
              Exit Sub
     End If
     
     Next
       
       
     ' Check recipients
     Dim recs As Outlook.Recipients
     Dim rec As Outlook.Recipient
     
     Set recs = oItem.Recipients
      
 
    For i = recs.Count To 1 Step -1
        For Each varKey In oDomain2CategoryDic.Keys()
        
            Set rec = recs.Item(i)
            If (InStr(rec.Address, varKey) > 0) Then
                oItem.Categories = oDomain2CategoryDic(varKey)
                oItem.Save
                Exit Sub
            End If
        
        Next
 
    Next
    
    ' Check Subject
    For Each varKey In oDomain2CategoryDic.Keys()
        
        If (InStr(oItem.Subject, Replace(oDic(varKey), "4", ""), vbTextCompare) > 0) Then ' vbTextCompare: Case insensitive
            oItem.Categories = oDic(varKey)
            oItem.Save
            Exit Sub
        End If
    
    Next
       
    
End Sub









