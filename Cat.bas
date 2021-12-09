Attribute VB_Name = "Cat"
Private oDomain2CategoryDic ' Global variable only for the module

Sub AutoMarkMeetingPrivate(objMeetingRequest As MeetingItem)
    Dim objMeeting As Outlook.AppointmentItem
 
    If TypeOf objMeetingRequest Is MeetingItem Then
       Set objMeeting = objMeetingRequest.GetAssociatedAppointment(True)
       objMeeting.Sensitivity = olPrivate
       objMeeting.Save
    End If
End Sub

Sub Cat_InitDic()
' Init Variable
    Set oDomain2CategoryDic = CreateObject("Scripting.Dictionary")
    oDomain2CategoryDic.Add "@customer1.com", "4Customer1"     'Add some domain <-> Category Mapping
    oDomain2CategoryDic.Add "@customer2.de", "4Customer2"
End Sub


' ---------------------------
Sub Cat_CheckItem(oItem As Object)
If TypeOf oItem Is MeetingItem Then
Dim objMeetingRequest As Object
   Set objMeetingRequest = oItem.GetAssociatedAppointment(True)
   Call Cat_CheckItem(objMeetingRequest)
End If ' MeetingRequest: categorize Appointment and MeetingRequest


' Check FromEmailAddress
Dim SenderEmailAddress As String
SenderEmailAddress = GetFromEmail(oItem)
Dim sCat As String

If IsEmpty(oDomain2CategoryDic) Then
    Call Cat_InitDic
End If
For Each varKey In oDomain2CategoryDic.Keys()
    If (InStr(LCase(SenderEmailAddress), varKey) > 0) Then
        sCat = oDomain2CategoryDic(varKey)
        Call SetCategory(oItem, sCat)
        'oItem.Categories = oDomain2CategoryDic(varKey)
        'oItem.Save
        Exit Sub
    End If
 
Next
   
 ' Check Recipients
If Cat_CheckRecip(oItem) Then
    Exit Sub
End If

' Check Subject
For Each varKey In oDomain2CategoryDic.Keys()
    sCat = oDomain2CategoryDic(varKey)
    If (InStr(1, oItem.Subject, Replace(sCat, "4", ""), vbTextCompare) > 0) Then ' vbTextCompare: Case insensitive
        Call SetCategory(oItem, sCat)
        Exit Sub
    End If
Next
End Sub

Function Cat_CheckRecip(oItem As Object) As Boolean

' Check recipients
Dim recs As Outlook.Recipients
Dim rec As Outlook.Recipient
Dim sCat As String

Set recs = oItem.Recipients

For i = recs.Count To 1 Step -1
    For Each varKey In oDomain2CategoryDic.Keys()
        Set rec = recs.Item(i)
        If (InStr(rec.Address, varKey) > 0) Then
            sCat = oDomain2CategoryDic(varKey)
            Call SetCategory(oItem, sCat)
            Cat_CheckRecip = True
            Exit Function
        End If
    Next
Next
Cat_CheckRecip = False
End Function

' ---------------------------
Sub Cat_CheckCurrentItems()
Set coll = GetCurrentItems
If coll.Count = 0 Then
    Exit Sub
End If
Dim objItem As Object
For Each objItem In coll
    Call Cat_CheckItem(objItem)
Next

End Sub

Sub SetCategory(oItem As Object, sCat As String)

If TypeOf oItem Is AppointmentItem Then
    If oItem.IsRecurring Then
        If oItem.RecurrenceState <> olApptMaster Then
            Set oItem = oItem.Parent
        End If
    End If
    
' ElseIf TypeOf oItem Is MailItem Then

End If
oItem.Categories = sCat
oItem.Save
End Sub


