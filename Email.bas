Attribute VB_Name = "Email"
' http://www.vboffice.net/en/developers/flag-email-for-follow-up/
Public Sub MarkItemAsTask(ByVal AddDays As Long, _
                          Optional TimeOfDay As String = "08:00", _
                          Optional Subject As String, _
                          Optional Mail As Outlook.MailItem _
                          )
    Dim Items As VBA.Collection
    Dim obj As Object
    Dim i As Long
    Dim dt As Date
    Dim tm As String
    Dim Icon As OlMarkInterval

    dt = DateAdd("d", AddDays, Date)
    tm = CStr(dt) & " " & TimeOfDay

    If AddDays < 1 Then
        Icon = olMarkToday
    ElseIf AddDays = 1 Then
        Icon = olMarkTomorrow
    ElseIf Weekday(Date, vbUseSystemDayOfWeek) + AddDays < 8 Then
        Icon = olMarkThisWeek
    Else
        Icon = olMarkNextWeek
    End If

    If Mail Is Nothing Then
        Set Items = GetCurrentItems
    Else
        Set Items = New VBA.Collection
        Items.Add Mail
    End If

    For Each obj In Items
        If TypeOf obj Is Outlook.MailItem Then
            Set Mail = obj
            Mail.MarkAsTask Icon
            Mail.TaskStartDate = tm
            Mail.TaskDueDate = tm
            If Len(Subject) Then
                Mail.TaskSubject = Subject
                Mail.FlagRequest = Subject
            End If
            Mail.ReminderTime = tm
            Mail.ReminderSet = True
            Mail.Save
        End If
    Next
End Sub

Sub SendAndFile()
    
    Dim Item As Outlook.MailItem
    Set Item = Application.ActiveInspector.CurrentItem
    Dim objNS As NameSpace
    Dim objFolder As MAPIFolder
    Set objNS = Application.Session
    Set objFolder = objNS.PickFolder
    If objFolder Is Nothing Then
        Exit Sub
    End If
    Set Item.SaveSentMessageFolder = objFolder
    Item.Send
    
End Sub

' https://www.slipstick.com/developer/code-samples/create-outlook-appointment-from-message/
'Private Sub CreateLogCalendar()
Public Sub CopyToTasksCalendar()
' Calls GetCurrentItem

    Dim objAppt As Outlook.AppointmentItem
    Dim Item As Object                           ' works with any outlook item

    ' OPTIONS
    Dim bAskAttach As Boolean
    bAskAttach = False                           ' Change to True if you want to be asked to attach. Preferred: False and keep link

    Dim bAskDelete As Boolean
    bAskDelete = False                           ' Change to True if you want to be asked if you want to delete original Item. Preferred: False alsways keep and use a link


    Set Item = GetCurrentItem()

    Set CalFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Folders("Tasks")
    Set objAppt = CalFolder.Items.Add(olAppointmentItem)
    With objAppt
        .Subject = Item.Subject
        '.Body = Item.Body
        .Start = Now
        '.End = Now
        '.ReminderSet = False
        '.BusyStatus = olFree ' Not needed in separate Tasks calendar
    
    End With


    If bAskAttach Then
        If MsgBox("Do you want to attach original item?", vbYesNo + vbQuestion) = vbYes Then
            objAppt.Attachments.Add Item
            objAppt.Body = Item.Body
            If MsgBox("Do you want to delete original item?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                Item.Delete
            End If
            objAppt.Display
        End If
    Else
        If Item.Attachments.Count > 0 Then
            If MsgBox("Do you want to copy attachments from original item to the Task?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                Call CopyAttachments(Item, objAppt)
            End If
        End If
    
        If bAskDelete Then
            If MsgBox("Do you want to delete original item?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                objAppt.Body = Item.Body
                Item.Delete
            End If
        Else
            ' Flag email
            If TypeOf Item Is Outlook.MailItem Then
                Item.MarkAsTask olMarkNoDate
                Item.FlagRequest = "Follow up in Calendar"
            End If
        
            ' Add Link to Email
            ' Create dummy email
            Dim olMail As Outlook.MailItem
            Set olMail = Outlook.CreateItem(olMailItem)
            olMail.Body = Item.Body
        
            sLink = "outlook:" & Item.EntryID
            sText = Item.Subject & " (" + Item.SenderName & ")"
            sHtml = "<a href=" & sLink & ">" & sText & "</a>"
            olMail.HTMLBody = sHtml & "<br>" & Item.HTMLBody
            olMail.Display                       'Required else change is not copied
            Sleep (500)
            ' Copy Body with Formatting : requires copy to Email then paste into Appointment
            Set objInsp = olMail.GetInspector
            If objInsp.EditorType = olEditorWord Then
                Set objDoc = objInsp.WordEditor
                Set objWord = objDoc.Application
                Set objSel = objWord.Selection
                With objSel
                    .WholeStory
                    .Copy
                End With
            End If
        
            ' Paste to Appointment with formatting
            'objAppt.Subject = objAppt.Subject & vbCrLf & sLink
            objAppt.Display 'show to add notes ' required at the beginning - else error at paste. objSel broken
            'Sleep (500)
            
            Set objInsp = objAppt.GetInspector
            Set objDoc = objInsp.WordEditor
            Set objSel = objDoc.Windows(1).Selection
            
            objSel.PasteAndFormat (wdFormatOriginalFormatting)
            olMail.Close (olDiscard)
        
        End If
    End If

End Sub

' https://www.slipstick.com/developer/code-samples/create-tasks-task-folders/
Public Sub CopyToTask()

    Dim olTask As Outlook.TaskItem
    Dim Item As Object                           ' works with any outlook item

    Set Item = GetCurrentItem()

    Set taskFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderTasks) '.Folders(tFolder)
    Set olTask = taskFolder.Items.Add(olTaskItem)
    With olTask
        .Subject = Item.Subject
    
        '.End = Now
        '.ReminderSet = False
        '.BusyStatus = olFree
    
    End With

    If MsgBox("Do you want to attach original item?", vbYesNo + vbQuestion) = vbYes Then
        olTask.Attachments.Add Item
        If MsgBox("Do you want to copy attachments from original item to the Task?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
            Item.Delete
        End If
        olTask.Body = Item.Body
        olTask.Display                           'show to add notes
    Else
    
        If Item.Attachments.Count > 0 Then
            If MsgBox("Do you want to copy attachments from original item to the Task?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                Call CopyAttachments(Item, olTask)
            End If
        End If
    
    
        If MsgBox("Do you want to delete original item?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Item.Delete
        Else
            ' Flag email
            If TypeOf Item Is Outlook.MailItem Then
                Item.MarkAsTask olMarkNoDate
                Item.FlagRequest = "Follow-up in Tasks"
            End If
        
            ' Add Link to Email
            Dim olMail As Outlook.MailItem
            Set olMail = Outlook.CreateItem(olMailItem)
            olMail.Body = Item.Body
        
            sLink = "outlook:" & Item.EntryID
            sText = Item.Subject & " (" + Item.SenderName & ")"
            sHtml = "<a href=" & sLink & ">" & sText & "</a>"
            olMail.HTMLBody = sHtml & "<br>" & Item.HTMLBody
            olMail.Display                       'Required else change is not copied
        
            ' Copy Body with Formatting
            Set objInsp = olMail.GetInspector
            If objInsp.EditorType = olEditorWord Then
                Set objDoc = objInsp.WordEditor
                Set objWord = objDoc.Application
                Set objSel = objWord.Selection
                With objSel
                    .WholeStory
                    .Copy
                End With
            End If
        
            ' Paste to Appointment with formatting
            'olAppt.Subject = olAppt.Subject & vbCrLf & sLink
            Set objInsp = olTask.GetInspector
            Set objDoc = objInsp.WordEditor
            Set objSel = objDoc.Windows(1).Selection
            olTask.Display                       'show to add notes ' required else error at paste
        
            objSel.PasteAndFormat (wdFormatOriginalFormatting)
        
            olMail.Delete
        End If
    End If

End Sub

Sub CopyAttachmentNames()
    
    Dim oItem As Object
    Set oItem = GetCurrentItem()
    Dim oAtt As Attachment
    Dim sHtml As String
    Dim sText As String
    
    sHtml = ""
    sText = ""
    For Each oAtt In oItem.Attachments
        sHtml = sHtml & "&lt;&lt;" & oAtt.Filename & "&gt;&gt; <br>"
        sText = sText & "<<" & oAtt.Filename & ">>" & vbCrLf
    Next oAtt
    
    Call CopyToClipboardHTML(sHtml, sText)
    
    
End Sub

' -----------------------------------------------------------
'  http://vboffice.net/en/developers/get-the-message-folder/
Public Sub GetItemsFolderPath()
    Dim obj As Object
    Dim F As Outlook.MAPIFolder
    Dim Msg$
    Set obj = Application.ActiveWindow
    If TypeOf obj Is Outlook.Inspector Then
        Set obj = obj.CurrentItem
    Else
        Set obj = obj.Selection(1)
    End If
    Set F = obj.Parent
    Msg = "The path is: " & F.FolderPath & vbCrLf
    Msg = Msg & "Switch to the folder?"
    If MsgBox(Msg, vbYesNo) = vbYes Then
        Set Application.ActiveExplorer.CurrentFolder = F
    End If
End Sub

Public Sub EditEmailSubject()
    Dim strSubject As String
    Dim objItem As Object

    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not TypeOf objItem Is Outlook.MailItem Then
        Exit Sub
    End If
    strSubject = objItem.Subject
    strSubject = InputBox("Edit Subject:", , strSubject)
    If strSubject = vbNullString Then
        ' User Cancelled
        Exit Sub
    End If
    objItem.Subject = strSubject
    objItem.Save
    Set objItem = Nothing
    '    Inspector.CommandBars.ExecuteMso ("EditMessage")
End Sub


