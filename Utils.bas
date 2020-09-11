Attribute VB_Name = "Utils"
Private Declare Function ShellExecute _
                         Lib "shell32.dll" Alias "ShellExecuteA" ( _
                         ByVal hWnd As Long, _
                         ByVal Operation As String, _
                         ByVal Filename As String, _
                         Optional ByVal Parameters As String, _
                         Optional ByVal Directory As String, _
                         Optional ByVal WindowStyle As Long = vbMinimizedFocus _
                         ) As Long



Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub CopyToClipboard(sText As String)
    ' Requires to add Microsoft Forms 2.0 as Reference see https://www.slipstick.com/developer/code-samples/paste-clipboard-contents-vba/

    Dim dObj As New DataObject
    dObj.SetText sText
    dObj.PutInClipboard


End Sub

Public Sub CopyToClipboardHTML(sHtml As String, Optional ByVal sText As String = "")
' Calls fclip.exe
    sTmpHtmlFile = Environ("temp") & "\Clipboard.html"
    file1 = FreeFile                             'Returns value of 1
    Open sTmpHtmlFile For Output As #file1
    Print #file1, sHtml
    Close #file1

    If sText = "" Then
        sText = sHtml
    End If

    sTmpTxtFile = Environ("temp") & "\Clipboard.txt"
    file1 = FreeFile                             'Returns value of 1
    Open sTmpTxtFile For Output As #file1
    Print #file1, sText
    Close #file1

    ' fclip.exe must be added to your System path (edit environment variable) or edit here with full path
    sCmd = "fclip.exe " & sTmpHtmlFile & " " & sTmpTxtFile

    RetVal = Shell(sCmd, vbMinimizedNoFocus)


End Sub

Function GetCurrentItem() As Object
    ' https://www.slipstick.com/developer/accept-or-decline-a-meeting-request-using-vba/
    Dim objApp As Outlook.Application
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
    Case "Explorer"
        Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
    Case "Inspector"
        Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    
    Set objApp = Nothing
End Function

Function GetCurrentItems() As VBA.Collection
    Dim c As VBA.Collection
    Dim Sel As Outlook.Selection
    Dim obj As Object
    Dim i&
    
    Set c = New VBA.Collection
    
    If TypeOf Application.ActiveWindow Is Outlook.Inspector Then
        c.Add Application.ActiveInspector.CurrentItem
    Else
        Set Sel = Application.ActiveExplorer.Selection
        If Not Sel Is Nothing Then
            For i = 1 To Sel.Count
                c.Add Sel(i)
            Next
        End If
    End If
    Set GetCurrentItems = c
End Function

' ----------------------------

' https://www.slipstick.com/developer/working-vba-nondefault-outlook-folders/
Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
    
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
    
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

Sub MoveDeleted()
    ' https://www.slipstick.com/developer/code-samples/move-deleted-items/
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objItem As MailItem
    
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Get GetCurrentItem function athttp://slipstick.me/e8mio
    Set objItem = GetCurrentItem()
    
    Set objDestFolder = objNamespace.GetDefaultFolder(olFolderInbox).Folders("completed")
    
    objItem.Move objDestFolder
    
    Set objDestFolder = Nothing
    
End Sub

Function GetLastDeleted(sDate) As Object
    sFilter = "[LastModificationTime] > '" & sDate & "'"
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderDeletedItems)
    
    Set myMtgReq = myFolder.Items.Find(sFilter)
    
End Function

' https://it.knightnet.org.uk/kb/ms-office/outlook-macro-open-web-url/
Sub openUrl(sUrl As String)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", sUrl)
End Sub

Sub RunLink(sUrl)
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False
    IE.Navigate sUrl
    Set IE = Nothing
End Sub

Sub CopyAttachments(objSourceItem, objTargetItem)
    ' http://www.outlookcode.com/d/code/copyatts.htm
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fldTemp = FSO.GetSpecialFolder(2)        ' TemporaryFolder
    strPath = fldTemp.Path & "\"
    For Each objAtt In objSourceItem.Attachments
        strFile = strPath & objAtt.Filename
        objAtt.SaveAsFile strFile
        objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
        FSO.DeleteFile strFile
    Next
    
    Set fldTemp = Nothing
    Set FSO = Nothing
End Sub

Function FileOpen(initialFilename As String, _
                  Optional sDesc As String = "Excel (*.xls*)", _
                  Optional sFilter As String = "*.xls*") As String
    With Word.Application.FileDialog(msoFileDialogOpen)
        .ButtonName = "&Open"
        .initialFilename = initialFilename
        .Filters.Clear
        .Filters.Add sDesc, sFilter, 1
        .Title = "File Open"
        .AllowMultiSelect = False
        If .Show = -1 Then
            FileOpen = .SelectedItems(1)
        Else
            FileOpen = ""
        End If
    End With
End Function

' https://stackoverflow.com/a/28237845/2043349
Function IsFile(ByVal fName As String) As Boolean
    'Returns TRUE if the provided name points to an existing file.
    'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

' https://stackoverflow.com/a/218199/2043349
Public Function URLEncode( _
       StringVal As String, _
       Optional SpaceAsPlus As Boolean = False _
       ) As String

    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim Result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)
            Select Case CharCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                Result(i) = Char
            Case 32
                Result(i) = Space
            Case 0 To 15
                Result(i) = "%0" & Hex(CharCode)
            Case Else
                Result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode = Join(Result, "")
    End If
End Function

Function GetEmailAddress()
' https://stackoverflow.com/questions/26519325/how-to-get-the-email-address-of-the-current-logged-in-user
Dim olFol As Outlook.Folder
Set olNS = Application.GetNamespace("MAPI") ' Outlook.NameSpace
Set olFol = olNS.GetDefaultFolder(olFolderInbox)

GetEmailAddress = olFol.Parent.Name
End Function

