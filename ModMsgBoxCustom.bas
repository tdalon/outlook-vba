Attribute VB_Name = "ModMsgBoxCustom"
' This module includes Private declarations for GetCurrentThreadId, SetWindowsHookEx, SetDlgItemText, CallNextHookEx, UnhookWindowsHookEx
' plus code for Public Sub MsgBoxCustom, Public Sub MsgBoxCustom_Set, Public Sub MsgBoxCustom_Reset
' plus code for Private Sub MsgBoxCustom_Init, Private Function MsgBoxCustom_Proc
' DEVELOPER: J. Woolley (for wellsr.com)
' https://wellsr.com/vba/2019/excel/create-advanced-vba-msgbox-custom-buttons/

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" _
        () As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
        (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" _
        (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As LongPtr) As Long
    Private hHook As LongPtr        ' handle to the Hook procedure (global variable)
#Else
    Private Declare Function GetCurrentThreadId Lib "kernel32" _
        () As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
    Private Declare Function CallNextHookEx Lib "user32" _
        (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" _
        (ByVal hHook As Long) As Long
    Private hHook As Long           ' handle to the Hook procedure (global variable)
#End If
' Hook flags (Computer Based Training)
Private Const WH_CBT = 5            ' hook type
Private Const HCBT_ACTIVATE = 5     ' activate window
' MsgBox constants (these are enumerated by VBA)
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7 (these are button IDs)
' for 1 button, use vbOKOnly = 0 (OK button with ID vbOK returned)
' for 2 buttons, use vbOKCancel = 1 (vbOK, vbCancel) or vbYesNo = 4 (vbYes, vbNo) or vbRetryCancel = 5 (vbRetry, vbCancel)
' for 3 buttons, use vbAbortRetryIgnore = 2 (vbAbort, vbRetry, vbIgnore) or vbYesNoCancel = 3 (vbYes, vbNo, vbCancel)
' Module level global variables
Private sMsgBoxDefaultLabel(1 To 7) As String
Private sMsgBoxCustomLabel(1 To 7) As String
Private bMsgBoxCustomInit As Boolean

Private Sub MsgBoxCustom_Init()
' Initialize default button labels for Public Sub MsgBoxCustom
    Dim nID As Integer
    Dim vA As Variant               ' base 0 array populated by Array function (must be Variant)
    vA = VBA.Array(vbNullString, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No")
    For nID = 1 To 7
        sMsgBoxDefaultLabel(nID) = vA(nID)
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Next nID
    bMsgBoxCustomInit = True
End Sub

Public Sub MsgBoxCustom_Set(ByVal nID As Integer, Optional ByVal vLabel As Variant)
' Set button nID label to CStr(vLabel) for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
' If vLabel is missing, button nID label will be set to default
' vLabel should not have more than 10 characters (approximately)
    If nID = 0 Then Call MsgBoxCustom_Init
    If nID < 1 Or nID > 7 Then Exit Sub
    If Not bMsgBoxCustomInit Then Call MsgBoxCustom_Init
    If IsMissing(vLabel) Then
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Else
        sMsgBoxCustomLabel(nID) = CStr(vLabel)
    End If
End Sub

Public Sub MsgBoxCustom_Reset(ByVal nID As Integer)
' Reset button nID to default label for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
    Call MsgBoxCustom_Set(nID)
End Sub

#If VBA7 Then
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
' Hook callback function for Public Function MsgBoxCustom
    Dim nID As Integer
    If lMsg = HCBT_ACTIVATE And bMsgBoxCustomInit Then
        For nID = 1 To 7
            SetDlgItemText wParam, nID, sMsgBoxCustomLabel(nID)
        Next nID
    End If
    MsgBoxCustom_Proc = CallNextHookEx(hHook, lMsg, wParam, lParam)
End Function

Public Sub MsgBoxCustom( _
    ByRef vID As Variant, _
    ByVal sPrompt As String, _
    Optional ByVal vButtons As Variant = 0, _
    Optional ByVal vTitle As Variant, _
    Optional ByVal vHelpfile As Variant, _
    Optional ByVal vContext As Variant = 0)
' Display standard VBA MsgBox with custom button labels
' Return vID as result from MsgBox corresponding to clicked button (ByRef...Variant is compatible with any type)
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' Arguments sPrompt, vButtons, vTitle, vHelpfile, and vContext match arguments of standard VBA MsgBox function
' This is Public Sub instead of Public Function so it will not be listed as a user-defined function (UDF)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxCustom_Proc, 0, GetCurrentThreadId)
    If IsMissing(vHelpfile) And IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons)
    ElseIf IsMissing(vHelpfile) Then
        vID = MsgBox(sPrompt, vButtons, vTitle)
    ElseIf IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons, , vHelpfile, vContext)
    Else
        vID = MsgBox(sPrompt, vButtons, vTitle, vHelpfile, vContext)
    End If
    If hHook <> 0 Then UnhookWindowsHookEx hHook
End Sub

