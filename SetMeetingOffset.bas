Attribute VB_Name = "SetMeetingOffset"
Public Sub InputMeetingOffsets()
    Dim intMinimumDuration As Integer
    Dim intStartOffset As Integer
    Dim intEndOffset As Integer

    Dim oMtgOffsets As New clsMeetingOffsets

    With oMtgOffsets
        .StartOffset = InputBox("Meeting start offset [min]:", "Meeting Offsets", .StartOffset)
        .EndOffset = InputBox("Meeting end offset [min]:", "Meeting Offsets", .EndOffset)
        .MinimumDuration = InputBox("Meeting offsets not applied if duration results to less than [min]:", "Meeting Offsets", .MinimumDuration)
    End With
End Sub

