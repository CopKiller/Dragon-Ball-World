Attribute VB_Name = "Conv_Handle"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' CONV EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleConvEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    With frmEditor_Conv
        Editor = EDITOR_CONV
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_CONVS
            .lstIndex.AddItem i & ": " & Trim$(Conversation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ConvEditorInit
    End With

End Sub

Public Sub HandleUpdateConv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Convnum As Long
    Dim buffer As clsBuffer
    Dim i As Long
    Dim X As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Convnum = buffer.ReadLong

    With Conversation(Convnum)
        .Name = buffer.ReadString
        .chatCount = buffer.ReadLong
        If .chatCount > 0 Then ReDim Conversation(Convnum).Conv(1 To .chatCount)

        For i = 1 To .chatCount
            .Conv(i).Talk = buffer.ReadString

            For X = 1 To 4
                .Conv(i).rText(X) = buffer.ReadString
                .Conv(i).rTarget(X) = buffer.ReadLong
            Next

            .Conv(i).EventType = buffer.ReadLong
            .Conv(i).EventNum = buffer.ReadLong
        Next

    End With

    buffer.Flush: Set buffer = Nothing
End Sub
