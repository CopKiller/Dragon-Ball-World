Attribute VB_Name = "Conv_Handle"
' :::::::::::::::::::::::::::::
' :: Request edit Conv packet ::
' :::::::::::::::::::::::::::::
Public Sub HandleRequestEditConv(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SConvEditor
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleRequestConvs(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendConvs index
End Sub

' :::::::::::::::::::::::
' :: Save Conv packet ::
' :::::::::::::::::::::::
Public Sub HandleSaveConv(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim x As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    ' Prevent hacking
    If n < 0 Or n > MAX_CONVS Then
        Exit Sub
    End If

    With Conversation(n)
        .Name = Buffer.ReadString
        .chatCount = Buffer.ReadLong
        ReDim .Conv(1 To .chatCount)
        For i = 1 To .chatCount
            .Conv(i).Talk = Buffer.ReadString
            For x = 1 To 4
                .Conv(i).rText(x) = Buffer.ReadString
                .Conv(i).rTarget(x) = Buffer.ReadLong
            Next
            .Conv(i).EventType = Buffer.ReadLong
            .Conv(i).EventNum = Buffer.ReadLong
        Next
    End With
    
    ' Save it
    Call SendUpdateConvToAll(n)
    Call SaveConv(n)
    Call AddLog(GetPlayerName(index) & " saved Conv #" & n & ".", ADMIN_LOG)
End Sub
