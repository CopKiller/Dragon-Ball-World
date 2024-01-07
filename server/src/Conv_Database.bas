Attribute VB_Name = "Conv_Database"
' ***********
' ** Convs **
' ***********
Public Sub SaveConv(ByVal convNum As Long)
    Dim filename As String
    Dim i As Long, x As Long, f As Long

    filename = App.Path & "\data\convs\conv" & convNum & ".dat"
    f = FreeFile

    Open filename For Binary As #f
    With Conversation(convNum)
        Put #f, , .Name
        Put #f, , .chatCount
        For i = 1 To .chatCount
            Put #f, , CLng(Len(.Conv(i).Talk))
            Put #f, , .Conv(i).Talk
            For x = 1 To 4
                Put #f, , CLng(Len(.Conv(i).rText(x)))
                Put #f, , .Conv(i).rText(x)
                Put #f, , .Conv(i).rTarget(x)
            Next
            Put #f, , .Conv(i).EventType
            Put #f, , .Conv(i).EventNum
        Next
    End With
    Close #f
End Sub

Public Sub SaveConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call SaveConv(i)
    Next
End Sub

Public Sub CheckConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        If Not FileExist(App.Path & "\data\convs\conv" & i & ".dat") Then
            Call SaveConv(i)
        End If
    Next
End Sub

Public Sub LoadConvs()
    Dim filename As String
    Dim i As Long, n As Long, x As Long, f As Long
    Dim sLen As Long

    Call CheckConvs

    For i = 1 To MAX_CONVS
        filename = App.Path & "\data\convs\conv" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        With Conversation(i)
            Get #f, , .Name
            Get #f, , .chatCount
            If .chatCount > 0 Then ReDim .Conv(1 To .chatCount)
            For n = 1 To .chatCount
                Get #f, , sLen
                .Conv(n).Talk = Space$(sLen)
                Get #f, , .Conv(n).Talk
                For x = 1 To 4
                    Get #f, , sLen
                    .Conv(n).rText(x) = Space$(sLen)
                    Get #f, , .Conv(n).rText(x)
                    Get #f, , .Conv(n).rTarget(x)
                Next
                Get #f, , .Conv(n).EventType
                Get #f, , .Conv(n).EventNum
            Next
        End With
        Close #f
    Next
End Sub

Public Sub ClearConv(ByVal index As Long)
    Dim i As Long
    
    Conversation(index) = EmptyConv
    Conversation(index).Name = vbNullString
    ReDim Conversation(index).Conv(1)
    
    Conversation(index).Conv(1).Talk = vbNullString
    For i = 1 To 4
        Conversation(index).Conv(1).rText(i) = vbNullString
    Next i
End Sub

Public Sub ClearConvs()
    Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub
