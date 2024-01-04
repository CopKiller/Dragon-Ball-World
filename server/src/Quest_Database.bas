Attribute VB_Name = "Quest_Database"
' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim f As Long, i As Long
    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    'Alatar v1.2
    Put #f, , Quest(QuestNum).Name
    Put #f, , Quest(QuestNum).Repeat
    Put #f, , Quest(QuestNum).QuestLog
    Put #f, , Quest(QuestNum).Speech
    For i = 1 To MAX_QUESTS_ITEMS
        Put #f, , Quest(QuestNum).GiveItem(i)
    Next
    For i = 1 To MAX_QUESTS_ITEMS
        Put #f, , Quest(QuestNum).TakeItem(i)
    Next
    Put #f, , Quest(QuestNum).RequiredLevel
    Put #f, , Quest(QuestNum).RequiredQuest
    For i = 1 To 5
        Put #f, , Quest(QuestNum).RequiredClass(i)
    Next
    For i = 1 To MAX_QUESTS_ITEMS
        Put #f, , Quest(QuestNum).RequiredItem(i)
    Next
    Put #f, , Quest(QuestNum).RewardExp
    For i = 1 To MAX_QUESTS_ITEMS
        Put #f, , Quest(QuestNum).RewardItem(i)
    Next
    For i = 1 To MAX_TASKS
        Put #f, , Quest(QuestNum).Task(i)
    Next
    '/Alatar v1.2
    Close #f
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Integer
    Dim f As Long, n As Long
    Dim sLen As Long

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        ' Clear
        Call ClearQuest(i)
        'Load
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f

        'Alatar v1.2
        Get #f, , Quest(i).Name
        Get #f, , Quest(i).Repeat
        Get #f, , Quest(i).QuestLog
        Get #f, , Quest(i).Speech
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , Quest(i).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , Quest(i).TakeItem(n)
        Next
        Get #f, , Quest(i).RequiredLevel
        Get #f, , Quest(i).RequiredQuest
        For n = 1 To 5
            Get #f, , Quest(i).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , Quest(i).RequiredItem(n)
        Next
        Get #f, , Quest(i).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #f, , Quest(i).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #f, , Quest(i).Task(n)
        Next
        '/Alatar v1.2
        Close #f
    Next
End Sub

Sub CheckQuests()
    Dim i As Long
    For i = 1 To MAX_QUESTS
        If Not FileExist(App.Path & "\data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Sub ClearQuest(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Quest(index).Name = vbNullString
    Quest(index).QuestLog = vbNullString
    
    For i = 1 To MAX_TASKS
        Quest(index).Task(i).TaskLog = vbNullString
        Quest(index).Task(i).TaskTimer.Msg = vbNullString
    Next i
End Sub

Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

Public Sub QuestCache_Create(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    Buffer.CompressBuffer
    QuestCache(QuestNum).Data = Buffer.ToArray()    'Buffers entire cache for its packet sending.
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendQuestAll(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteBytes QuestCache(QuestNum).Data    'Sends the entire cache as 1 packet.
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteBytes QuestCache(QuestNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
