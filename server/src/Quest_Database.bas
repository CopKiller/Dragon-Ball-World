Attribute VB_Name = "Quest_Database"
' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim I As Long
    For I = 1 To MAX_QUESTS
        Call SaveQuest(I)
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim FileName As String
    Dim F As Long, I As Long
    FileName = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    'Alatar v1.2
    Put #F, , Quest(QuestNum).Name
    Put #F, , Quest(QuestNum).Repeat
    Put #F, , Quest(QuestNum).QuestLog
    Put #F, , Quest(QuestNum).Speech
    For I = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).GiveItem(I)
    Next
    For I = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).TakeItem(I)
    Next
    Put #F, , Quest(QuestNum).RequiredLevel
    Put #F, , Quest(QuestNum).RequiredQuest
    For I = 1 To 5
        Put #F, , Quest(QuestNum).RequiredClass(I)
    Next
    For I = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).RequiredItem(I)
    Next
    Put #F, , Quest(QuestNum).RewardExp
    For I = 1 To MAX_QUESTS_ITEMS
        Put #F, , Quest(QuestNum).RewardItem(I)
    Next
    For I = 1 To MAX_TASKS
        Put #F, , Quest(QuestNum).Task(I)
    Next
    '/Alatar v1.2
    Close #F
End Sub

Sub LoadQuests()
    Dim FileName As String
    Dim I As Integer
    Dim F As Long, n As Long
    Dim sLen As Long

    Call CheckQuests

    For I = 1 To MAX_QUESTS
        ' Clear
        Call ClearQuest(I)
        'Load
        FileName = App.Path & "\data\quests\quest" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F

        'Alatar v1.2
        Get #F, , Quest(I).Name
        Get #F, , Quest(I).Repeat
        Get #F, , Quest(I).QuestLog
        Get #F, , Quest(I).Speech
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).TakeItem(n)
        Next
        Get #F, , Quest(I).RequiredLevel
        Get #F, , Quest(I).RequiredQuest
        For n = 1 To 5
            Get #F, , Quest(I).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).RequiredItem(n)
        Next
        Get #F, , Quest(I).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #F, , Quest(I).Task(n)
        Next
        '/Alatar v1.2
        Close #F
    Next
End Sub

Sub CheckQuests()
    Dim I As Long
    For I = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & I & ".dat") Then
            Call SaveQuest(I)
        End If
    Next
End Sub

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
End Sub

Sub ClearQuests()
    Dim I As Long

    For I = 1 To MAX_QUESTS
        Call ClearQuest(I)
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

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteBytes QuestCache(QuestNum).Data    'Sends the entire cache as 1 packet.
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
