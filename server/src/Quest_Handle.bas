Attribute VB_Name = "Quest_Handle"
Option Explicit

Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If

    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing

    ' Save it
    Call QuestCache_Create(n)
    Call SendQuestAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests Index
End Sub

Sub HandlePlayerCancelQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim QuestNum As Long, I As Long, n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong

    Call ResetPlayerTaskTimer(Index, QuestNum)
    Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED    '2
    Player(Index).PlayerQuest(QuestNum).ActualTask = 1
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

    PlayerMsg Index, Trim$(Quest(QuestNum).Name) & " has been canceled!", BrightGreen
    For I = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).GiveItem(I).Item > 0 Then
            If HasItem(Index, Quest(QuestNum).GiveItem(I).Item) > 0 Then
                If Item(Quest(QuestNum).GiveItem(I).Item).Stackable > 0 Then
                    TakeInvItem Index, Quest(QuestNum).GiveItem(I).Item, Quest(QuestNum).GiveItem(I).Value
                Else
                    For n = 1 To Quest(QuestNum).GiveItem(I).Value
                        TakeInvItem Index, Quest(QuestNum).GiveItem(I).Item, 1
                    Next
                End If
            End If
        End If
    Next

    SavePlayer Index
    SendQuestCancel Index, QuestNum

    Set Buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests Index
End Sub
