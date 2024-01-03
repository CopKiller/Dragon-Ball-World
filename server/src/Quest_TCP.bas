Attribute VB_Name = "Quest_TCP"
Option Explicit

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long, Optional ByVal QuestSelectLst As Integer = 0)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    
    Buffer.WriteInteger QuestSelectLst

    For i = 1 To MAX_QUESTS

        If Player(Index).PlayerQuest(i).Status > 0 Then
            Buffer.WriteLong i
            Buffer.WriteLong Player(Index).PlayerQuest(i).Status
            Buffer.WriteLong Player(Index).PlayerQuest(i).ActualTask
            Buffer.WriteLong Player(Index).PlayerQuest(i).CurrentCount


            Buffer.WriteByte Player(Index).PlayerQuest(i).TaskTimer.Active
            Buffer.WriteLong Player(Index).PlayerQuest(i).TaskTimer.Timer
        End If
    Next

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal Index As Long, ByVal QuestNum As Long, Optional ByVal QuestSelectLst As Integer = 0)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    
    Buffer.WriteInteger QuestSelectLst

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(Index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub SendQuestCancel(ByVal Index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestCancel

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(Index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal Index As Long, ByVal QuestNum As Long, ByVal Message As String, Optional ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(Message)
    Buffer.WriteLong saycolour
    Buffer.WriteString "[Quest] "
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub
