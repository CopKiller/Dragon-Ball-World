Attribute VB_Name = "Quest_TCP"
Option Explicit

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next
End Sub

Public Sub SendPlayerQuests(ByVal index As Long, Optional ByVal QuestSelectLst As Integer = 0)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    
    Buffer.WriteInteger QuestSelectLst

    For i = 1 To MAX_QUESTS

        If Player(index).PlayerQuest(i).Status > 0 Then
            Buffer.WriteLong i
            Buffer.WriteLong Player(index).PlayerQuest(i).Status
            Buffer.WriteLong Player(index).PlayerQuest(i).ActualTask
            Buffer.WriteLong Player(index).PlayerQuest(i).CurrentCount


            Buffer.WriteByte Player(index).PlayerQuest(i).TaskTimer.Active
            Buffer.WriteLong Player(index).PlayerQuest(i).TaskTimer.Timer
        End If
    Next

    SendDataTo index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal index As Long, ByVal QuestNum As Long, Optional ByVal QuestSelectLst As Integer = 0)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerQuest
    
    Buffer.WriteInteger QuestSelectLst

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendQuestCancel(ByVal index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestCancel

    Buffer.WriteLong QuestNum
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).Status
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).ActualTask
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).CurrentCount

    Buffer.WriteByte Player(index).PlayerQuest(QuestNum).TaskTimer.Active
    Buffer.WriteLong Player(index).PlayerQuest(QuestNum).TaskTimer.Timer

    SendDataTo index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal index As Long, ByVal QuestNum As Long, ByVal Message As String, Optional ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SQuestMessage
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(Message)
    Buffer.WriteLong saycolour
    Buffer.WriteString "[Quest] "
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub
