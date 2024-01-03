Attribute VB_Name = "Quest_TCP"
Option Explicit

Public Sub UpdateQuestLog()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CQuestLogUpdate
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub PlayerCancelQuest()
    Dim QuestName As String

    With Windows(GetWindowIndex("winQuest"))

        If QuestSelect = 0 Then Exit Sub

        If Not .Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).visible Then Exit Sub

        QuestName = .Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).text

        Dim buffer As clsBuffer
        Set buffer = New clsBuffer

        buffer.WriteLong CPlayerHandleQuest
        buffer.WriteLong FindQuestIndex(QuestName)
        SendData buffer.ToArray()
        Set buffer = Nothing

    End With
End Sub
