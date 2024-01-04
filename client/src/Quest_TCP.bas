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
    Dim X As Long, Y As Long, Width As Long, Height As Long
    Dim i As Long, questName As String

    With Windows(GetWindowIndex("winQuest"))

      '  If Not .Window.activated Then Exit Sub

        For i = 1 To MAX_QUESTS_WINDOW
            If .Controls(GetControlIndex("winQuest", "btnCancel" & i)).visible Then
                X = .Window.Left + .Controls(GetControlIndex("winQuest", "btnCancel" & i)).Left
                Y = .Window.Top + .Controls(GetControlIndex("winQuest", "btnCancel" & i)).Top
                Width = .Controls(GetControlIndex("winQuest", "btnCancel" & i)).Width
                Height = .Controls(GetControlIndex("winQuest", "btnCancel" & i)).Height


                If GlobalX >= X And GlobalX <= X + Width Then
                    If GlobalY >= Y And GlobalY <= Y + Height Then
                        questName = .Controls(GetControlIndex("winQuest", "lblList" & i)).text
                        Call Dialogue(questName, "Cancelar Quest", "Deseja cancelar agora?", TypeQUESTCANCEL, styleyesno, FindQuestIndex(questName))
                        Exit Sub
                    End If
                End If
            End If
        Next i
    End With
End Sub

Public Sub CancelQuest(ByVal questNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong CPlayerHandleQuest
    buffer.WriteLong questNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
