Attribute VB_Name = "Quest_Logic"
Option Explicit

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_STARTED Then    'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED Or Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

Public Sub ShowRewardDesc()
    Dim X As Integer, Y As Integer, Width As Integer, Height As Integer, i As Integer
    Dim itemNum As Long
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_ITEMS
            If .Controls(GetControlIndex("winQuest", "picReward" & i)).visible Then
                X = .Window.Left + .Controls(GetControlIndex("winQuest", "picReward" & i)).Left
                Y = .Window.Top + .Controls(GetControlIndex("winQuest", "picReward" & i)).Top
                Width = .Controls(GetControlIndex("winQuest", "picReward" & i)).Width
                Height = .Controls(GetControlIndex("winQuest", "picReward" & i)).Height
                If GlobalX >= X And GlobalX <= X + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                    itemNum = .Controls(GetControlIndex("winQuest", "picReward" & i)).Value
                    ShowItemDesc GlobalX, GlobalY, itemNum, False
                End If
            End If
        Next i
    End With
End Sub

Public Sub ShowRewardItem()
    Dim X As Integer, Y As Integer, Width As Integer, Height As Integer, i As Integer
    Dim itemNum As Long
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_ITEMS
            If .Controls(GetControlIndex("winQuest", "picReward" & i)).visible Then
                itemNum = .Controls(GetControlIndex("winQuest", "picReward" & i)).Value
                If itemNum > 0 And itemNum <= CountItem Then
                    X = .Window.Left + .Controls(GetControlIndex("winQuest", "picReward" & i)).Left
                    Y = .Window.Top + .Controls(GetControlIndex("winQuest", "picReward" & i)).Top
                    Width = .Controls(GetControlIndex("winQuest", "picReward" & i)).Width
                    Height = .Controls(GetControlIndex("winQuest", "picReward" & i)).Height
                    RenderTexture TextureItem(Item(itemNum).pic), X, Y, 0, 0, Width, Height, PIC_X, PIC_Y
                End If
            End If
        Next i
    End With
End Sub

Public Function FindQuestIndex(ByVal QuestName As String) As Integer
    Dim i As Integer

    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            If QuestInProgress(i) Then
                FindQuestIndex = i
                Exit Function
            End If
        End If
    Next i

End Function

Public Function GetQuestObjetives(ByVal QuestNum As Integer) As String
    Dim i As Byte
    Dim SString As String

    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED_BUT Then
        GetQuestObjetives = "Objetivos ja foram concluidos, voce pode iniciar a missão novamente!"
        Exit Function
    ElseIf Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED Then
        GetQuestObjetives = "Objetivos ja foram concluidos, siga para proxima missao!"
        Exit Function
    End If

    For i = 1 To MAX_TASKS
        If i > Player(MyIndex).PlayerQuest(QuestNum).ActualTask Then

            If Quest(QuestNum).Task(i).Order <> 0 Then
                If i = (Player(MyIndex).PlayerQuest(QuestNum).ActualTask + 1) Then
                    SString = "PROX.:" & Space(1)
                End If
            End If

            Select Case Quest(QuestNum).Task(i).Order
            Case 0    'None

            Case QUEST_TYPE_GOSLAY
                SString = SString & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & "/"
            Case QUEST_TYPE_GOGATHER
                SString = SString & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & "/"
            Case QUEST_TYPE_GOTALK
                SString = SString & "Falar com" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & "/"
            Case QUEST_TYPE_GOREACH
                SString = SString & Quest(QuestNum).Task(i).TaskLog & "/"
            Case QUEST_TYPE_GOGIVE
                SString = SString & "Entregar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1) & "Ao NPC" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & "/"
            Case QUEST_TYPE_GOKILL
                SString = SString & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Jogadores" & "/"
            Case QUEST_TYPE_GOTRAIN
                SString = SString & "Treinar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Vezes na Resource" & Space(1) & Trim$(Resource(Quest(QuestNum).Task(i).Resource).Name) & "/"
            Case QUEST_TYPE_GOGET
                SString = SString & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Item(s) do NPC" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & "/"
            End Select
        End If
    Next i

    GetQuestObjetives = SString

End Function

Public Function GetQuestObjetiveCurrent(ByVal QuestNum As Integer) As String
    Dim i As Byte

    If Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED_BUT Then
        Exit Function
    ElseIf Player(MyIndex).PlayerQuest(QuestNum).status = QUEST_COMPLETED Then
        Exit Function
    End If

    For i = 1 To MAX_TASKS
        If i = Player(MyIndex).PlayerQuest(QuestNum).ActualTask Then

            Select Case Quest(QuestNum).Task(i).Order
            Case 0    'None
                GetQuestObjetiveCurrent = "ATUAL: Nenhum(a)"
            Case QUEST_TYPE_GOSLAY
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & Space(1)
            Case QUEST_TYPE_GOGATHER
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1)
            Case QUEST_TYPE_GOTALK
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Falar com" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & Space(1)
            Case QUEST_TYPE_GOREACH
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & Quest(QuestNum).Task(i).TaskLog & Space(1)
            Case QUEST_TYPE_GOGIVE
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Entregar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & Trim$(Item(Quest(QuestNum).Task(i).Item).Name) & Space(1) & "Ao NPC" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & Space(1)
            Case QUEST_TYPE_GOKILL
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Derrotar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Jogadores" & Space(1)
            Case QUEST_TYPE_GOTRAIN
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Treinar" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Vezes na Resource" & Space(1) & Trim$(Resource(Quest(QuestNum).Task(i).Resource).Name) & Space(1)
            Case QUEST_TYPE_GOGET
                GetQuestObjetiveCurrent = "ATUAL:" & Space(1) & "Obter" & Space(1) & Quest(QuestNum).Task(i).Amount & Space(1) & "Item(s) do NPC" & Space(1) & Trim$(Npc(Quest(QuestNum).Task(i).Npc).Name) & Space(1)
            End Select

            Exit Function
        End If
    Next i

End Function

Public Sub CalculateQuestTimer()
    Dim i As Integer

    With Windows(GetWindowIndex("winQuest"))

        If QuestSelect > 0 Then
            i = FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & QuestSelect)).text)
            If i > 0 And i <= MAX_QUESTS Then
                If LenB(Trim$(Quest(i).Name)) > 0 Then
                    If Player(MyIndex).PlayerQuest(i).status = QUEST_STARTED Then
                        If Player(MyIndex).PlayerQuest(i).TaskTimer.active = YES Then
                            If Player(MyIndex).PlayerQuest(i).TaskTimer.timer > 0 Then
                                Player(MyIndex).PlayerQuest(i).TaskTimer.timer = Player(MyIndex).PlayerQuest(i).TaskTimer.timer - 1
                                QuestNameToFinish = "Quest: " & Trim$(Quest(i).Name)
                                QuestTimeToFinish = "Tempo da Task: " & SecondsToHMS(CLng(Player(MyIndex).PlayerQuest(i).TaskTimer.timer))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

