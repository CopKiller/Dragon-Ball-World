Attribute VB_Name = "Quest_Logic"
Option Explicit

Public Function CanStartQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    Dim I As Long, n As Long
    CanStartQuest = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If QuestInProgress(Index, QuestNum) Then
        Call QuestMessage(Index, QuestNum, "Você ja iniciou a quest, precisa termina-la!", BrightRed)
        Exit Function
    End If

    'check if now a completed quest can be repeated
    Select Case Player(Index).PlayerQuest(QuestNum).Status
    Case QUEST_COMPLETED    ' Normal?
        If Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_BUT    ' Repetível?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_DIARY    ' Diaria?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 3 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    Case QUEST_COMPLETED_TIME    ' Tempo pra refazer?
        If Quest(QuestNum).Repeat = 0 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        ElseIf Quest(QuestNum).Repeat = 1 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
        ElseIf Quest(QuestNum).Repeat = 2 Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        End If
    End Select

    ' Fazer o processamento da quest diaria e quest por tempo!
    Select Case Player(Index).PlayerQuest(QuestNum).Status
    Case QUEST_COMPLETED_DIARY
        If Format(Player(Index).PlayerQuest(QuestNum).Data, "dd/mm/yyyy") <> CStr(Date) Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        Else
            PlayerMsg Index, "Você ja realizou essa missão hoje, volte novamente amanhã!", BrightRed
            Exit Function
        End If
    Case QUEST_COMPLETED_TIME
        If DateDiff("s", Player(Index).PlayerQuest(QuestNum).Data, Now) >= Quest(QuestNum).Time Then
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED
        Else
            PlayerMsg Index, "Aguarde: " & SecondsToHMS(Quest(QuestNum).Time - DateDiff("s", Player(Index).PlayerQuest(QuestNum).Data, Now)), BrightRed
            Exit Function
        End If
    End Select

    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(QuestNum).RequiredLevel <= Player(Index).Level Then

            'Check if item is needed
            For I = 1 To MAX_QUESTS_ITEMS
                If Quest(QuestNum).RequiredItem(I).Item > 0 Then
                    'if we don't have it at all then
                    If HasItem(Index, Quest(QuestNum).RequiredItem(I).Item) = 0 Then
                        PlayerMsg Index, "You need " & Trim$(Item(Quest(QuestNum).RequiredItem(I).Item).Name) & " to take this quest!", BrightRed
                        Exit Function
                    End If
                End If
            Next

            'Check if previous quest is needed
            If Quest(QuestNum).RequiredQuest > 0 And Quest(QuestNum).RequiredQuest <= MAX_QUESTS Then
                If Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_NOT_STARTED Or Player(Index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_STARTED Then
                    PlayerMsg Index, "You need to complete the " & Trim$(Quest(Quest(QuestNum).RequiredQuest).Name) & " quest in order to take this quest!", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg Index, "You need to be a higher level to take this quest!", BrightRed
        End If
    Else
        PlayerMsg Index, "You can't start that quest again!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal Index As Long, QuestNum As Long) As Boolean
    CanEndQuest = False
    If Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED Then
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function

    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim I As Long
    GetQuestNum = 0

    For I = 1 To MAX_QUESTS
        If Trim$(Quest(I).Name) = Trim$(QuestName) Then
            GetQuestNum = I
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim I As Long
    GetItemNum = 0

    For I = 1 To MAX_ITEMS
        If Trim$(Item(I).Name) = Trim$(ItemName) Then
            GetItemNum = I
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal Index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim I As Long

    For I = 1 To MAX_QUESTS
        If QuestInProgress(Index, I) Then
            If TaskType = Quest(I).Task(Player(Index).PlayerQuest(I).ActualTask).Order Then
                Call CheckTask(Index, I, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal Index As Long, ByVal QuestNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, I As Long
    ActualTask = Player(Index).PlayerQuest(QuestNum).ActualTask

    Select Case TaskType
    Case QUEST_TYPE_GOSLAY    'Kill X amount of X npc's.

        'is npc's defeated id is the same as the npc i have to kill?
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Npc Then
            'Count +1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            'show msg
            QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Npc(TargetIndex).Name) + " killed.", Yellow
            'did i finish the work?
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", Green
                'is the quest's end?
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    'otherwise continue to the next task
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOGATHER    'Gather X amount of X item.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Item Then

            'reset the count first
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

            'Check inventory for the items
            For I = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, I) = TargetIndex Then
                    If Item(I).Type = ITEM_TYPE_CURRENCY Then
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, I)
                    Else
                        'If is the correct item add it to the count
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                    End If
                End If
            Next

            QuestMessage Index, QuestNum, "You have " + Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow

            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", Green
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOTALK    'Interact with X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Npc Then
            QuestMessage Index, QuestNum, "Task completed", Green
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else
                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOREACH    'Reach X map.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Map Then
            QuestMessage Index, QuestNum, "Task completed", Green
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else

                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOGIVE    'Give X amount of X item to X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Npc Then

            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

            For I = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, I) = Quest(QuestNum).Task(ActualTask).Item Then
                    If Item(I).Type = ITEM_TYPE_CURRENCY Then
                        If GetPlayerInvItemValue(Index, I) >= Quest(QuestNum).Task(ActualTask).Amount Then
                            Player(Index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(Index, I)
                        End If
                    Else
                        'If is the correct item add it to the count
                        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
                    End If
                End If
            Next

            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                'if we have enough items, then remove them and finish the task
                If Item(Quest(QuestNum).Task(ActualTask).Item).Type = ITEM_TYPE_CURRENCY Then
                    TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                Else
                    'If it's not a currency then remove all the items
                    For I = 1 To Quest(QuestNum).Task(ActualTask).Amount
                        TakeInvItem Index, Quest(QuestNum).Task(ActualTask).Item, 1
                    Next
                End If

                QuestMessage Index, QuestNum, "You gave " + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                QuestMessage Index, QuestNum, "Task completed", Green

                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOKILL    'Kill X amount of players.
        Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
        QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " players killed.", Yellow
        If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
            QuestMessage Index, QuestNum, "Task completed", Green
            If CanEndQuest(Index, QuestNum) Then
                EndQuest Index, QuestNum
            Else
                Call ResetPlayerTaskTimer(Index, QuestNum)
                Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                Call SetPlayerTaskTimer(Index, QuestNum)
                'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
            End If
        End If

    Case QUEST_TYPE_GOTRAIN    'Hit X amount of times X resource.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Resource Then
            Player(Index).PlayerQuest(QuestNum).CurrentCount = Player(Index).PlayerQuest(QuestNum).CurrentCount + 1
            QuestMessage Index, QuestNum, Trim$(Player(Index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " hits.", Yellow
            If Player(Index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage Index, QuestNum, "Task completed", Green
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else
                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    Case QUEST_TYPE_GOGET    'Get X amount of X item from X npc.
        If TargetIndex = Quest(QuestNum).Task(ActualTask).Npc Then
            If GiveInvItem(Index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount, 0) Then
                QuestMessage Index, QuestNum, Quest(QuestNum).Task(ActualTask).TaskLog, Yellow
                If CanEndQuest(Index, QuestNum) Then
                    EndQuest Index, QuestNum
                Else

                    Call ResetPlayerTaskTimer(Index, QuestNum)
                    Player(Index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    Call SetPlayerTaskTimer(Index, QuestNum)
                    'QuestMessage index, QuestNum, "New Task: " & Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).TaskLog, Yellow
                    SendMessageTo Index, "New Task:", Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskLog
                End If
            End If
        End If

    End Select
    SavePlayer Index
    SendPlayerQuest Index, QuestNum
End Sub

Public Sub EndQuest(ByVal Index As Long, ByVal QuestNum As Long)
    Dim I As Long, n As Long

    ' Reseta os dados da data pra ser somente usado onde necessitar!
    Player(Index).PlayerQuest(QuestNum).Data = vbNullString

    If Quest(QuestNum).Repeat = 0 Then    ' Normal?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED
    ElseIf Quest(QuestNum).Repeat = 1 Then    ' Repetível?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
    ElseIf Quest(QuestNum).Repeat = 2 Then    ' Diaria?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_DIARY
        Player(Index).PlayerQuest(QuestNum).Data = Now
    ElseIf Quest(QuestNum).Repeat = 3 Then    ' Tempo pra refazer?
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_TIME
        Player(Index).PlayerQuest(QuestNum).Data = Now
    End If

    'reset counters to 0
    Call ResetPlayerTaskTimer(Index, QuestNum)
    Player(Index).PlayerQuest(QuestNum).ActualTask = 0
    Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

    'give experience
    GivePlayerEXP Index, Quest(QuestNum).RewardExp

    'give levels
    If Quest(QuestNum).RewardLevel > 0 Then
        CheckPlayerLevelUp Index, Quest(QuestNum).RewardLevel
    End If

    'remove items on the end
    For I = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).TakeItem(I).Item > 0 Then
            If HasItem(Index, Quest(QuestNum).TakeItem(I).Item) > 0 Then
                If Item(Quest(QuestNum).TakeItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                    TakeInvItem Index, Quest(QuestNum).TakeItem(I).Item, Quest(QuestNum).TakeItem(I).Value
                Else
                    For n = 1 To Quest(QuestNum).TakeItem(I).Value
                        TakeInvItem Index, Quest(QuestNum).TakeItem(I).Item, 1
                    Next
                End If
            End If
        End If
    Next

    'give rewards
    For I = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).RewardItem(I).Item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                PlayerMsg Index, "You have no inventory space.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If Item(Quest(QuestNum).RewardItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                    GiveInvItem Index, Quest(QuestNum).RewardItem(I).Item, Quest(QuestNum).RewardItem(I).Value, 0
                Else
                    'if not, create a new loop and store the item in a new slot if is possible
                    For n = 1 To Quest(QuestNum).RewardItem(I).Value
                        If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit For
                        Else
                            GiveInvItem Index, Quest(QuestNum).RewardItem(I).Item, 1, 0
                        End If
                    Next
                End If
            End If
        End If
    Next

    ' Give Spell Reward
    If Quest(QuestNum).RewardSpell > 0 Then
        Call GivePlayerSpell(Index, Quest(QuestNum).RewardSpell)
    End If

    'show ending message
    'QuestMessage Index, QuestNum, "Parabens, Você concluiu a missão!", LightGreen
    If Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_DIARY Then
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao, volte amanha para completar novamente!"
    ElseIf Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_TIME Then
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao, volte daqui: " & SecondsToHMS(Quest(QuestNum).Time) & " e complete novamente!"
    ElseIf Player(Index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao. Esta missão é repetitiva!"
    Else
        SendMessageTo Index, Trim$(Quest(QuestNum).Name), "Parabens, Voce concluiu a missao!"
    End If

    SavePlayer Index
    SendEXP Index
    SendStats Index
    SendPlayerQuest Index, QuestNum
End Sub

Public Function GivePlayerSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim I As Long, FreeSlot As Long
    
    GivePlayerSpell = False

    ' Se o usuário já estiver com a magia, atualiza o level.
    For I = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(I).Spell = SpellNum Then
            Call PlayerMsg(Index, "Você já possui essa habilidade", BrightRed)
            GivePlayerSpell = True
            Exit Function
        End If
    Next

    ' Procura por um slot vazio.
    For I = 1 To MAX_PLAYER_SPELLS
        If Player(Index).Spell(I).Spell = 0 Then
            FreeSlot = I
            Exit For
        End If
    Next

    If FreeSlot <> 0 Then
        Player(Index).Spell(FreeSlot).Spell = SpellNum

        Call PlayerMsg(Index, "A habilidade " & Trim$(Spell(SpellNum).Name) & " foi adquirida", BrightGreen)
        Call SendPlayerSpells(Index)
        GivePlayerSpell = True
    Else
        Call PlayerMsg(Index, "Não há espaço suficiente para novas magias", BrightRed)
    End If

End Function

Public Sub StartQuest(ByVal Index As Long, ByVal QuestNum As Long, ByVal Order As Byte)
    Dim I As Long, n As Long
    Dim RemoveStartItems As Boolean

    If Order = 1 Then    'Iniciar
        RemoveStartItems = False
        For I = 1 To MAX_QUESTS_ITEMS

            If Quest(QuestNum).RewardItem(I).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                    QuestMessage Index, QuestNum, "Você não tem espaço na mochila, drope algo para pegar a quest.", Red
                    Exit For
                End If
            End If

            If Quest(QuestNum).GiveItem(I).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).GiveItem(I).Item) = 0 Then
                    QuestMessage Index, QuestNum, "Você não tem espaço na mochila, drope algo para pegar a quest.", Red
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem Index, Quest(QuestNum).GiveItem(I).Item, Quest(QuestNum).GiveItem(I).Value, 0
                    Else
                        GiveInvItem Index, Quest(QuestNum).GiveItem(I).Item, 1, 0
                    End If
                End If
            End If


        Next

        If RemoveStartItems = False Then    'this means everything went ok
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED    '1
            Player(Index).PlayerQuest(QuestNum).ActualTask = 1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
            QuestMessage Index, QuestNum, "Nova missão aceita, olhe seu QuestLog!", BrightGreen

            Call SetPlayerTaskTimer(Index, QuestNum)
        End If

    ElseIf Order = 2 Then
        Call ResetPlayerTaskTimer(Index, QuestNum)
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED    '2
        Player(Index).PlayerQuest(QuestNum).ActualTask = 1
        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0

        RemoveStartItems = True    'avoid exploits
        QuestMessage Index, QuestNum, " foi cancelada!", Yellow
    End If

    If RemoveStartItems = True Then
        For I = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(I).Item > 0 Then
                If HasItem(Index, Quest(QuestNum).GiveItem(I).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem Index, Quest(QuestNum).GiveItem(I).Item, Quest(QuestNum).GiveItem(I).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(I).Value
                            TakeInvItem Index, Quest(QuestNum).GiveItem(I).Item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If

    SavePlayer Index
    SendPlayerQuest Index, QuestNum, QuestNum
End Sub

Public Sub ResetPlayerTaskTimer(ByVal Index As Long, ByVal QuestNum As Integer)
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Active = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.mapnum = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.ResetType = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Teleport = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.Timer = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.TimerType = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.x = 0
    Player(Index).PlayerQuest(QuestNum).TaskTimer.y = 0
End Sub

Public Sub SetPlayerTaskTimer(ByVal Index As Long, ByVal QuestNum As Integer)
    With Player(Index).PlayerQuest(QuestNum).TaskTimer
        .Active = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Active
        .Teleport = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Teleport
        .mapnum = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.mapnum
        .x = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.x
        .y = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.y
        .ResetType = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.ResetType


        .TimerType = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.TimerType

        ' Converter o tipo de contador pelo menor pra ter um melhor processamento pelo loop
        If .TimerType = TaskType.Day Then
            .Timer = (((Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 24) * 60) * 60)
        ElseIf .TimerType = TaskType.Hour Then
            .Timer = ((Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 60) * 60)
        ElseIf .TimerType = TaskType.Minutes Then
            .Timer = (Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer * 60)
        Else    ' segundos já pré configurado no editor
            .Timer = Quest(QuestNum).Task(Player(Index).PlayerQuest(QuestNum).ActualTask).TaskTimer.Timer
        End If
    End With
End Sub

Public Sub CheckPlayerTaskTimer(ByVal Index As Long)
    Dim I As Integer

    If IsPlaying(Index) Then
        For I = 1 To MAX_QUESTS
            If LenB(Trim$(Quest(I).Name)) > 0 Then
                If QuestInProgress(Index, I) Then
                    With Player(Index).PlayerQuest(I).TaskTimer

                        If .Active = YES Then
                            If .Timer > 0 Then
                                .Timer = .Timer - 1
                            End If

                            If .Timer <= 0 Then
                                If .Teleport = YES Then
                                    If .mapnum > 0 And .mapnum <= MAX_MAPS Then
                                        Call PlayerWarp(Index, .mapnum, .x, .y)
                                    Else
                                        Call PlayerWarp(Index, START_MAP, START_X, START_Y)
                                    End If
                                End If

                                ' 0=Resetar Task ; 1=Resetar Quest.
                                If .ResetType = 0 Then
                                    Player(Index).PlayerQuest(I).CurrentCount = 0    ' Retornar a zero a contagem do objetivo da task.
                                    .Timer = Quest(I).Task(Player(Index).PlayerQuest(I).ActualTask).TaskTimer.Timer    ' Resetar o tempo que o jogador vai refazê-lá.
                                ElseIf .ResetType = 1 Then
                                    Call ResetPlayerTaskTimer(Index, I)    ' Resetar todo os dados da task das variaveis do jogador!
                                    Call StartQuest(Index, I, 2)    ' Cancelar a quest toda!
                                End If

                                ' enviar a mensagem do editor de task
                                If Trim$(Quest(I).Task(Player(Index).PlayerQuest(I).ActualTask).TaskTimer.Msg) <> vbNullString Then
                                    Call SendMessageTo(Index, Trim$(Quest(I).Name), Trim$(Quest(I).Task(Player(Index).PlayerQuest(I).ActualTask).TaskTimer.Msg))
                                End If

                                Call SendPlayerQuest(Index, I)
                            End If
                        Else
                            If .Teleport = YES Then
                                If .mapnum > 0 And .mapnum <= MAX_MAPS Then
                                    Call PlayerWarp(Index, .mapnum, .x, .y)
                                Else
                                    Call PlayerWarp(Index, START_MAP, START_X, START_Y)
                                End If

                                Call ResetPlayerTaskTimer(Index, I)
                            End If
                        End If
                    End With
                End If
            End If
        Next I
    End If

End Sub

