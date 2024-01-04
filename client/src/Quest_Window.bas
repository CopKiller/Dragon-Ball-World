Attribute VB_Name = "Quest_Window"
Option Explicit

Private Const QuestOffsetX As Long = 20
Private Const QuestOffsetY As Long = 10

Private Const ListOffsetY As Integer = 25
Private Const RewardOffsetX As Integer = 21
Private Const ListX As Integer = 18
Private Const ListY As Integer = 43

Private Const DescriptionX As Integer = 182
Private Const DescriptionY As Integer = 43

' Quantidade de quests mostradas na janela
Public Const MAX_QUESTS_WINDOW As Byte = 13

Private Const QuestMouseMoveColour = Yellow
Private Const QuestMouseDownColour = Yellow
Private Const QuestDefaultColour = White

Public QuestSelect As Byte

Public QuestTimeToFinish As String
Public QuestNameToFinish As String

Public Sub CreateWindow_Quest()
    Dim i As Long
    
    ' Create window
    CreateWindow "winQuest", "Quests em Andamento...", GetOrder_Win, 0, 0, 436, 406, TextureItem(23), False, Fonts.Default, , 2, 7, DesignTypes.DesignWindowNormalIcon, DesignTypes.DesignWindowNormalIcon, DesignTypes.DesignWindowNormalIcon

    ' Centralise it
    CentraliseWindow windowCount

    ' Set the index for spawning controls
    SetzOrder_Con 1

    ' Close button
    CreateButton windowCount, "btnClose", Windows(windowCount).Window.Width - 39, 2, 36, 36, , , , , , , TextureGUI(3), TextureGUI(4), TextureGUI(5), , , GetAddress(AddressOf btnMenu_Quest), , , GetAddress(AddressOf btnMenu_Quest)

    ' Parchment
    CreatePictureBox windowCount, "picList", ListX - 14, ListY + 1, 175, 358, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , , GetAddress(AddressOf lblList_ClearColour)
    CreatePictureBox windowCount, "picDescription", DescriptionX, DescriptionY, 250, 358, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , , , GetAddress(AddressOf lblList_ClearColour)

    ' Shadow
    CreatePictureBox windowCount, "picShadow_1", ListX + 1, ListY + 10, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestList", ListX - 14, ListY + 7, 175, 25, "Quest's Name", rockwellDec_15, White, Alignment.alignCentre

    ' Shadow descrição
    CreatePictureBox windowCount, "picShadow_1", ListX + 215, ListY + 10, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestDes", ListX - -200, ListY + 7, 175, 25, "Description", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox windowCount, "picBackground", ListX - -175, ListY + 25, 219, 124, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestDescription1", ListX - -175, ListY + 25, 219, 124, "", rockwellDec_15, White, Alignment.alignCentre

    ' Shadow Objective
    CreatePictureBox windowCount, "picShadow_1", ListX + 215, ListY + 158, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestObj", ListX - -200, ListY + 155, 175, 25, "Objective", rockwellDec_15, Yellow, Alignment.alignCentre
    CreatePictureBox windowCount, "picBackgroun2", ListX - -175, ListY + 175, 219, 78, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestDescription2", ListX - -175, ListY + 175, 219, 124, "", rockwellDec_15, Yellow, Alignment.alignCentre

    ' Text Rewards
    CreatePictureBox windowCount, "picShadow_1", ListX + 215, ListY + 260, 142, 9, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestRew", ListX - -200, ListY + 257, 175, 25, "Rewards", rockwellDec_15, BrightGreen, Alignment.alignCentre
    CreatePictureBox windowCount, "picBackground3", ListX - -175, ListY + 276, 219, 70, , , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment
    CreateLabel windowCount, "lblQuestDescription3", ListX - -175, ListY + 276, 219, 70, "", rockwellDec_15, BrightGreen, Alignment.alignCentre

    
    For i = 1 To MAX_QUESTS_ITEMS
        CreatePictureBox windowCount, "picReward" & i, (ListX + 160) + (RewardOffsetX * i), ListY + 320, 20, 20, True, , , , , , , DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, DesignTypes.DesignBlackParchment, , GetAddress(AddressOf ShowRewardDesc), , GetAddress(AddressOf ShowRewardDesc), , GetAddress(AddressOf ShowRewardItem)
    Next i

    For i = 1 To MAX_QUESTS_WINDOW
        CreatePictureBox windowCount, "picList" & i, ListX, ListY + (ListOffsetY * i), 130, 20, False, , , , , , , DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, DesignTypes.DesignTextInput, , , 0
        CreateLabel windowCount, "lblList" & i, ListX, ListY + (ListOffsetY * i) + 3, 130, 20, "Vazio", , , Alignment.alignCentre, False, , , , , GetAddress(AddressOf lblList_MouseDown), GetAddress(AddressOf lblList_MouseMove)
        CreateButton windowCount, "btnCancel" & i, ListX + 130, ListY + (ListOffsetY * i), 20, 20, "X", rockwellDec_15, White, , False, , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf PlayerCancelQuest)
    Next i

    ' Btns
    'CreateButton windowCount, "btnCancel", 238, 385, 134, 20, "Cancel Quest", rockwellDec_15, White, , , , , , , DesignTypes.DesignRedNormal, DesignTypes.DesignRedHover, DesignTypes.DesignRedClick, , , GetAddress(AddressOf PlayerCancelQuest)
End Sub

Public Sub lblList_MouseMove()
    Dim i As Byte, X As Long, Y As Long, Width As Long, Height As Long
    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To MAX_QUESTS_WINDOW

            X = .Window.Left + .Controls(GetControlIndex("winQuest", "lblList" & i)).Left
            Y = .Window.Top + .Controls(GetControlIndex("winQuest", "lblList" & i)).Top
            Width = .Controls(GetControlIndex("winQuest", "lblList" & i)).Width
            Height = .Controls(GetControlIndex("winQuest", "lblList" & i)).Height


            If QuestSelect <> i Then
                If GlobalX >= X And GlobalX <= X + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseMoveColour
                Else
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
                End If
            End If

        Next i
    End With
End Sub

Public Sub lblList_MouseDown()
    Dim i As Byte, X As Long, Y As Long, Width As Long, Height As Long
    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To MAX_QUESTS_WINDOW

            X = .Window.Left + .Controls(GetControlIndex("winQuest", "lblList" & i)).Left
            Y = .Window.Top + .Controls(GetControlIndex("winQuest", "lblList" & i)).Top
            Width = .Controls(GetControlIndex("winQuest", "lblList" & i)).Width
            Height = .Controls(GetControlIndex("winQuest", "lblList" & i)).Height

            If GlobalX >= X And GlobalX <= X + Width And GlobalY >= Y And GlobalY <= Y + Height Then
                If QuestSelect = i Then
                    .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
                    QuestSelect = 0
                    QuestTimeToFinish = vbNullString
                    QuestNameToFinish = vbNullString
                    ClearQuestLogBox
                    Exit For
                End If
                QuestSelect = i
                If Player(MyIndex).PlayerQuest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).text)).TaskTimer.timer > 0 Then
                    QuestNameToFinish = "Quest: " & Trim$(Quest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).text)).Name)
                    QuestTimeToFinish = "Tempo da Task: " & SecondsToHMS(CLng(Player(MyIndex).PlayerQuest(FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).text)).TaskTimer.timer))
                End If
                LoadQuestLogBox QuestSelect
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseDownColour
            Else
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
            End If

        Next i
    End With
End Sub

Public Sub lblList_ClearColour()
    Dim i As Byte
    With Windows(GetWindowIndex("winQuest"))
        For i = 1 To MAX_QUESTS_WINDOW
            If .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestMouseMoveColour Then
                .Controls(GetControlIndex("winQuest", "lblList" & i)).textColour = QuestDefaultColour
            End If
        Next i
    End With
End Sub

Public Sub RefreshQuestWindow()
    Dim i As Long, n As Long, LastQuest As Integer

    With Windows(GetWindowIndex("winQuest"))

        For n = 1 To MAX_QUESTS_WINDOW

            'clear
            .Controls(GetControlIndex("winQuest", "lblList" & n)).text = "Vazio"
            .Controls(GetControlIndex("winQuest", "lblList" & n)).textColour = QuestDefaultColour
            .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = False
            .Controls(GetControlIndex("winQuest", "picList" & n)).visible = False
            .Controls(GetControlIndex("winQuest", "btnCancel" & n)).visible = False
            ClearQuestLogBox

            For i = 1 To MAX_QUESTS
                If QuestInProgress(i) Then
                    If LastQuest < i Then
                        .Controls(GetControlIndex("winQuest", "lblList" & n)).text = Trim$(Quest(i).Name)
                        .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = True
                        .Controls(GetControlIndex("winQuest", "picList" & n)).visible = True
                        
                        .Controls(GetControlIndex("winQuest", "btnCancel" & n)).visible = True
                        LastQuest = i
                        QuestSelect = n
                        Exit For
                    End If
                End If
            Next i

            If .Controls(GetControlIndex("winQuest", "lblList" & n)).visible = False Then Exit For
        Next n
        
        If QuestSelect > 0 Then
            LoadQuestLogBox QuestSelect
        End If

    End With
End Sub

Public Sub ClearQuestLogBox()
    Dim i As Byte

    With Windows(GetWindowIndex("winQuest"))

        For i = 1 To 3
            .Controls(GetControlIndex("winQuest", "lblQuestDescription" & i)).text = ""
        Next i

        For i = 1 To MAX_QUESTS_ITEMS
            .Controls(GetControlIndex("winQuest", "picReward" & i)).visible = False
        Next i

    End With
End Sub

Public Sub LoadQuestLogBox(ByVal QuestSelected As Byte)
    Dim questNum As Long, i As Long
    Dim QuestString As String

    ' Clear window first
    ClearQuestLogBox

    With Windows(GetWindowIndex("winQuest"))

        questNum = FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & QuestSelected)).text)

        If questNum = 0 Then Exit Sub

        'Descrição da quest
        QuestString = Trim$(Quest(questNum).Speech)
        .Controls(GetControlIndex("winQuest", "lblQuestDescription1")).text = QuestString

        'Objetivo da Task
        If Player(MyIndex).PlayerQuest(questNum).ActualTask > 0 Then
            QuestString = GetQuestObjetiveCurrent(questNum) & GetQuestObjetives(questNum)
        End If

        .Controls(GetControlIndex("winQuest", "lblQuestDescription2")).text = QuestString

        'Recompensa da quest
        QuestString = "Exp: " & Quest(questNum).RewardExp & vbNewLine & "Level(s): " & Quest(questNum).RewardLevel
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(questNum).RewardItem(i).Item > 0 Then
                .Controls(GetControlIndex("winQuest", "picReward" & i)).Value = Quest(questNum).RewardItem(i).Item
                .Controls(GetControlIndex("winQuest", "picReward" & i)).visible = True
            End If
        Next i
        .Controls(GetControlIndex("winQuest", "lblQuestDescription3")).text = QuestString

    End With
End Sub

Public Sub SelectLastQuest(ByVal QuestID As Integer)
    Dim i As Integer
    
    With Windows(GetWindowIndex("winQuest"))
    
    For i = 1 To MAX_QUESTS_WINDOW
        If QuestID = FindQuestIndex(.Controls(GetControlIndex("winQuest", "lblList" & i)).text) Then
            QuestSelect = i
            Exit Sub
        End If
    Next i
    
    End With
    
End Sub
