Attribute VB_Name = "Quest_Editor"
Option Explicit

Public Sub SendRequestEditQuest()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditQuest
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong CSaveQuest
    buffer.WriteLong QuestNum
    buffer.WriteBytes QuestData
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

Sub SendRequestQuests()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestQuests
    SendData buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub HandleQuestEditor()
    Dim i As Long

    With frmEditor_Quest
        Editor = EDITOR_QUEST
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

End Sub

Public Sub QuestEditorInit()
    Dim i As Long

    If frmEditor_Quest.visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(Quest(EditorIndex).Name)

        .optRepeat(Quest(EditorIndex).Repeat).Value = True
        .txtSegs = Quest(EditorIndex).Time

        .txtQuestLog = Trim$(Quest(EditorIndex).QuestLog)
        .txtSpeech.text = Trim$(Quest(EditorIndex).Speech)

        .scrlReqLevel.Value = Quest(EditorIndex).RequiredLevel
        .scrlReqQuest.Value = Quest(EditorIndex).RequiredQuest
        For i = 1 To 5
            .scrlReqClass.Value = Quest(EditorIndex).RequiredClass(i)
        Next

        .txtExp.text = Quest(EditorIndex).RewardExp
        .txtLevel.text = Quest(EditorIndex).RewardLevel

        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        UpdateQuestClass

        '/Alatar v1.2

        'load task nº1
        .scrlTotalTasks.Value = 1
        LoadTask EditorIndex, 1

    End With

    Quest_Changed(EditorIndex) = True

End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim i As Long

    frmEditor_Quest.lstGiveItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).GiveItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim i As Long

    frmEditor_Quest.lstTakeItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).TakeItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim i As Long

    frmEditor_Quest.lstItemRew.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RewardItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim i As Long

    frmEditor_Quest.lstReqItem.Clear

    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RequiredItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestClass()
    Dim i As Long

    frmEditor_Quest.lstReqClass.Clear

    For i = 1 To 5
        If Quest(EditorIndex).RequiredClass(i) = 0 Then
            frmEditor_Quest.lstReqClass.AddItem "-"
        Else
            frmEditor_Quest.lstReqClass.AddItem Trim$(Trim$(Class(Quest(EditorIndex).RequiredClass(i)).Name))
        End If
    Next
End Sub
'/Alatar v1.2

Public Sub QuestEditorOk()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next

    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest

End Sub

Public Sub QuestEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2    ' 2 = boolean length
End Sub

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = Quest(QuestNum).Task(TaskNum)

    With frmEditor_Quest
        'Load the task type
        .optTask(TaskToLoad.Order).Value = True
        'Load textboxes
        .txtTaskLog.text = "" & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.Value = 0
        .scrlItem.Value = 0
        .scrlMap.Value = 0
        .scrlResource.Value = 0
        .scrlAmount.Value = 0
        .scrlNPC.enabled = False
        .scrlItem.enabled = False
        .scrlMap.enabled = False
        .scrlResource.enabled = False
        .scrlAmount.enabled = False

        ' Quest Timer
        .chkTaskTimer.Value = TaskToLoad.TaskTimer.Active
        .optTaskTimer(TaskToLoad.TaskTimer.TimerType).Value = True
        .txtTaskTimer.text = CLng(TaskToLoad.TaskTimer.Timer)
        .chkTaskTeleport = TaskToLoad.TaskTimer.Teleport
        .txtTaskTeleport.text = CInt(TaskToLoad.TaskTimer.Teleport)
        .optReset(TaskToLoad.TaskTimer.ResetType).Value = True
        .txtTaskTeleport = CInt(TaskToLoad.TaskTimer.MapNum)
        .txtTaskX.text = CByte(TaskToLoad.TaskTimer.X)
        .txtTaskY.text = CByte(TaskToLoad.TaskTimer.Y)
        .txtMsg.text = Trim$(CStr(TaskToLoad.TaskTimer.Msg))

        If TaskToLoad.QuestEnd = True Then
            .chkEnd.Value = 1
        Else
            .chkEnd.Value = 0
        End If

        Select Case TaskToLoad.Order
        Case 0    'Nothing

        Case QUEST_TYPE_GOSLAY    '1
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOGATHER    '2
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOTALK    '3
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC

        Case QUEST_TYPE_GOREACH    '4
            .scrlMap.enabled = True
            .scrlMap.Value = TaskToLoad.Map

        Case QUEST_TYPE_GOGIVE    '5
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC

        Case QUEST_TYPE_GOKILL    '6
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOTRAIN    '7
            .scrlResource.enabled = True
            .scrlResource.Value = TaskToLoad.Resource
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        Case QUEST_TYPE_GOGET    '8
            .scrlNPC.enabled = True
            .scrlNPC.Value = TaskToLoad.NPC
            .scrlItem.enabled = True
            .scrlItem.Value = TaskToLoad.Item
            .scrlAmount.enabled = True
            .scrlAmount.Value = TaskToLoad.Amount

        End Select
    End With
End Sub
