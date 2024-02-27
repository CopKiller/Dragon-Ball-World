Attribute VB_Name = "Quest_UDT"
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10    'Alatar v1.2

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3
Public Const QUEST_COMPLETED_DIARY As Byte = 4
Public Const QUEST_COMPLETED_TIME As Byte = 5

Public Const TASK_DEFEAT_LENGTH As Byte = 100

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    Value As Long
End Type
'/Alatar v1.2

Public Type TaskTimerRec
    Active As Byte            ' Is Active?
    TimerType As Byte         ' 0=Days; 1=Hours; 2=Minutes; 3=Seconds.
    timer As Long             ' Time with /\

    Teleport As Byte          ' Teleport cannot end task in time.
    mapnum As Integer         ' Map Number to teleport /\
    ResetType As Byte         ' 0=Resetar Task ; 1=Resetar Quest.
    X As Byte
    Y As Byte

    Msg As String * TASK_DEFEAT_LENGTH
End Type

Public Type TaskRec
    Order As Byte
    Npc As Integer
    Item As Integer
    Map As Integer
    Resource As Integer
    Amount As Long
    TaskLog As String * 150
    QuestEnd As Boolean

    ' Task Timer
    TaskTimer As TaskTimerRec
End Type

Public Type QuestRec
    'Alatar v1.2
    Name As String * NAME_LENGTH
    Repeat As Byte
    Time As Long
    QuestLog As String * 100
    Speech As String * 200
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec

    RequiredLevel As Integer
    RequiredQuest As Integer
    RequiredClass(1 To 5) As Integer
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec

    RewardExp As Long
    RewardLevel As Integer
    RewardSpell As Integer
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec

    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2

End Type

Public Type PlayerQuestRec
    Status As Byte
    ActualTask As Byte
    CurrentCount As Long    'Used to handle the Amount property
    Data As String * 19    ' Salva o now que tem 19 dígitos, pra usar como comparação na hora de iniciar novamente a quest

    TaskTimer As TaskTimerRec
End Type

