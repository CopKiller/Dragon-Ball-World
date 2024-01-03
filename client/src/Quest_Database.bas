Attribute VB_Name = "Quest_Database"
Option Explicit

Sub ClearQuest(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).QuestLog = vbNullString
    Quest(Index).Speech = vbNullString
    
    For i = 1 To MAX_TASKS
        Quest(Index).Task(i).TaskLog = vbNullString
    Next i
End Sub

Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub
