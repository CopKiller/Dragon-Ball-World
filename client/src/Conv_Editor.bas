Attribute VB_Name = "Conv_Editor"
Option Explicit

Public Conv_Changed(1 To MAX_CONVS) As Boolean

' /////////////////
' // Conv Editor //
' /////////////////
Public Sub ConvEditorInit()
    Dim i As Long, n As Long

    If frmEditor_Conv.visible = False Then Exit Sub
    EditorIndex = frmEditor_Conv.lstIndex.ListIndex + 1

    With frmEditor_Conv
        .txtName.text = Trim$(Conversation(EditorIndex).Name)

        If Conversation(EditorIndex).chatCount = 0 Then
            Conversation(EditorIndex).chatCount = 1
            Call InitConversationMode(EditorIndex, ClearAndRedimensionEmpty)
        End If

        Call ConvRealoadData(EditorIndex, .scrlConv.Value)

        Call ConvReloadEventOptions(EditorIndex, .scrlConv.Value)
    End With

    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvRealoadData(ByVal EditorIndex As Long, ByVal CurConv As Long)
    Dim i As Long, n As Long

    With frmEditor_Conv
        For n = 1 To 4
            .cmbReply(n).Clear
            .cmbReply(n).AddItem "None"

            For i = 1 To Conversation(EditorIndex).chatCount
                .cmbReply(n).AddItem i
            Next
        Next

        If Conversation(EditorIndex).chatCount = 0 Then Conversation(EditorIndex).chatCount = 1
        .scrlChatCount = Conversation(EditorIndex).chatCount
        .scrlConv.max = Conversation(EditorIndex).chatCount
        
        If CurConv > .scrlConv.max Then CurConv = .scrlConv.max
        .scrlConv.Value = CurConv
        .txtConv = Trim$(Conversation(EditorIndex).Conv(CurConv).Talk)

        For i = 1 To 4
            .txtReply(i).text = Trim$(Conversation(EditorIndex).Conv(CurConv).rText(i))
            .cmbReply(i).ListIndex = Conversation(EditorIndex).Conv(CurConv).rTarget(i)
        Next
    End With
End Sub

Public Sub ConvReloadEventOptions(ByVal EditorIndex As Long, ByVal CurConv As Long)
    Dim i As Long
    
    With Conversation(EditorIndex).Conv(CurConv)
        frmEditor_Conv.cmbEvent.ListIndex = .EventType

        If frmEditor_Conv.cmbEvent.ListIndex = EventType.Event_OpenShop Then
            ' build EventNum combo
            frmEditor_Conv.cmbEventNum.Clear
            frmEditor_Conv.cmbEventNum.AddItem "None"
            frmEditor_Conv.cmbEventNum.visible = True

            For i = 1 To MAX_SHOPS
                frmEditor_Conv.cmbEventNum.AddItem i & ": " & Trim$(Shop(i).Name)
            Next

            frmEditor_Conv.cmbEventNum.ListIndex = .EventNum
        ElseIf frmEditor_Conv.cmbEvent.ListIndex = EventType.Event_GiveQuest Then
            ' build EventNum combo
            frmEditor_Conv.cmbEventNum.Clear
            frmEditor_Conv.cmbEventNum.AddItem "None"
            frmEditor_Conv.cmbEventNum.visible = True

            For i = 1 To MAX_QUESTS
                frmEditor_Conv.cmbEventNum.AddItem i & ": " & Trim$(Quest(i).Name)
            Next

            frmEditor_Conv.cmbEventNum.ListIndex = .EventNum
        Else
            frmEditor_Conv.cmbEventNum.Clear
            frmEditor_Conv.cmbEventNum.visible = False
        End If
    End With
End Sub

Public Sub ConvEditorOk()
    Dim i As Long

    For i = 1 To MAX_CONVS

        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If

    Next

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
End Sub

Public Sub ConvEditorCancel()
    Editor = 0
    Unload frmEditor_Conv
    ClearChanged_Conv
    ClearConvs
    SendRequestConvs
End Sub

Public Sub ClearChanged_Conv()
    ZeroMemory Conv_Changed(1), MAX_CONVS * 2 ' 2 = boolean length
End Sub
