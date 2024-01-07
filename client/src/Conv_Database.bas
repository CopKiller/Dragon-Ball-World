Attribute VB_Name = "Conv_Database"
Option Explicit

Public Enum EditorConversationMode
    ClearAndRedimensionEmpty
    AddRedimensionToChat
End Enum
Public Sub InitConversationMode(ByVal EditorIndex As Long, ByVal conversationMode As EditorConversationMode, _
                                Optional ByVal valueAttribute As Long = 1)
    Dim i As Long
    
    ' Tratamento de toda a limpeza e redimensionamento no âmbito das conversas
    Select Case conversationMode
        Case EditorConversationMode.ClearAndRedimensionEmpty
            Conversation(EditorIndex) = EmptyConv
            Conversation(EditorIndex).Name = vbNullString
            ReDim Conversation(EditorIndex).Conv(valueAttribute)
            Conversation(EditorIndex).Conv(valueAttribute).Talk = vbNullString
            For i = 1 To 4
                Conversation(EditorIndex).Conv(valueAttribute).rText(i) = vbNullString
            Next i
        Case EditorConversationMode.AddRedimensionToChat
            If valueAttribute > UBound(Conversation(EditorIndex).Conv) Then
                ReDim Preserve Conversation(EditorIndex).Conv(LBound(Conversation(EditorIndex).Conv) To UBound(Conversation(EditorIndex).Conv) + 1) As ConvRec
                Conversation(EditorIndex).Conv(UBound(Conversation(EditorIndex).Conv)).Talk = vbNullString
                For i = 1 To 4
                    Conversation(EditorIndex).Conv(UBound(Conversation(EditorIndex).Conv)).rText(i) = vbNullString
                Next i
            ElseIf valueAttribute = UBound(Conversation(EditorIndex).Conv) Then
                ' Evita saída prematura
            Else
                ' Redimensiona sem adicionar nova dimensão
                ReDim Preserve Conversation(EditorIndex).Conv(LBound(Conversation(EditorIndex).Conv) To UBound(Conversation(EditorIndex).Conv) - 1) As ConvRec
            End If
    End Select
End Sub

Public Sub ClearConvs()
    Dim i As Long
    For i = 1 To MAX_CONVS
        Call InitConversationMode(i, ClearAndRedimensionEmpty)
    Next
End Sub
