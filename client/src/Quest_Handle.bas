Attribute VB_Name = "Quest_Handle"
Option Explicit

Public Sub HandleUpdateQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    'zlib
    buffer.DecompressBuffer

    n = buffer.ReadLong
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
End Sub

Public Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, QuestNum As Long, QSelected As Integer

    Set buffer = New clsBuffer

    buffer.WriteBytes data()
    
    ' Recebe se começou a quest e seleciona ela na lista
    QSelected = buffer.ReadInteger

    For i = 1 To MAX_QUESTS
        QuestNum = buffer.ReadLong

        If QuestNum > 0 Then
            Player(MyIndex).PlayerQuest(QuestNum).status = buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).ActualTask = buffer.ReadLong
            Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = buffer.ReadLong

            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = buffer.ReadByte
            Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = buffer.ReadLong

            QuestTimeToFinish = vbNullString
            QuestNameToFinish = vbNullString
            QuestSelect = QuestNum
        End If
    Next

    RefreshQuestWindow
    
    If QSelected > 0 Then
        SelectLastQuest QSelected
    End If

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleQuestMessage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, QuestNum As Long, header As String, saycolour As Long
    Dim message As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    message = Trim$(buffer.ReadString)
    saycolour = buffer.ReadLong
    header = buffer.ReadString

    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)

    AddText ColourChar & GetColStr(Gold) & header & Trim$(Quest(QuestNum).Name) & " : " & ColourChar & GetColStr(saycolour) & message, Grey, , ChatChannel.chQuest

    Set buffer = Nothing
End Sub

Public Sub HandleQuestCancel(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim QuestNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    QuestNum = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).status = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).ActualTask = buffer.ReadLong
    Player(MyIndex).PlayerQuest(QuestNum).CurrentCount = buffer.ReadLong

    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Active = buffer.ReadByte
    Player(MyIndex).PlayerQuest(QuestNum).TaskTimer.Timer = buffer.ReadLong

    QuestTimeToFinish = vbNullString
    QuestNameToFinish = vbNullString

    RefreshQuestWindow

    Set buffer = Nothing
End Sub


Public Sub UpdateOffers(Index_Offer)
    Dim i As Long
    
    If Index_Offer <> Offer_HighIndex Then
        For i = Index_Offer To MAX_OFFER
            If i <> Offer_HighIndex And i < MAX_OFFER Then
                inOffer(i) = inOffer(i + 1)
                inOfferType(i) = inOfferType(i + 1)
                inOfferInvite(i) = inOfferInvite(i + 1)
            Else
                inOffer(i) = 0
                inOfferType(i) = 0
                inOfferInvite(i) = 0
            End If
        Next
    Else
        inOffer(Offer_HighIndex) = 0
        inOfferType(Offer_HighIndex) = 0
        inOfferInvite(Offer_HighIndex) = 0
    End If
    
    Call SetOfferHighIndex
    If Offer_HighIndex > 0 Then
        For i = 1 To Offer_HighIndex
            Call UpdateWindowOffer(i)
        Next
    Else
        Call UpdateWindowOffer(0)
    End If
End Sub

Function FindOpenOfferSlot() As Long
    Dim i As Long
    FindOpenOfferSlot = 0

    For i = 1 To MAX_OFFER
        If inOffer(i) = 0 Then
            FindOpenOfferSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub SetOfferHighIndex()
    Dim i As Integer
    Dim X As Integer
    
    For i = 0 To MAX_OFFER
        X = MAX_OFFER - i
        If X > 0 Then
            If inOffer(X) <> 0 Then
                Offer_HighIndex = X
            Exit Sub
            End If
        End If

    Next i

    Offer_HighIndex = 0
End Sub

