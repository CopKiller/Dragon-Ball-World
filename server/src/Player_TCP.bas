Attribute VB_Name = "Player_TCP"
Public Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteLong GetPlayerClass(Index)
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    PlayerData = Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Function

Public Sub SendPlayerData(ByVal Index As Long)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Public Sub SendPlayerData_Party(partynum As Long)
    Dim i As Long, X As Long
    ' loop through all the party members
    For i = 1 To Party(partynum).MemberCount
        For X = 1 To Party(partynum).MemberCount
            SendDataTo Party(partynum).Member(X), PlayerData(Party(partynum).Member(i))
        Next
    Next
End Sub

Public Sub SendPlayerVariables(ByVal Index As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerVariables
    For i = 1 To MAX_BYTE
        Buffer.WriteLong Player(Index).Variable(i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendAttack(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong Index
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTarget(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).Target
    Buffer.WriteLong TempPlayer(Index).TargetType
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendHotbar(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).Hotbar(i).Slot
        Buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerXYToMap(ByVal Index As Long, Optional ByVal ImpactedDir As Byte = 0)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteByte ImpactedDir
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendMapNpcXY(ByVal Index As Long, ByVal MapNum As Long, Optional ByVal ImpactedDir As Byte = 0)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapNpcDataXY

    Buffer.WriteLong Index
    Buffer.WriteLong MapNpc(MapNum).Npc(Index).X
    Buffer.WriteLong MapNpc(MapNum).Npc(Index).Y
    Buffer.WriteLong MapNpc(MapNum).Npc(Index).Dir
    Buffer.WriteByte ImpactedDir

    SendDataToMap MapNum, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
    
    ' check if they're in a party
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
End Sub

Public Sub SendEXP(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStats(ByVal Index As Long)
    Dim i As Long
    Dim packet As String
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteByte Player(Index).Inv(i).Bound
    Next

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, invSlot)
    Buffer.WriteByte Player(Index).Inv(invSlot).Bound
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(Index, Shield)
    Buffer.WriteLong GetPlayerEquipment(Index, Pants)
    Buffer.WriteLong GetPlayerEquipment(Index, Feet)
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong Player(Index).Spell(i).Spell
        Buffer.WriteLong Player(Index).Spell(i).Uses
    Next

    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim tradeTarget As Long
    Dim totalWorth As Long, multiplier As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyUpdate(ByVal partynum As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partynum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partynum).Member(i)
    Next
    Buffer.WriteLong Party(partynum).MemberCount
    
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyUpdateTo(ByVal Index As Long)
    Dim Buffer As clsBuffer, i As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partynum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partynum).Member(i)
        Next
        Buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyVitals(ByVal partynum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, i)
        Buffer.WriteLong Player(Index).Vital(i)
    Next
    
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendDataToParty(ByVal partynum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Party(partynum).MemberCount
        If Party(partynum).Member(i) > 0 Then
            Call SendDataTo(Party(partynum).Member(i), Data)
        End If
    Next
End Sub

Public Sub SendBlood(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStunned(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Player(Index).Bank(i).Num
        Buffer.WriteLong Player(Index).Bank(i).Value
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Public Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStartTutorial(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStartTutorial
    
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendProjectile(ByVal MapNum As Long, ByVal ProjectileSlot As Long, Optional ByVal IsDirectional As Boolean = False)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SProjectileAttack
    Buffer.WriteLong ProjectileSlot
    Buffer.WriteLong MapProjectile_HighIndex
    With MapProjectile(ProjectileSlot)
        Buffer.WriteLong .Owner
        Buffer.WriteLong .OwnerType
        Buffer.WriteLong .direction
        Buffer.WriteLong .Graphic
        Buffer.WriteByte Spell(.spellNum).IsAoE
        Buffer.WriteLong .Rotate
        Buffer.WriteLong .RotateSpeed
        Buffer.WriteLong .Speed
        Buffer.WriteLong .Duration
        Buffer.WriteLong .X
        Buffer.WriteLong .Y
        Buffer.WriteLong .tX
        Buffer.WriteLong .tY
        Buffer.WriteByte IsDirectional
        For i = 1 To 4
            Buffer.WriteLong .ProjectileOffset(i).X
            Buffer.WriteLong .ProjectileOffset(i).Y
        Next
    End With
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
