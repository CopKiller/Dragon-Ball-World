Attribute VB_Name = "Player_Actions"
Option Explicit

Public Sub InitChat(ByVal index As Long, ByVal mapnum As Long, ByVal mapNpcNum As Long, Optional ByVal remoteChat As Boolean = False)
    Dim npcNum As Long
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
    ' check if we can chat
    If Npc(npcNum).Conv = 0 Then Exit Sub
    If Len(Trim$(Conversation(Npc(npcNum).Conv).Name)) = 0 Then Exit Sub
    
    If Not remoteChat Then
        With MapNpc(mapnum).Npc(mapNpcNum)
            .c_inChatWith = index
            .c_lastDir = .Dir
            If GetPlayerY(index) = .y - 1 Then
                .Dir = DIR_UP
            ElseIf GetPlayerY(index) = .y + 1 Then
                .Dir = DIR_DOWN
            ElseIf GetPlayerX(index) = .x - 1 Then
                .Dir = DIR_LEFT
            ElseIf GetPlayerX(index) = .x + 1 Then
                .Dir = DIR_RIGHT
            End If
            ' send NPC's dir to the map
            NpcDir mapnum, mapNpcNum, .Dir
        End With
    End If
    
    ' Set chat value to Npc
    TempPlayer(index).inChatWith = npcNum
    TempPlayer(index).c_mapNpcNum = mapNpcNum
    TempPlayer(index).c_mapNum = mapnum
    ' set to the root chat
    TempPlayer(index).curChat = 1
    ' send the root chat
    sendChat index
End Sub

Public Sub chatOption(ByVal index As Long, ByVal chatOption As Long)
    Dim exitChat As Boolean
    Dim convNum As Long
    Dim curChat As Long
    
    If TempPlayer(index).inChatWith = 0 Then Exit Sub
    
    convNum = Npc(TempPlayer(index).inChatWith).Conv
    curChat = TempPlayer(index).curChat
    
    exitChat = False
    
    ' follow route
    If Conversation(convNum).Conv(curChat).rTarget(chatOption) = 0 Then
        exitChat = True
    Else
        TempPlayer(index).curChat = Conversation(convNum).Conv(curChat).rTarget(chatOption)
    End If
    
    ' if exiting chat, clear temp values
    If exitChat Then
        TempPlayer(index).inChatWith = 0
        TempPlayer(index).curChat = 0
        ' send chat update
        sendChat index
        ' send npc dir
        With MapNpc(TempPlayer(index).c_mapNum).Npc(TempPlayer(index).c_mapNpcNum)
            If .c_inChatWith = index Then
                .c_inChatWith = 0
                .Dir = .c_lastDir
                NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .Dir
            End If
        End With
        ' clear last of data
        TempPlayer(index).c_mapNpcNum = 0
        TempPlayer(index).c_mapNum = 0
        ' exit out early so we don't send chat update twice
        Exit Sub
    End If
    
    ' send update to the client
    sendChat index
End Sub

Public Sub sendChat(ByVal index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim mainText As String
    Dim optText(1 To 4) As String
    'Dim P_GENDER As String
    'Dim P_NAME As String
    'Dim P_CLASS As String
    Dim i As Long

    If TempPlayer(index).inChatWith > 0 Then
        convNum = Npc(TempPlayer(index).inChatWith).Conv
        curChat = TempPlayer(index).curChat

        ' check for unique events and trigger them early
        If Conversation(convNum).Conv(curChat).EventType > 0 Then
            Select Case Conversation(convNum).Conv(curChat).EventType
            Case 1    ' Open Shop
                If Conversation(convNum).Conv(curChat).EventNum > 0 Then    ' shop exists?
                    SendOpenShop index, Conversation(convNum).Conv(curChat).EventNum
                    TempPlayer(index).InShop = Conversation(convNum).Conv(curChat).EventNum    ' stops movement and the like
                End If
                ' exit out early so we don't send chat update twice
                ClosePlayerChat index
                Exit Sub
            Case 2    ' Open Bank
                SendBank index
                TempPlayer(index).InBank = True
                ' exit out early'
                ClosePlayerChat index
                Exit Sub
            Case 3    ' Give Quest
                If Conversation(convNum).Conv(curChat).EventNum > 0 Then
                    If QuestInProgress(index, Conversation(convNum).Conv(curChat).EventNum) Then
                        'if the quest is in progress show the meanwhile message (speech2)
                        mainText = Trim$(Quest(Conversation(convNum).Conv(curChat).EventNum).Task(Player(index).PlayerQuest(Conversation(convNum).Conv(curChat).EventNum).ActualTask).TaskLog)
                        SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    If CanStartQuest(index, Conversation(convNum).Conv(curChat).EventNum) Then
                        'if can start show the request message (speech1)
                        StartQuest index, Conversation(convNum).Conv(curChat).EventNum, 1
                        mainText = Trim$(Quest(Conversation(convNum).Conv(curChat).EventNum).QuestLog)
                        SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    mainText = "Voce nao cumpre algum requisito, verifique o log no chat!"
                    SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                    Exit Sub
                End If
            Case 4    ' unique script
            Case 5    ' Dar Dica
            Case 6    ' Dar Quest
            Case 7    ' Iniciar Viagem
            Case 8    ' Abrir mix
            End Select
        End If

Continue:
        ' cache player's details
        'If Player(index).Sex = SEX_MALE Then
        '    P_GENDER = "man"
        'Else
        '    P_GENDER = "woman"
        'End If
        'P_NAME = Trim$(Player(index).Name)
        'P_CLASS = Trim$(Class(Player(index).Class).Name)

        mainText = Conversation(convNum).Conv(curChat).Talk
        For i = 1 To 4
            optText(i) = Conversation(convNum).Conv(curChat).rText(i)
        Next
    End If

    SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
    Exit Sub
End Sub

Public Sub ClosePlayerChat(ByVal index As Long)
    ' exit the chat
    TempPlayer(index).inChatWith = 0
    TempPlayer(index).curChat = 0
    ' send chat update
    sendChat index
    ' send npc dir
    With MapNpc(TempPlayer(index).c_mapNum).Npc(TempPlayer(index).c_mapNpcNum)
        If .c_inChatWith = index Then
            .c_inChatWith = 0
            .Dir = .c_lastDir
            NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .Dir
        End If
    End With
    ' clear last of data
    TempPlayer(index).c_mapNpcNum = 0
    TempPlayer(index).c_mapNum = 0
    Exit Sub
End Sub

Sub UseChar(ByVal index As Long, ByVal charNum As Long)
    If Not IsPlaying(index) Then
        Call LoadPlayer(index, charNum)
        TempPlayer(index).charNum = charNum
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendConvs(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendPlayerVariables(index)
    Call SendQuests(index)
    Call SendPlayerQuests(index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", White)
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    If GetPlayerMap(index) > 0 And GetPlayerMap(index) <= MAX_MAPS Then
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            SendResourceCacheTo index, i
        Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    
    ' tell them to do the damn tutorial
    If Player(index).TutorialState = 0 Then SendStartTutorial index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(index) <= 0 Or GetPlayerMap(index) > MAX_MAPS Then Exit Sub
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index

        ' save and clear data.
        Call SavePlayer(index)

        ' Send a global message that he/she left
        Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", White)

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
    Call ClearAccount(index)
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MapData.maxX Then x = Map(mapnum).MapData.maxX
    If y > Map(mapnum).MapData.maxY Then y = Map(mapnum).MapData.maxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    SendCheckForMap index, mapnum
End Sub

Public Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long, x As Long, y As Long, moved As Byte, MovedSoFar As Boolean, newMapX As Byte, newMapY As Byte
    Dim TileType As Long, vitalType As Long, colour As Long, Amount As Long, canMoveResult As Long, i As Long
    Dim fonte As fonts

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    moved = NO
    mapnum = GetPlayerMap(index)
    
    If mapnum = 0 Then Exit Sub
    
    ' check if they're casting a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        SendCancelAnimation index
        TempPlayer(index).spellBuffer.Spell = 0
        TempPlayer(index).spellBuffer.target = 0
        TempPlayer(index).spellBuffer.Timer = 0
        TempPlayer(index).spellBuffer.tType = 0
        
        Call SendClearSpellBufferTo(index)
        Call ClearPlayerFrameToMapBut(index)
    End If
    
    ' check directions
    canMoveResult = CanMove(index, Dir)
    If canMoveResult = 1 Then
        Select Case Dir
            Case DIR_UP
                Call SetPlayerY(index, GetPlayerY(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN
                Call SetPlayerY(index, GetPlayerY(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_LEFT
                Call SetPlayerX(index, GetPlayerX(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_RIGHT
                Call SetPlayerX(index, GetPlayerX(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_UP_LEFT
                Call SetPlayerY(index, GetPlayerY(index) - 1)
                Call SetPlayerX(index, GetPlayerX(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_UP_RIGHT
                Call SetPlayerY(index, GetPlayerY(index) - 1)
                Call SetPlayerX(index, GetPlayerX(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN_LEFT
                Call SetPlayerY(index, GetPlayerY(index) + 1)
                Call SetPlayerX(index, GetPlayerX(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN_RIGHT
                Call SetPlayerY(index, GetPlayerY(index) + 1)
                Call SetPlayerX(index, GetPlayerX(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
                
        End Select
    End If
    
    With Map(GetPlayerMap(index)).TileData.Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
            Call PlayerWarp(index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, x, y, 1
                'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            vitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(index, vitalType) = GetPlayerMaxVital(index, vitalType) Then
                If vitalType = Vitals.HP Then
                    colour = BrightGreen
                    fonte = health
                Else
                    colour = BrightBlue
                    fonte = energy
                End If
                SendActionMsg GetPlayerMap(index), "+" & Amount, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, fonte, index
                SetPlayerVital index, vitalType, GetPlayerVital(index, vitalType) + Amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(index, vitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, fonts.Damage, index
            If GetPlayerVital(index, HP) - Amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            moved = YES
        End If
    End With

    ' They tried to hack
    If moved = NO And canMoveResult <> 2 Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If
End Sub

Public Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal direction As Long)
    If direction < DIR_UP Or direction > DIR_DOWN_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.maxY Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MapData.maxX Then Exit Sub
        Case DIR_UP_LEFT
            If GetPlayerY(index) = 0 Then Exit Sub
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_UP_RIGHT
            If GetPlayerY(index) = 0 Then Exit Sub
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MapData.maxX Then Exit Sub
        Case DIR_DOWN_LEFT
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.maxY Then Exit Sub
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN_RIGHT
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.maxY Then Exit Sub
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MapData.maxX Then Exit Sub
    End Select
    
    PlayerMove index, direction, movement, True
End Sub

Public Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Public Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Player(index).Inv(i).Bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Player(index).Inv(invSlot).Bound = 0
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If Item(ItemNum).BindType = 1 Then ' bind on pickup
                Player(index).Inv(i).Bound = 1
                PlayerMsg index, "This item is now bound to your soul.", BrightRed
            Else
                Player(index).Inv(i).Bound = 0
            End If
        Else
            Player(index).Inv(i).Bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Public Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i).Spell = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)
    
    If mapnum = 0 Then Exit Sub

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) And (MapItem(mapnum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).Num)
                            
                            ' check tasks
                            Call CheckTasks(index, QUEST_TYPE_GOGATHER, MapItem(mapnum, i).Num)
                            
                            If Item(GetPlayerInvItemNum(index, n)).Type <> ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(index).Inv(n).Bound = 0
                            If Item(GetPlayerInvItemNum(index, n)).BindType = 1 Or MapItem(mapnum, i).Bound Then
                                Player(index).Inv(n).Bound = 1
                                If Not Trim$(MapItem(mapnum, i).playerName) = Trim$(GetPlayerName(index)) Then
                                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If
                            

                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), alert
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If Item(GetPlayerInvItemNum(index, invNum)).BindType > 0 Then
                If Player(index).Inv(invNum).Bound = 1 Then
                    PlayerMsg index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
                If Player(index).Inv(invNum).Bound > 0 Then
                    MapItem(GetPlayerMap(index), i).Bound = True
                Else
                    MapItem(GetPlayerMap(index), i).Bound = False
                End If

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                        Player(index).Inv(invNum).Bound = 0
                    Else
                        MapItem(GetPlayerMap(index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                    Player(index).Inv(invNum).Bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn, MapItem(GetPlayerMap(index), i).Bound)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long, Optional ByVal level_count As Long)
    Dim i As Long, PontosPorLevel As Byte
    Dim expRollover As Long

    PontosPorLevel = 3

    ' Caso queira adicionar levels diretamente!
    If level_count > 0 Then
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + level_count) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + (level_count * PontosPorLevel))
        GoTo Continue
    End If

    ' Adiciona level pela experi�ncia, m�todo normal de um rpg
    level_count = 0
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)

        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + PontosPorLevel)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop

Continue:
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
            'Call SendDiscordMsg(Levelup, Index, "has gained " & level_count & " level!")
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
            'Call SendDiscordMsg(Levelup, Index, "has gained " & level_count & " levels!")
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    SendVital index, HP

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            PlayerMapDropItem index, GetPlayerEquipment(index, i), 0
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index)).MapData
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBufferTo(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long

    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).TileData.Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            'If GetPlayerEquipment(index, Weapon) > 0 Then
            ' If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

            If Resource(Resource_index).ToolRequired > 0 Then
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then
                        Damage = Item(GetPlayerEquipment(index, Weapon)).Data2
                    Else
                        PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                        Exit Sub
                    End If
                Else
                    PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
                    Exit Sub
                End If
            Else
                'Damage = GetPlayerDamage(index)
                Damage = 1
            End If

            ' inv space?
            If Resource(Resource_index).ItemReward > 0 Then
                If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                    PlayerMsg index, "You have no inventory space.", BrightRed
                    Exit Sub
                End If
            End If

            ' check if already cut down
            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then

                rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y

                'Damage = Item(GetPlayerEquipment(index, Weapon)).Data2

                ' check if damage is more than health
                If Damage > 0 Then
                    ' cut it down!
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                        SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32), Damage
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1    ' Cut
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                        SendResourceCacheToMap GetPlayerMap(index), Resource_num
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), alert
                        End If
                        ' carry on
                        GiveInvItem index, Resource(Resource_index).ItemReward, 1
                        SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                    Else
                        ' just do the damage
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                        SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32), Damage
                        SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                    End If
                    ' send the sound
                    SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                Else
                    ' too weak
                    SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32), alert
                End If
            Else
                ' send message if it exists
                If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), alert
                End If
            End If

            '  Else
            '      PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
            '  End If

            'Else
            '    PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            'End If
        End If
    End If
End Sub

Public Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal Amount As Long)
    Dim BankSlot As Long, ItemNum As Long

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If Amount < 1 Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub EquipItem(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Dim ItemNum As Long, tempItem As Long
    
    ItemNum = GetPlayerInvItemNum(index, invNum)
    If ItemNum < 0 And ItemNum > MAX_ITEMS Then Exit Sub
    If Not IsPlayerItemRequerimentsOK(index, ItemNum) Then Exit Sub
    
    If GetPlayerEquipment(index, EquipmentSlot) > 0 Then
        tempItem = GetPlayerEquipment(index, EquipmentSlot)
    End If

    Call SetPlayerEquipment(index, ItemNum, EquipmentSlot)
    
    Call PlayerMsg(index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen)
    
    ' tell them if it's soulbound
    If Item(ItemNum).BindType = 2 Then ' BoE
        If Player(index).Inv(invNum).Bound = 0 Then
            PlayerMsg index, "This item is now bound to your soul.", BrightRed
        End If
    End If
    
    Call TakeInvItem(index, ItemNum, 0)

    If tempItem > 0 Then
        If Item(tempItem).BindType > 0 Then
            Call GiveInvItem(index, tempItem, 0, True) ' give back the stored item
            tempItem = 0
        Else
            Call GiveInvItem(index, tempItem, 0)
            tempItem = 0
        End If
    End If

    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    
    ' send vitals
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then Call SendPartyVitals(TempPlayer(index).inParty, index)
    
    ' send the sound
    Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
    Dim n As Long, i As Long
    Dim tempItem As Long
    Dim x As Long, y As Long
    Dim ItemNum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        ItemNum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_WEAPON To ITEM_TYPE_FEET
                Call EquipItem(index, invNum, Item(ItemNum).Type)
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, health
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, energy
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(ItemNum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, fonts.exp
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.maxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.maxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If ItemNum = Map(GetPlayerMap(index)).TileData.Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, x, y, 1
                        'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).TileData.Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, ItemNum, 0)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_UNIQUE
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Go through with it
                Unique_Item index, ItemNum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(ItemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                    
                        ' make sure they don't already know it
                        For i = 1 To MAX_PLAYER_SPELLS
                            If Player(index).Spell(i).Spell > 0 Then
                                If Player(index).Spell(i).Spell = n Then
                                    PlayerMsg index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                                If Spell(Player(index).Spell(i).Spell).UniqueIndex = Spell(n).UniqueIndex Then
                                    PlayerMsg index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        Next
                    
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq


                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Player(index).Spell(i).Spell = n
                                    Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, ItemNum, 0)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells index
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_FOOD
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' make sure they're not in combat
                If TempPlayer(index).stopRegen Then
                    PlayerMsg index, "You cannot eat whilst in combat.", BrightRed
                    Exit Sub
                End If
                
                ' make sure not full hp
                x = Item(ItemNum).HPorSP
                If Player(index).Vital(x) >= GetPlayerMaxVital(index, x) Then
                    PlayerMsg index, "You don't need to eat this at the moment.", BrightRed
                    Exit Sub
                End If
                
                ' set the player's food
                If Item(ItemNum).HPorSP = 2 Then 'mp
                    If Not TempPlayer(index).foodItem(Vitals.MP) = ItemNum Then
                        TempPlayer(index).foodItem(Vitals.MP) = ItemNum
                        TempPlayer(index).foodTick(Vitals.MP) = 0
                        TempPlayer(index).foodTimer(Vitals.MP) = GetTickCount
                    Else
                        PlayerMsg index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                Else ' hp
                    If Not TempPlayer(index).foodItem(Vitals.HP) = ItemNum Then
                        TempPlayer(index).foodItem(Vitals.HP) = ItemNum
                        TempPlayer(index).foodTick(Vitals.HP) = 0
                        TempPlayer(index).foodTimer(Vitals.HP) = GetTickCount
                    Else
                        PlayerMsg index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' take the item
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
        End Select
    End If
End Sub
