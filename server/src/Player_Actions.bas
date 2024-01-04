Attribute VB_Name = "Player_Actions"
Option Explicit

Public Sub InitChat(ByVal Index As Long, ByVal mapnum As Long, ByVal mapNpcNum As Long, Optional ByVal remoteChat As Boolean = False)
    Dim npcNum As Long
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
    ' check if we can chat
    If Npc(npcNum).Conv = 0 Then Exit Sub
    If Len(Trim$(Conv(Npc(npcNum).Conv).Name)) = 0 Then Exit Sub
    
    If Not remoteChat Then
        With MapNpc(mapnum).Npc(mapNpcNum)
            .c_inChatWith = Index
            .c_lastDir = .Dir
            If GetPlayerY(Index) = .y - 1 Then
                .Dir = DIR_UP
            ElseIf GetPlayerY(Index) = .y + 1 Then
                .Dir = DIR_DOWN
            ElseIf GetPlayerX(Index) = .x - 1 Then
                .Dir = DIR_LEFT
            ElseIf GetPlayerX(Index) = .x + 1 Then
                .Dir = DIR_RIGHT
            End If
            ' send NPC's dir to the map
            NpcDir mapnum, mapNpcNum, .Dir
        End With
    End If
    
    ' Set chat value to Npc
    TempPlayer(Index).inChatWith = npcNum
    TempPlayer(Index).c_mapNpcNum = mapNpcNum
    TempPlayer(Index).c_mapNum = mapnum
    ' set to the root chat
    TempPlayer(Index).curChat = 1
    ' send the root chat
    sendChat Index
End Sub

Public Sub chatOption(ByVal Index As Long, ByVal chatOption As Long)
    Dim exitChat As Boolean
    Dim convNum As Long
    Dim curChat As Long
    
    If TempPlayer(Index).inChatWith = 0 Then Exit Sub
    
    convNum = Npc(TempPlayer(Index).inChatWith).Conv
    curChat = TempPlayer(Index).curChat
    
    exitChat = False
    
    ' follow route
    If Conv(convNum).Conv(curChat).rTarget(chatOption) = 0 Then
        exitChat = True
    Else
        TempPlayer(Index).curChat = Conv(convNum).Conv(curChat).rTarget(chatOption)
    End If
    
    ' if exiting chat, clear temp values
    If exitChat Then
        TempPlayer(Index).inChatWith = 0
        TempPlayer(Index).curChat = 0
        ' send chat update
        sendChat Index
        ' send npc dir
        With MapNpc(TempPlayer(Index).c_mapNum).Npc(TempPlayer(Index).c_mapNpcNum)
            If .c_inChatWith = Index Then
                .c_inChatWith = 0
                .Dir = .c_lastDir
                NpcDir TempPlayer(Index).c_mapNum, TempPlayer(Index).c_mapNpcNum, .Dir
            End If
        End With
        ' clear last of data
        TempPlayer(Index).c_mapNpcNum = 0
        TempPlayer(Index).c_mapNum = 0
        ' exit out early so we don't send chat update twice
        Exit Sub
    End If
    
    ' send update to the client
    sendChat Index
End Sub

Public Sub sendChat(ByVal Index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim mainText As String
    Dim optText(1 To 4) As String
    'Dim P_GENDER As String
    'Dim P_NAME As String
    'Dim P_CLASS As String
    Dim I As Long

    If TempPlayer(Index).inChatWith > 0 Then
        convNum = Npc(TempPlayer(Index).inChatWith).Conv
        curChat = TempPlayer(Index).curChat

        ' check for unique events and trigger them early
        If Conv(convNum).Conv(curChat).EventType > 0 Then
            Select Case Conv(convNum).Conv(curChat).EventType
            Case 1    ' Open Shop
                If Conv(convNum).Conv(curChat).EventNum > 0 Then    ' shop exists?
                    SendOpenShop Index, Conv(convNum).Conv(curChat).EventNum
                    TempPlayer(Index).InShop = Conv(convNum).Conv(curChat).EventNum    ' stops movement and the like
                End If
                ' exit out early so we don't send chat update twice
                ClosePlayerChat Index
                Exit Sub
            Case 2    ' Open Bank
                SendBank Index
                TempPlayer(Index).InBank = True
                ' exit out early'
                ClosePlayerChat Index
                Exit Sub
            Case 3    ' Give Quest
                If Conv(convNum).Conv(curChat).EventNum > 0 Then
                    If QuestInProgress(Index, Conv(convNum).Conv(curChat).EventNum) Then
                        'if the quest is in progress show the meanwhile message (speech2)
                        mainText = Trim$(Quest(Conv(convNum).Conv(curChat).EventNum).Task(Player(Index).PlayerQuest(Conv(convNum).Conv(curChat).EventNum).ActualTask).TaskLog)
                        SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    If CanStartQuest(Index, Conv(convNum).Conv(curChat).EventNum) Then
                        'if can start show the request message (speech1)
                        StartQuest Index, Conv(convNum).Conv(curChat).EventNum, 1
                        mainText = Trim$(Quest(Conv(convNum).Conv(curChat).EventNum).QuestLog)
                        SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
                        Exit Sub
                    End If
                    mainText = "Voce nao cumpre algum requisito, verifique o log no chat!"
                    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
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

        mainText = Conv(convNum).Conv(curChat).Conv
        For I = 1 To 4
            optText(I) = Conv(convNum).Conv(curChat).rText(I)
        Next
    End If

    SendChatUpdate Index, TempPlayer(Index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
    Exit Sub
End Sub

Public Sub ClosePlayerChat(ByVal Index As Long)
    ' exit the chat
    TempPlayer(Index).inChatWith = 0
    TempPlayer(Index).curChat = 0
    ' send chat update
    sendChat Index
    ' send npc dir
    With MapNpc(TempPlayer(Index).c_mapNum).Npc(TempPlayer(Index).c_mapNpcNum)
        If .c_inChatWith = Index Then
            .c_inChatWith = 0
            .Dir = .c_lastDir
            NpcDir TempPlayer(Index).c_mapNum, TempPlayer(Index).c_mapNpcNum, .Dir
        End If
    End With
    ' clear last of data
    TempPlayer(Index).c_mapNpcNum = 0
    TempPlayer(Index).c_mapNum = 0
    Exit Sub
End Sub

Sub UseChar(ByVal Index As Long, ByVal charNum As Long)
    If Not IsPlaying(Index) Then
        Call LoadPlayer(Index, charNum)
        TempPlayer(Index).charNum = charNum
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim I As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendConvs(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendPlayerVariables(Index)
    Call SendQuests(Index)
    Call SendPlayerQuests(Index)
    
    ' send vitals, exp + stats
    For I = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, I)
    Next
    SendEXP Index
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", White)
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    If GetPlayerMap(Index) > 0 And GetPlayerMap(Index) <= MAX_MAPS Then
        For I = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            SendResourceCacheTo Index, I
        Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
    
    ' tell them to do the damn tutorial
    If Player(Index).TutorialState = 0 Then SendStartTutorial Index
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, I As Long
    Dim tradeTarget As Long
    
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(Index) <= 0 Or GetPlayerMap(Index) > MAX_MAPS Then Exit Sub
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For I = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(I).Num = 0
                TempPlayer(tradeTarget).TradeOffer(I).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index

        ' save and clear data.
        Call SavePlayer(Index)

        ' Send a global message that he/she left
        Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", White)

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
    Call ClearAccount(Index)
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim I As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MapData.MaxX Then x = Map(mapnum).MapData.MaxX
    If y > Map(mapnum).MapData.MaxY Then y = Map(mapnum).MapData.MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If GetPlayerMap(I) = mapnum Then
                    SendMapEquipmentTo I, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        ' Regenerate all NPCs' health
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(I).Num > 0 Then
                MapNpc(OldMap).Npc(I).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(I).Num, Vitals.HP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(Index).GettingMap = YES
    SendCheckForMap Index, mapnum
End Sub

Public Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long, x As Long, y As Long, moved As Byte, MovedSoFar As Boolean, newMapX As Byte, newMapY As Byte
    Dim TileType As Long, vitalType As Long, colour As Long, Amount As Long, canMoveResult As Long, I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    moved = NO
    mapnum = GetPlayerMap(Index)
    
    If mapnum = 0 Then Exit Sub
    
    ' check if they're casting a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        SendCancelAnimation Index
        SendClearSpellBuffer Index
        TempPlayer(Index).spellBuffer.Spell = 0
        TempPlayer(Index).spellBuffer.Target = 0
        TempPlayer(Index).spellBuffer.Timer = 0
        TempPlayer(Index).spellBuffer.tType = 0
    End If
    
    ' check directions
    canMoveResult = CanMove(Index, Dir)
    If canMoveResult = 1 Then
        Select Case Dir
            Case DIR_UP
                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN
                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_LEFT
                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_RIGHT
                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_UP_LEFT
                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_UP_RIGHT
                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN_LEFT
                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN_RIGHT
                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                SendPlayerMove Index, movement, sendToSelf
                moved = YES
                
        End Select
    End If
    
    With Map(GetPlayerMap(Index)).TileData.Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(Index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                SendMapKey Index, x, y, 1
                'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, x
                    TempPlayer(Index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            vitalType = .Data1
            Amount = .Data2
            If Not GetPlayerVital(Index, vitalType) = GetPlayerMaxVital(Index, vitalType) Then
                If vitalType = Vitals.HP Then
                    colour = BrightGreen
                Else
                    colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & Amount, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, vitalType, GetPlayerVital(Index, vitalType) + Amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(Index, vitalType)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .Data1
            SendActionMsg GetPlayerMap(Index), "-" & Amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - Amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - Amount
                PlayerMsg Index, "You're injured by a trap.", BrightRed
                Call SendVital(Index, HP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            moved = YES
        End If
    End With

    ' They tried to hack
    If moved = NO And canMoveResult <> 2 Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If
End Sub

Public Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal direction As Long)
    If direction < DIR_UP Or direction > DIR_DOWN_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case direction
        Case DIR_UP
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
        Case DIR_UP_LEFT
            If GetPlayerY(Index) = 0 Then Exit Sub
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_UP_RIGHT
            If GetPlayerY(Index) = 0 Then Exit Sub
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
        Case DIR_DOWN_LEFT
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN_RIGHT
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MapData.MaxY Then Exit Sub
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MapData.MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, direction, movement, True
End Sub

Public Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For I = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                FindOpenInvSlot = I
                Exit Function
            End If

        Next

    End If

    For I = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If

    Next

End Function

Public Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, I) = ItemNum Then
                FindOpenBankSlot = I
                Exit Function
            End If
        Next I

    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, I) = 0 Then
            FindOpenBankSlot = I
            Exit Function
        End If
    Next I

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim I As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, I) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) - ItemVal)
                    Call SendInventoryUpdate(Index, I)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, I, 0)
                Call SetPlayerInvItemValue(Index, I, 0)
                Player(Index).Inv(I).Bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(Index, I)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim I As Long
    Dim n As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invSlot, GetPlayerInvItemValue(Index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, invSlot, 0)
        Call SetPlayerInvItemValue(Index, invSlot, 0)
        Player(Index).Inv(invSlot).Bound = 0
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    I = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If I <> 0 Then
        Call SetPlayerInvItemNum(Index, I, ItemNum)
        Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If Item(ItemNum).BindType = 1 Then ' bind on pickup
                Player(Index).Inv(I).Bound = 1
                PlayerMsg Index, "This item is now bound to your soul.", BrightRed
            Else
                Player(Index).Inv(I).Bound = 0
            End If
        Else
            Player(Index).Inv(I).Bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(Index, I)
        GiveInvItem = True
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Public Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS

        If Player(Index).Spell(I).Spell = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim I As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    mapnum = GetPlayerMap(Index)
    
    If mapnum = 0 Then Exit Sub

    For I = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, I).Num > 0) And (MapItem(mapnum, I).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, I) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, I).x = GetPlayerX(Index)) Then
                    If (MapItem(mapnum, I).y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(mapnum, I).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(mapnum, I).Num)
                            
                            ' check tasks
                            Call CheckTasks(Index, QUEST_TYPE_GOGATHER, MapItem(mapnum, I).Num)
                            
                            If Item(GetPlayerInvItemNum(Index, n)).Type <> ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(mapnum, I).Value)
                                Msg = MapItem(mapnum, I).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(Index).Inv(n).Bound = 0
                            If Item(GetPlayerInvItemNum(Index, n)).BindType = 1 Or MapItem(mapnum, I).Bound Then
                                Player(Index).Inv(n).Bound = 1
                                If Not Trim$(MapItem(mapnum, I).playerName) = Trim$(GetPlayerName(Index)) Then
                                    PlayerMsg Index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If
                            

                            ' Erase item from the map
                            ClearMapItem I, mapnum
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(I, 0, 0, GetPlayerMap(Index), 0, 0)
                            SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, invNum) > 0) Then
        If (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If Item(GetPlayerInvItemNum(Index, invNum)).BindType > 0 Then
                If Player(Index).Inv(invNum).Bound = 1 Then
                    PlayerMsg Index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            I = FindOpenMapItemSlot(GetPlayerMap(Index))

            If I <> 0 Then
                MapItem(GetPlayerMap(Index), I).Num = GetPlayerInvItemNum(Index, invNum)
                MapItem(GetPlayerMap(Index), I).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), I).y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), I).playerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), I).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), I).canDespawn = True
                MapItem(GetPlayerMap(Index), I).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
                If Player(Index).Inv(invNum).Bound > 0 Then
                    MapItem(GetPlayerMap(Index), I).Bound = True
                Else
                    MapItem(GetPlayerMap(Index), I).Bound = False
                End If

                If Item(GetPlayerInvItemNum(Index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, invNum) Then
                        MapItem(GetPlayerMap(Index), I).Value = GetPlayerInvItemValue(Index, invNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, invNum, 0)
                        Call SetPlayerInvItemValue(Index, invNum, 0)
                        Player(Index).Inv(invNum).Bound = 0
                    Else
                        MapItem(GetPlayerMap(Index), I).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, invNum, GetPlayerInvItemValue(Index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), I).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, invNum, 0)
                    Call SetPlayerInvItemValue(Index, invNum, 0)
                    Player(Index).Inv(invNum).Bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(I, MapItem(GetPlayerMap(Index), I).Num, Amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), I).canDespawn, MapItem(GetPlayerMap(Index), I).Bound)
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long, Optional ByVal level_count As Long)
    Dim I As Long, PontosPorLevel As Byte
    Dim expRollover As Long

    PontosPorLevel = 3

    ' Caso queira adicionar levels diretamente!
    If level_count > 0 Then
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + level_count) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + (level_count * PontosPorLevel))
        GoTo Continue
    End If

    ' Adiciona level pela experi�ncia, m�todo normal de um rpg
    level_count = 0
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)

        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If

        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + PontosPorLevel)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop

Continue:
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
            'Call SendDiscordMsg(Levelup, Index, "has gained " & level_count & " level!")
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
            'Call SendDiscordMsg(Levelup, Index, "has gained " & level_count & " levels!")
        End If
        SendEXP Index
        SendPlayerData Index
    End If
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim I As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    SendVital Index, HP

    ' Drop all worn items
    For I = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(Index, I) > 0 Then
            PlayerMapDropItem Index, GetPlayerEquipment(Index, I), 0
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    With Map(GetPlayerMap(Index)).MapData
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp Index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(I)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.Target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

End Sub

Sub CheckResource(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim I As Long
    Dim Damage As Long
    
    If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).TileData.Tile(x, y).Data1

        ' Get the cache number
        For I = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(I).x = x Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(I).y = y Then
                    Resource_num = I
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                        
                        Damage = Item(GetPlayerEquipment(Index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                End If
                                ' carry on
                                GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                    End If

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Public Sub GiveBankItem(ByVal Index As Long, ByVal invSlot As Long, ByVal Amount As Long)
    Dim BankSlot As Long, ItemNum As Long

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)

    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If Amount < 1 Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            End If
        End If
    End If
    
    SavePlayer Index
    SendBank Index

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), Amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SavePlayer Index
    SendBank Index

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(Index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(Index, GetPlayerExp(Index) - exp)
        SendEXP Index
        Call PlayerMsg(Index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(Index)
End Sub

Public Sub EquipItem(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Dim ItemNum As Long, tempItem As Long
    
    ItemNum = GetPlayerInvItemNum(Index, invNum)
    If ItemNum < 0 And ItemNum > MAX_ITEMS Then Exit Sub
    If Not IsPlayerItemRequerimentsOK(Index, ItemNum) Then Exit Sub
    
    If GetPlayerEquipment(Index, EquipmentSlot) > 0 Then
        tempItem = GetPlayerEquipment(Index, EquipmentSlot)
    End If

    Call SetPlayerEquipment(Index, ItemNum, EquipmentSlot)
    
    Call PlayerMsg(Index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen)
    
    ' tell them if it's soulbound
    If Item(ItemNum).BindType = 2 Then ' BoE
        If Player(Index).Inv(invNum).Bound = 0 Then
            PlayerMsg Index, "This item is now bound to your soul.", BrightRed
        End If
    End If
    
    Call TakeInvItem(Index, ItemNum, 0)

    If tempItem > 0 Then
        If Item(tempItem).BindType > 0 Then
            Call GiveInvItem(Index, tempItem, 0, True) ' give back the stored item
            tempItem = 0
        Else
            Call GiveInvItem(Index, tempItem, 0)
            tempItem = 0
        End If
    End If

    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    
    ' send vitals
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then Call SendPartyVitals(TempPlayer(Index).inParty, Index)
    
    ' send the sound
    Call SendPlayerSound(Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum)
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal invNum As Long)
    Dim n As Long, I As Long
    Dim tempItem As Long
    Dim x As Long, y As Long
    Dim ItemNum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, invNum) > 0) And (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, invNum)).Data2
        ItemNum = GetPlayerInvItemNum(Index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_WEAPON To ITEM_TYPE_FEET
                Call EquipItem(Index, invNum, Item(ItemNum).Type)
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, I) < Item(ItemNum).Stat_Req(I) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp Index, GetPlayerExp(Index) + Item(ItemNum).AddEXP
                    CheckPlayerLevelUp Index
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendEXP Index
                End If
                Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call TakeInvItem(Index, Player(Index).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, I) < Item(ItemNum).Stat_Req(I) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(Index)
                    Case DIR_UP

                        If GetPlayerY(Index) > 0 Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(Index) < Map(GetPlayerMap(Index)).MapData.MaxY Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT

                        If GetPlayerX(Index) > 0 Then
                            x = GetPlayerX(Index) - 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT

                        If GetPlayerX(Index) < Map(GetPlayerMap(Index)).MapData.MaxX Then
                            x = GetPlayerX(Index) + 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If ItemNum = Map(GetPlayerMap(Index)).TileData.Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        SendMapKey Index, x, y, 1
                        'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).TileData.Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(Index, ItemNum, 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_UNIQUE
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, I) < Item(ItemNum).Stat_Req(I) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Go through with it
                Unique_Item Index, ItemNum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, I) < Item(ItemNum).Stat_Req(I) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(ItemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                    
                        ' make sure they don't already know it
                        For I = 1 To MAX_PLAYER_SPELLS
                            If Player(Index).Spell(I).Spell > 0 Then
                                If Player(Index).Spell(I).Spell = n Then
                                    PlayerMsg Index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                                If Spell(Player(Index).Spell(I).Spell).UniqueIndex = Spell(n).UniqueIndex Then
                                    PlayerMsg Index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        Next
                    
                        ' Make sure they are the right level
                        I = Spell(n).LevelReq


                        If I <= GetPlayerLevel(Index) Then
                            I = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If I > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Player(Index).Spell(I).Spell = n
                                    Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                    Call TakeInvItem(Index, ItemNum, 0)
                                    Call PlayerMsg(Index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells Index
                                Else
                                    Call PlayerMsg(Index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "You must be level " & I & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(Index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_FOOD
                ' stat requirements
                For I = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, I) < Item(ItemNum).Stat_Req(I) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' make sure they're not in combat
                If TempPlayer(Index).stopRegen Then
                    PlayerMsg Index, "You cannot eat whilst in combat.", BrightRed
                    Exit Sub
                End If
                
                ' make sure not full hp
                x = Item(ItemNum).HPorSP
                If Player(Index).Vital(x) >= GetPlayerMaxVital(Index, x) Then
                    PlayerMsg Index, "You don't need to eat this at the moment.", BrightRed
                    Exit Sub
                End If
                
                ' set the player's food
                If Item(ItemNum).HPorSP = 2 Then 'mp
                    If Not TempPlayer(Index).foodItem(Vitals.MP) = ItemNum Then
                        TempPlayer(Index).foodItem(Vitals.MP) = ItemNum
                        TempPlayer(Index).foodTick(Vitals.MP) = 0
                        TempPlayer(Index).foodTimer(Vitals.MP) = GetTickCount
                    Else
                        PlayerMsg Index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                Else ' hp
                    If Not TempPlayer(Index).foodItem(Vitals.HP) = ItemNum Then
                        TempPlayer(Index).foodItem(Vitals.HP) = ItemNum
                        TempPlayer(Index).foodTick(Vitals.HP) = 0
                        TempPlayer(Index).foodTimer(Vitals.HP) = GetTickCount
                    Else
                        PlayerMsg Index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' take the item
                Call TakeInvItem(Index, Player(Index).Inv(invNum).Num, 0)
        End Select
    End If
End Sub
