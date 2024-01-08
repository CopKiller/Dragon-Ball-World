Attribute VB_Name = "Player_Combat"
' ################################
' ##      Basic Calculations    ##
' ################################

Public Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
                Case 2 ' Wizard
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 5 + 65
                Case 3 ' Whisperer
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Wizard
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
                Case 3 ' Whisperer
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Public Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = 10 '(GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = 10 '(GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Public Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = Item(weaponNum).Data2 + (((Item(weaponNum).Data2 / 100) * 5) * GetPlayerStat(index, Strength))
    Else
        GetPlayerDamage = 1 + (((0.01) * 5) * GetPlayerStat(index, Strength))
    End If

End Function

Public Function GetPlayerDefence(ByVal index As Long) As Long
    Dim Defence As Long, i As Long, ItemNum As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' base defence
    For i = 1 To Equipment.Equipment_Count - 1
        If i <> Equipment.Weapon Then
            ItemNum = GetPlayerEquipment(index, i)
            If ItemNum > 0 Then
                If Item(ItemNum).Data2 > 0 Then
                    Defence = Defence + Item(ItemNum).Data2
                End If
            End If
        End If
    Next
    
    ' divide by 3
    Defence = Defence / 3
    
    ' floor it at 1
    If Defence < 1 Then Defence = 1
    
    ' add in a player's agility
    GetPlayerDefence = Defence + (((Defence / 100) * 2.5) * (GetPlayerStat(index, Agility) / 2))
End Function

Public Function GetPlayerSpellDamage(ByVal index As Long, ByVal spellNum As Long) As Long
    Dim Damage As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' return damage
    Damage = Spell(spellNum).Vital
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetPlayerSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim rndNum As Long

    CanPlayerCrit = False

    Rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= Rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim rndNum As Long

    CanPlayerDodge = False

    Rate = GetPlayerStat(index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim rndNum As Long

    CanPlayerParry = False

    Rate = GetPlayerStat(index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= Rate Then
        CanPlayerParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        'If CanNpcDodge(npcNum) Then
        '    SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).Npc(mapNpcNum).x * 32), (MapNpc(mapnum).Npc(mapNpcNum).y * 32)
        '    Exit Sub
        'End If
        'If CanNpcParry(npcNum) Then
        '    SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).Npc(mapNpcNum).x * 32), (MapNpc(mapnum).Npc(mapNpcNum).y * 32)
        '    Exit Sub
        'End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - RAND((GetNpcDefence(npcNum) / 100) * 10, (GetNpcDefence(npcNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32), alert
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            Exit Function
        End If
    End If

    ' exit out early
    If isSpell Then
        If npcNum > 0 Then
            If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                CanPlayerAttackNpc = True
                Exit Function
            End If
        End If
    End If

    ' attack speed from weapon
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    If npcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
        Case DIR_UP
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y + 1
        Case DIR_DOWN
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y - 1
        Case DIR_LEFT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x + 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y
        Case DIR_RIGHT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x - 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y
        Case DIR_UP_RIGHT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x - 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y + 1
        Case DIR_UP_LEFT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x + 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y + 1
        Case DIR_DOWN_RIGHT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x - 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y - 1
        Case DIR_DOWN_LEFT
            NpcX = MapNpc(mapnum).Npc(mapNpcNum).x + 1
            NpcY = MapNpc(mapnum).Npc(mapNpcNum).y - 1
        End Select

        If NpcX = GetPlayerX(Attacker) Then
            If NpcY = GetPlayerY(Attacker) Then
                If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                ElseIf Npc(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then
                    ' init quest tasks
                    Call CheckTasks(Attacker, QUEST_TYPE_GOTALK, npcNum)
                    Call CheckTasks(Attacker, QUEST_TYPE_GOGIVE, npcNum)
                    Call CheckTasks(Attacker, QUEST_TYPE_GOGET, npcNum)
                    ' init conversation if it's friendly
                    If Npc(npcNum).Conv > 0 Then
                        InitChat Attacker, mapnum, mapNpcNum
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    Name = Trim$(Npc(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).Npc(mapNpcNum).x * 32), (MapNpc(mapnum).Npc(mapNpcNum).y * 32), fonts.Damage
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Attacker, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, Npc(npcNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, exp, Npc(npcNum).Level
        End If
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(npcNum).DropItem(n) = 0 Then Exit For
            If Rnd <= Npc(npcNum).DropChance(n) Then
                Call SpawnItem(Npc(npcNum).DropItem(n), Npc(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, GetPlayerName(Attacker))
            End If
        Next
        
        ' destroy map npcs
        If Map(mapnum).MapData.Moral = MAP_MORAL_BOSS Then
            If mapNpcNum = Map(mapnum).MapData.BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(mapnum).MapData.Npc(i) > 0 Then
                        ' only kill dangerous npcs
                        If Npc(Map(mapnum).MapData.Npc(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(Map(mapnum).MapData.Npc(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(mapnum).Npc(i).Num = 0
                            MapNpc(mapnum).Npc(i).SpawnWait = GetTickCount
                            MapNpc(mapnum).Npc(i).Vital(Vitals.HP) = 0
                            ' send kill command
                            SendNpcDeath mapnum, i
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).Npc(mapNpcNum).Num = 0
        MapNpc(mapnum).Npc(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).Npc(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).Npc(mapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' check task
        Call CheckTasks(Attacker, QUEST_TYPE_GOSLAY, npcNum)
        
        ' send death to the map
        SendNpcDeath mapnum, mapNpcNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = mapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).Npc(mapNpcNum).x * 32), (MapNpc(mapnum).Npc(mapNpcNum).y * 32), fonts.Damage
        SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Attacker, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapnum).Npc(mapNpcNum).TargetType = 1 ' player
        MapNpc(mapnum).Npc(mapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(mapnum).Npc(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).Npc(i).Num = MapNpc(mapnum).Npc(mapNpcNum).Num Then
                    MapNpc(mapnum).Npc(i).Target = Attacker
                    MapNpc(mapnum).Npc(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).Npc(mapNpcNum).stopRegen = True
        MapNpc(mapnum).Npc(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellNum, Attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
        
        ' set the player's target if they don't have one
        If TempPlayer(Attacker).Target = 0 Then
            TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).Target = mapNpcNum
            SendTarget Attacker
        End If
    End If

    If spellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub
' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim npcNum As Long, mapnum As Long, Damage As Long, Defence As Long, blockAmount As Long ' percent %

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
    
        mapnum = GetPlayerMap(Attacker)
        
        'Block Amount
        blockAmount = CanPlayerBlockHit(Victim, Damage, TARGET_TYPE_PLAYER, Attacker)

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        If blockAmount > 0 Then
            Damage = Damage - blockAmount
            SendActionMsg mapnum, "BLOQUEOU -" & blockAmount, Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32), alert
        End If
        
        ' take away armour
        Defence = GetPlayerDefence(Victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32), alert
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal isSpell As Boolean = False) As Boolean
Dim partynum As Long, i As Long

    If Not isSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If Victim = Attacker Then
        PlayerMsg Attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not isSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UP_LEFT
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UP_RIGHT
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN_LEFT
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN_RIGHT
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).MapData.Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 5 Then
        Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 5 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' make sure not in your party
    partynum = TempPlayer(Attacker).inParty
    If partynum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) > 0 Then
                If Victim = Party(partynum).Member(i) Then
                    PlayerMsg Attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32), fonts.Damage
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, GetPlayerLevel(Victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp, GetPlayerLevel(Victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        ' check task
        Call CheckTasks(Attacker, QUEST_TYPE_GOKILL, Victim)

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If spellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellNum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32), fonts.Damage
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer Victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Player Victim, spellNum, Attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(Attacker).Target = 0 Then
            TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER
            TempPlayer(Attacker).Target = Victim
            SendTarget Attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long, mapnum As Long, ClassReq As Long
    Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, TargetType As Byte, Target As Long
    Dim PlayerProjectileSlot As Long, ProjectileSlot As Long
    
    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellNum = Player(index).Spell(spellSlot).Spell
    mapnum = GetPlayerMap(index)
    
    If spellNum <= 0 Or spellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellNum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(index).spellBuffer.Spell = spellSlot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    mpCost = Spell(spellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < mpCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Type = SPELL_TYPE_PROJECTILE Then
        If Spell(spellNum).Range = 0 Then Exit Sub
        
        If Spell(spellNum).Projectile.Ammo > 0 Then
            If HasItem(index, Spell(spellNum).Projectile.Ammo) Then
                TakeInvItem index, Spell(spellNum).Projectile.Ammo, 1
            Else
                PlayerMsg index, "Suas ferramentas acabaram.", BrightRed
                Exit Sub
            End If
        End If
        TempPlayer(index).SpellCastType = 0
    Else
        If Spell(spellNum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellNum).IsAoE Then
                TempPlayer(Idex).SpellCastType = 2 ' targetted
            Else
                TempPlayer(Idex).SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellNum).IsAoE Then
                TempPlayer(Idex).SpellCastType = 0 ' self-cast
            Else
                TempPlayer(Idex).SpellCastType = 1 ' self-cast AoE
            End If
        End If
    End If
    
    TargetType = TempPlayer(index).TargetType
    Target = TempPlayer(index).Target
    Range = Spell(spellNum).Range
    HasBuffered = False
    
    Select Case TempPlayer(index).SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If Spell(spellNum).Type = SPELL_TYPE_HEALHP Or Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                    Target = index
                    TargetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).Npc(Target).x, MapNpc(mapnum).Npc(Target).y) Then
                        PlayerMsg index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        ' go through spell types
                        If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                            HasBuffered = True
                        Else
                            If CanPlayerAttackNpc(index, Target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index, 1
        TempPlayer(index).spellBuffer.Spell = spellSlot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.Target = Target
        TempPlayer(index).spellBuffer.tType = TargetType
        
        If Spell(spellNum).CastFrame > 0 Then
            TempPlayer(index).PlayerFrame = Spell(spellNum).CastFrame
            SendPlayerFrameToMapBut index
        End If
        
        If Spell(spellNum).Projectile.projectileType = ProjectileTypeEnum.GekiDama Then
            SendPlayerConjureProjectileCustomToMapBut index, ProjectileTypeEnum.GekiDama, Spell(spellNum).Projectile.Graphic
        End If
        
        Exit Sub
    Else
        SendClearSpellBufferTo index
    End If
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellSlot As Long, ByVal Target As Long, ByVal TargetType As Byte)
    Dim spellNum As Long, mpCost As Long, LevelReq As Long
    Dim mapnum As Long, Vital As Long, DidCast As Boolean, ClassReq As Long
    Dim AccessReq As Long, i As Long, AoE As Long, Range As Long
    Dim vitalType As Byte, increment As Boolean, x As Long, y As Long
    Dim Buffer As clsBuffer, SpellCastType As Long
    Dim ProjectileSlot As Long, PlayerProjectileSlot As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    spellNum = Player(index).Spell(spellSlot).Spell
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellNum) Then Exit Sub

    mpCost = Spell(spellNum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < mpCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
'    ' find out what kind of spell it is! self cast, target or AOE
'    If Spell(spellNum).Range > 0 Then
'        ' ranged attack, single target or aoe?
'        If Not Spell(spellNum).IsAoE Then
'            SpellCastType = 2 ' targetted
'        Else
'            SpellCastType = 3 ' targetted aoe
'        End If
'    Else
'        If Not Spell(spellNum).IsAoE Then
'            SpellCastType = 0 ' self-cast
'        Else
'            SpellCastType = 1 ' self-cast AoE
'        End If
'    End If
    
    ' get damage
    Vital = GetPlayerSpellDamage(index, spellNum)
    
    ' store data
    AoE = Spell(spellNum).RadiusX
    Range = Spell(spellNum).Range
    
    Select Case TempPlayer(index).SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(spellNum).Map, Spell(spellNum).x, Spell(spellNum).y
                    SendAnimation GetPlayerMap(index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
                Case SPELL_TYPE_PROJECTILE
                    SpellPlayer_Projectile index, spellNum, GetPlayerMap(index)
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If TargetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).Npc(Target).x
                    y = MapNpc(mapnum).Npc(Target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBufferTo index
                End If
            End If
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, spellNum
                                            DidCast = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).Npc(i).Num > 0 Then
                            If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).Npc(i).x, MapNpc(mapnum).Npc(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, Vital, spellNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        vitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        vitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        vitalType = Vitals.MP
                        increment = False
                    End If
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect vitalType, increment, i, Vital, spellNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(mapnum).Npc(i).Num > 0 Then
                                If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, x, y, MapNpc(mapnum).Npc(i).x, MapNpc(mapnum).Npc(i).y) Then
                                        SpellNpc_Effect vitalType, increment, i, Vital, spellNum, mapnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
                    End If
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            
            If TargetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).Npc(Target).x
                y = MapNpc(mapnum).Npc(Target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBufferTo index
                Exit Sub
            End If
            
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        vitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        vitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        vitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, True) Then
                                SpellPlayer_Effect vitalType, increment, Target, Vital, spellNum
                                DidCast = True
                            End If
                        Else
                            SpellPlayer_Effect vitalType, increment, Target, Vital, spellNum
                            DidCast = True
                        End If
                    Else
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, True) Then
                                SpellNpc_Effect vitalType, increment, Target, Vital, spellNum, mapnum
                                DidCast = True
                            End If
                        Else
                            SpellNpc_Effect vitalType, increment, Target, Vital, spellNum, mapnum
                            DidCast = True
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - mpCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
        TempPlayer(index).SpellCD(spellSlot) = GetTickCount + (Spell(spellNum).CDTime * 1000)
        Call SendCooldown(index, spellSlot)
        
        ' if has a next rank then increment usage
        SetPlayerSpellUsage index, spellSlot
    End If
End Sub

Public Sub SetPlayerSpellUsage(ByVal index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long, i As Long
    spellNum = Player(index).Spell(spellSlot).Spell
    ' if has a next rank then increment usage
    If Spell(spellNum).NextRank > 0 Then
        If Player(index).Spell(spellSlot).Uses < Spell(spellNum).NextUses - 1 Then
            Player(index).Spell(spellSlot).Uses = Player(index).Spell(spellSlot).Uses + 1
        Else
            If GetPlayerLevel(index) >= Spell(Spell(spellNum).NextRank).LevelReq Then
                Player(index).Spell(spellSlot).Spell = Spell(spellNum).NextRank
                Player(index).Spell(spellSlot).Uses = 0
                PlayerMsg index, "Your spell has ranked up!", Blue
                ' update hotbar
                For i = 1 To MAX_HOTBAR
                    If Player(index).Hotbar(i).Slot > 0 Then
                        If Player(index).Hotbar(i).sType = 2 Then ' spell
                            If Spell(Player(index).Hotbar(i).Slot).UniqueIndex = Spell(Spell(spellNum).NextRank).UniqueIndex Then
                                Player(index).Hotbar(i).Slot = Spell(spellNum).NextRank
                                SendHotbar index
                            End If
                        End If
                    End If
                Next
            Else
                Player(index).Spell(spellSlot).Uses = Spell(spellNum).NextUses
            End If
        End If
        SendPlayerSpells index
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellNum As Long)
    Dim sSymbol As String * 1
    Dim colour As Long
    Dim fonte As fonts

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then colour = BrightGreen: fonte = health
            If Vital = Vitals.MP Then colour = BrightBlue: fonte = energy
        Else
            sSymbol = "-"
            colour = Blue
            fonte = Damage
        End If
    
        SendAnimation GetPlayerMap(index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, fonte
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellNum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Player index, spellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
        
        ' send update
        SendVital index, Vital
    End If
End Sub

Public Sub SpellPlayer_Projectile(ByVal index As Long, spellNum As Long, mapnum As Long)
    Dim TargetType As Byte, Target As Long, Range As Byte, i As Long
    Dim xT As Long, yT As Long, Rotate As Long
    
    Dim ProjectileIndex As Integer
    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > MapProjectile_HighIndex Then
            MapProjectile_HighIndex = ProjectileIndex
            ReDim Preserve MapProjectile(1 To MapProjectile_HighIndex)
            Exit Do
        End If
        
    Loop While MapProjectile(ProjectileIndex).Graphic > 0
    
    TargetType = TempPlayer(index).TargetType
    Target = TempPlayer(index).Target
    Range = Spell(spellNum).Range
    
    With MapProjectile(ProjectileIndex)
        
        ' SE É UMA PROJECTILE
        If Spell(spellNum).Projectile.Speed < 5000 Then
            ' DEFINE OS VALORES INICIAIS
            .x = GetPlayerX(index)
            .y = GetPlayerY(index)
            ' SE TEMOS UMA PROJECTILE DE DANO EM AREA
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(index)
                    Case DIR_UP
                        .tX = GetPlayerX(index)
                        If GetPlayerY(index) - Spell(spellNum).Range >= 0 Then
                            .tY = GetPlayerY(index) - Spell(spellNum).Range
                        Else
                            .tY = 0
                        End If
                    Case DIR_DOWN
                        .tX = GetPlayerX(index)
                        If GetPlayerY(index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxY Then
                            .tY = GetPlayerY(index) + Spell(spellNum).Range
                        Else
                            .tY = Map(mapnum).MapData.MaxY
                        End If
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        If GetPlayerX(index) - Spell(spellNum).Range >= 0 Then
                            .tX = GetPlayerX(index) - Spell(spellNum).Range
                        Else
                            .tX = 0
                        End If
                        .tY = GetPlayerY(index)
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        If GetPlayerX(index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxX Then
                            .tX = GetPlayerX(index) + Spell(spellNum).Range
                        Else
                            .tX = Map(mapnum).MapData.MaxX
                        End If
                        .tY = GetPlayerY(index)
                End Select
            ' DEFINIR A POSIÇÃO DO ALVO
            Else
                ' SE TEMOS UM ALVO
                If Target > 0 Then
                    ' SE É UM ALVO DO TIPO PLAYER
                    If TargetType = TARGET_TYPE_PLAYER Then
                        ' SE ESTÁ FORA DE ALCANCE
                        If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                            
                        ' SE ESTÁ DENTRO DO ALCANCE
                        Else
                            .tX = GetPlayerX(TempPlayer(index).Target)
                            .tY = GetPlayerY(TempPlayer(index).Target)
                        End If
                    ' SE É UM ALVO DO TIPO NPC
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        ' SE ESTÁ FORA DA ALCANCE
                        If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).Npc(Target).x, MapNpc(mapnum).Npc(Target).y) Then
                            yT = MapNpc(mapnum).Npc(Target).y
                            xT = MapNpc(mapnum).Npc(Target).x
                            Do
                                ' Up left
                                If GetPlayerY(index) < yT And GetPlayerX(index) < xT Then
                                    yT = yT - 1
                                    xT = xT - 1
                                End If
                                    
                                ' Up right
                                If GetPlayerY(index) < yT And GetPlayerX(index) > xT Then
                                    yT = yT - 1
                                    xT = xT + 1
                                End If
                                    
                                ' Down left
                                If GetPlayerY(index) > yT And GetPlayerX(index) < xT Then
                                    yT = yT + 1
                                    xT = xT - 1
                                End If
                                    
                                ' Down right
                                If GetPlayerY(index) > yT And GetPlayerX(index) > xT Then
                                    yT = yT + 1
                                    xT = xT + 1
                                End If
                                    
                                ' Up
                                If GetPlayerY(index) < yT Then
                                    yT = yT - 1
                                End If
                                    
                                ' Down
                                If GetPlayerY(index) > yT Then
                                    yT = yT + 1
                                End If
                                    
                                ' left
                                If GetPlayerX(index) < xT Then
                                    xT = xT - 1
                                End If
                                    
                                ' right
                                If GetPlayerX(index) > xT Then
                                    xT = xT + 1
                                End If
                                
                            Loop Until isInRange(Range, GetPlayerX(index), GetPlayerY(index), xT, yT)
                            .tX = xT
                            .tY = yT
                        ' SE ESTÁ DENTRO DO ALCANCE
                        Else
                            .tX = MapNpc(mapnum).Npc(TempPlayer(index).Target).x
                            .tY = MapNpc(mapnum).Npc(TempPlayer(index).Target).y
                        End If
                    End If
                ' SE NÃO TEMOS UM ALVO DEFINIR O ALVO NO ALCANCE MÁXIMO
                Else
                    Select Case GetPlayerDir(index)
                        Case DIR_UP
                            .tX = GetPlayerX(index)
                            If GetPlayerY(index) - Spell(spellNum).Range >= 0 Then
                                .tY = GetPlayerY(index) - Spell(spellNum).Range
                            Else
                                .tY = 0
                            End If
                        Case DIR_DOWN
                            .tX = GetPlayerX(index)
                            If GetPlayerY(index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxY Then
                                .tY = GetPlayerY(index) + Spell(spellNum).Range
                            Else
                                .tY = Map(mapnum).MapData.MaxY
                            End If
                        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                            If GetPlayerX(index) - Spell(spellNum).Range >= 0 Then
                                .tX = GetPlayerX(index) - Spell(spellNum).Range
                            Else
                                .tX = 0
                            End If
                            .tY = GetPlayerY(index)
                        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                            If GetPlayerX(index) + Spell(spellNum).Range <= Map(mapnum).MapData.MaxX Then
                                .tX = GetPlayerX(index) + Spell(spellNum).Range
                            Else
                                .tX = Map(mapnum).MapData.MaxX
                            End If
                            .tY = GetPlayerY(index)
                    End Select
                End If
            End If 'If Spell(spellNum).IsAoE Then
            
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(index)
                    Case DIR_UP
                        .xTargetAoE = (GetPlayerX(index) - Int(Spell(spellNum).DirectionAoE(DIR_UP + 1).x / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - 1) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).y
                        Else
                            .Duration = 1
                        End If
                    Case DIR_DOWN
                        .xTargetAoE = (GetPlayerX(index) - Int(Spell(spellNum).DirectionAoE(DIR_DOWN + 1).x / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) + 1) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).y
                        Else
                            .Duration = 1
                        End If
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        .xTargetAoE = (GetPlayerX(index) - 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - Int(Spell(spellNum).DirectionAoE(DIR_LEFT + 1).y) / 2) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).x
                        Else
                            .Duration = 1
                        End If
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        .xTargetAoE = (GetPlayerX(index) + 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - Int(Spell(spellNum).DirectionAoE(DIR_RIGHT + 1).y / 2)) * PIC_Y
                        If Spell(spellNum).Projectile.RecuringDamage Then
                            .Duration = Spell(spellNum).DirectionAoE(DIR_UP + 1).x
                        Else
                            .Duration = 1
                        End If
                End Select
            End If
            
            ' DEFINE O ANGULO INICIAL DE ROTAÇÃO
            .Rotate = Engine_GetAngle(.x, .y, .tX, .tY)
            ' DEFINE A VELOCIDADE DE ROTAÇÃO
            .RotateSpeed = Spell(spellNum).Projectile.Rotation
            
            ' DEFINE O LADO QUE O PLAYER DEVE VIRAR ANTES DE SOLTAR A SKILL
            If .Rotate >= 315 And .Rotate <= 360 Then
                Call SetPlayerDir(index, DIR_UP)
            ElseIf .Rotate >= 0 And .Rotate <= 45 Then
                Call SetPlayerDir(index, DIR_UP)
            ElseIf .Rotate >= 225 And .Rotate <= 315 Then
                Call SetPlayerDir(index, DIR_LEFT)
            ElseIf .Rotate >= 135 And .Rotate <= 225 Then
                Call SetPlayerDir(index, DIR_DOWN)
            ElseIf .Rotate >= 45 And .Rotate <= 135 Then
                Call SetPlayerDir(index, DIR_RIGHT)
            End If
            
            Dim Buffer As clsBuffer
            
            Set Buffer = New clsBuffer
            Buffer.WriteLong SPlayerDir
            Buffer.WriteLong index
            Buffer.WriteLong GetPlayerDir(index)
            Call SendDataToMap(mapnum, Buffer.ToArray())
            Set Buffer = Nothing
        ' SE É UMA TRAP
        Else
            If Spell(spellNum).IsAoE Then
                Select Case GetPlayerDir(index)
                    Case DIR_UP
                        .xTargetAoE = (GetPlayerX(index) - Int(Spell(spellNum).DirectionAoE(DIR_UP + 1).x / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - 1) * PIC_Y
                    Case DIR_DOWN
                        .xTargetAoE = (GetPlayerX(index) - Int(Spell(spellNum).DirectionAoE(DIR_DOWN + 1).x / 2)) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) + 1) * PIC_Y
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        .xTargetAoE = (GetPlayerX(index) - 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - Int(Spell(spellNum).DirectionAoE(DIR_LEFT + 1).y) / 2) * PIC_Y
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        .xTargetAoE = (GetPlayerX(index) + 1) * PIC_X
                        .yTargetAoE = (GetPlayerY(index) - Int(Spell(spellNum).DirectionAoE(DIR_RIGHT + 1).y / 2)) * PIC_Y
                End Select
            End If
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    .x = GetPlayerX(index)
                    If GetPlayerY(index) - 1 < 0 Then
                        Exit Sub
                    Else
                        .y = GetPlayerY(index) - 1
                    End If
                Case DIR_DOWN
                    .x = GetPlayerX(index)
                    If GetPlayerY(index) + 1 > Map(mapnum).MapData.MaxY Then
                        Exit Sub
                    Else
                        .y = GetPlayerY(index) + 1
                    End If
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    If GetPlayerX(index) - 1 < 0 Then
                        Exit Sub
                    Else
                        .x = GetPlayerX(index) - 1
                    End If
                    .y = GetPlayerY(index)
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    If GetPlayerX(index) + 1 > Map(mapnum).MapData.MaxX Then
                        Exit Sub
                    Else
                        .x = GetPlayerX(index) + 1
                    End If
                    .y = GetPlayerY(index)
            End Select
            
            .Duration = Spell(spellNum).Projectile.Despawn
        End If
        
        ' DEFINE OS DADOS DO DONO DA PROJECTILE
        .Owner = index
        .OwnerType = TARGET_TYPE_PLAYER
        ' DEFINE A DIREÇÃO DA PROJECTILE
        .direction = GetPlayerDir(index)
        ' DEFINE O GRÁFICO DO PROJECTILE
        .Graphic = Spell(spellNum).Projectile.Graphic
        ' DEFINE A VELOCIDADE DA PROJECTILE
        .Speed = Spell(spellNum).Projectile.Speed
        ' ALTERA AS POSIÇÕES DA X,Y E TARGET X,Y
        .x = .x * PIC_X
        .y = .y * PIC_Y
        .tX = .tX * PIC_X
        .tY = .tY * PIC_Y
        .spellNum = spellNum
        ' DEFINE OS OFFSET DE X E Y PARA EXIBIR NA POSIÇÃO CERTA NO MAPA
        For i = 1 To 4
            .ProjectileOffset(i).x = Spell(spellNum).Projectile.ProjectileOffset(i).x
            .ProjectileOffset(i).y = Spell(spellNum).Projectile.ProjectileOffset(i).y
        Next
        
        ' DEFINE O MAPA DA PROJECTILE
        .mapnum = mapnum
        
        Call SendProjectile(mapnum, ProjectileIndex, Spell(spellNum).IsDirectional)
        If .Speed >= 5000 Then
            .Duration = .Duration + tick
        End If
    End With
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, GetPlayerSpellDamage(.Caster, .Spell)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(index).Map, "+" & GetPlayerSpellDamage(.Caster, .Spell), BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32, health
                Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell)
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(spellNum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If
End Sub

Sub MakeImpact(ByVal index As Long, ByVal ImpactValue As Byte, ByVal TargetType As Byte, Optional ByVal mapnum As Long, Optional ByVal Attacker As Long, Optional ByVal NpcToPlayer As Boolean)
    Dim i As Long, x As Long, y As Long, Dir As Byte
    Dim XDif, YDif As Long

    If TargetType = TARGET_TYPE_PLAYER Then
        x = Player(index).x
        y = Player(index).y

        If NpcToPlayer = True Then
            XDif = x - MapNpc(mapnum).Npc(Attacker).x
            YDif = y - MapNpc(mapnum).Npc(Attacker).y
        Else
            XDif = x - Player(Attacker).x
            YDif = y - Player(Attacker).y
        End If

        If XDif = 0 Then
            If YDif < 0 Then Dir = DIR_UP
            If YDif > 0 Then Dir = DIR_DOWN
        Else
            If XDif < 0 Then Dir = DIR_LEFT
            If XDif > 0 Then Dir = DIR_RIGHT
        End If

        For i = 1 To ImpactValue
            Select Case Dir
            Case DIR_UP: y = y - 1
            Case DIR_DOWN: y = y + 1
            Case DIR_LEFT: x = x - 1
            Case DIR_RIGHT: x = x + 1
            End Select

            If x > 0 And x < Map(Player(index).Map).MapData.MaxX Then
                If y > 0 And y < Map(Player(index).Map).MapData.MaxY Then
                    If Map(Player(index).Map).TileData.Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        Player(index).x = x
                        Player(index).y = y
                    Else
                        Exit For
                    End If
                End If
            End If
        Next i

        'TempPlayer(index).ImpactedBy = Attacker
        TempPlayer(index).ImpactedTick = GetTickCount + 100
        SendPlayerXYToMap index, Dir + 1
    End If

    If TargetType = TARGET_TYPE_NPC Then
        If index < 1 Then Exit Sub
        x = MapNpc(mapnum).Npc(index).x
        y = MapNpc(mapnum).Npc(index).y

        XDif = x - Player(Attacker).x
        YDif = y - Player(Attacker).y

        If XDif = 0 Then
            If YDif < 0 Then Dir = DIR_UP
            If YDif > 0 Then Dir = DIR_DOWN
        Else
            If XDif < 0 Then Dir = DIR_LEFT
            If XDif > 0 Then Dir = DIR_RIGHT
        End If

        For i = 1 To ImpactValue
            Select Case Dir
            Case DIR_UP: y = y - 1
            Case DIR_DOWN: y = y + 1
            Case DIR_LEFT: x = x - 1
            Case DIR_RIGHT: x = x + 1
            End Select

            If mapnum > 0 Then
                If x > 0 And x < Map(mapnum).MapData.MaxX Then
                    If y > 0 And y < Map(mapnum).MapData.MaxY Then
                        If Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                            MapNpc(mapnum).Npc(index).x = x
                            MapNpc(mapnum).Npc(index).y = y

                        Else
                            Exit For
                        End If
                    End If
                End If
            End If
        Next i

        'MapNpc(mapnum).Npc(index).ImpactedBy = Attacker
        MapNpc(mapnum).Npc(index).ImpactedTick = GetTickCount + 100
        SendMapNpcXY index, mapnum, Dir + 1
    End If

End Sub

