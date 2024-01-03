Attribute VB_Name = "Client_Handle"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SMapNpcDataXY) = GetAddress(AddressOf HandleMapNpcDataXY)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SChatUpdate) = GetAddress(AddressOf HandleChatUpdate)
    HandleDataSub(SConvEditor) = GetAddress(AddressOf HandleConvEditor)
    HandleDataSub(SUpdateConv) = GetAddress(AddressOf HandleUpdateConv)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SPlayerChars) = GetAddress(AddressOf HandlePlayerChars)
    HandleDataSub(SCancelAnimation) = GetAddress(AddressOf HandleCancelAnimation)
    HandleDataSub(SPlayerVariables) = GetAddress(AddressOf HandlePlayerVariables)
    HandleDataSub(SProjectileAttack) = GetAddress(AddressOf HandleProjectile)
    'Quest
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SQuestMessage) = GetAddress(AddressOf HandleQuestMessage)
    HandleDataSub(SQuestCancel) = GetAddress(AddressOf HandleQuestCancel)
    ' Message Window
    HandleDataSub(SMessage) = GetAddress(AddressOf HandleMessageWindow)
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgType = Buffer.ReadLong

    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMsgCOUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long
    
    SetStatus vbNullString
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    dialogue_index = Buffer.ReadLong
    menuReset = Buffer.ReadLong
    kick = Buffer.ReadLong
    
    Buffer.Flush: Set Buffer = Nothing
    
    If menuReset > 0 Then
        HideWindows
        Select Case menuReset
            Case MenuCount.menuLogin
                ShowWindow GetWindowIndex("winLogin")
            Case MenuCount.menuChars
                ShowWindow GetWindowIndex("winCharacters")
            Case MenuCount.menuClass
                ShowWindow GetWindowIndex("winClasses")
            Case MenuCount.menuNewChar
                ShowWindow GetWindowIndex("winNewChar")
            Case MenuCount.menuMain
                ShowWindow GetWindowIndex("winLogin")
            Case MenuCount.menuRegister
                ShowWindow GetWindowIndex("winRegister")
        End Select
    Else
        If kick > 0 Or inMenu = True Then
            ShowWindow GetWindowIndex("winLogin")
            DialogueAlert dialogue_index
            logoutGame
            Exit Sub
        End If
    End If
    
    DialogueAlert dialogue_index
End Sub

Private Sub HandleMessageWindow(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim WindowName As String
    Dim message As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    WindowName = Buffer.ReadString
    message = Buffer.ReadString

    Set Buffer = Nothing

    ShowMessageWindow WindowName, message
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    Call SetStatus("Receiving game data.")
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = Buffer.ReadLong
            Next

            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InGame = True
    inMenu = False
    SetStatus vbNullString
    ' show gui
    ShowWindow GetWindowIndex("winBars"), , False
    ShowWindow GetWindowIndex("winMenu"), , False
    ShowWindow GetWindowIndex("winHotbar"), , False
    ShowWindow GetWindowIndex("winChatSmall"), , False
    ' enter loop
    GameLoop
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        PlayerInv(i).bound = Buffer.ReadByte
    Next
    
    SetGoldLabel

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = Buffer.ReadByte
    Buffer.Flush: Set Buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Pants)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Feet)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Pants)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Feet)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 Then
        BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / 209)) * 209
    Else
        BarWidth_GuiHP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.MP) > 0 Then
        BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / 209)) * 209
    Else
        BarWidth_GuiSP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, Buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    ' set max width
    If GetPlayerLevel(MyIndex) <= MAX_LEVELS Then
        If GetPlayerExp(MyIndex) > 0 Then
            BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / 209) / (TNL / 209)) * 209
        Else
            BarWidth_GuiEXP_Max = 0
        End If
    Else
        BarWidth_GuiEXP_Max = 209
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerClass(i, Buffer.ReadLong)

    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, Buffer.ReadLong
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirLeft = False
        DirDown = False
        DirRight = False
        ' set form
        With Windows(GetWindowIndex("winCharacter"))
            .Controls(GetControlIndex("winCharacter", "lblName")).text = "Name: " & Trim$(GetPlayerName(MyIndex))
            .Controls(GetControlIndex("winCharacter", "lblClass")).text = "Class: " & Trim$(Class(GetPlayerClass(MyIndex)).Name)
            .Controls(GetControlIndex("winCharacter", "lblLevel")).text = "Level: " & GetPlayerLevel(MyIndex)
            .Controls(GetControlIndex("winCharacter", "lblGuild")).text = "Guild: " & "None"
            .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
            .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
            .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Experience: " & Player(MyIndex).EXP & "/" & TNL
            ' stats
            For X = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "lblStat_" & X)).text = GetPlayerStat(MyIndex, X)
            Next
            ' points
            .Controls(GetControlIndex("winCharacter", "lblPoints")).text = GetPlayerPOINTS(MyIndex)
            ' grey out buttons
            If GetPlayerPOINTS(MyIndex) = 0 Then
                For X = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = True
                Next
            Else
                For X = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & X)).visible = False
                Next
            End If
        End With
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)

        Case DIR_UP
            Player(i).yOffset = PIC_Y

        Case DIR_DOWN
            Player(i).yOffset = PIC_Y * -1

        Case DIR_LEFT
            Player(i).xOffset = PIC_X

        Case DIR_RIGHT
            Player(i).xOffset = PIC_X * -1
        
        Case DIR_UP_LEFT
            Player(i).yOffset = PIC_Y
            Player(i).xOffset = PIC_X
            
        Case DIR_UP_RIGHT
            Player(i).yOffset = PIC_Y
            Player(i).xOffset = PIC_X * -1

        Case DIR_DOWN_LEFT
            Player(i).yOffset = PIC_Y * -1
            Player(i).xOffset = PIC_X

        Case DIR_DOWN_RIGHT
            Player(i).yOffset = PIC_Y * -1
            Player(i).xOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim dir As Long, dirBySpell As Byte
    Dim Buffer As clsBuffer
    Dim thePlayer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    dirBySpell = Buffer.ReadByte

    Dim playerX As Long
    Dim playerY As Long

    With Player(thePlayer)

        If dirBySpell > 0 Then
        Select Case dirBySpell

        Case DIR_UP + 1
                playerY = GetPlayerY(thePlayer)
                playerY = playerY - Y
                playerY = playerY * PIC_Y
                .yOffset = playerY

        Case DIR_DOWN + 1
                playerY = GetPlayerY(thePlayer)
                playerY = Y - playerY
                playerY = playerY * PIC_Y
                .yOffset = playerY * -1

        Case DIR_LEFT + 1
                playerX = GetPlayerX(thePlayer)
                playerX = playerX - X
                playerX = playerX * PIC_X
                .xOffset = playerX

        Case DIR_RIGHT + 1
                playerX = GetPlayerX(thePlayer)
                playerX = X - playerX
                playerX = playerX * PIC_X
                .xOffset = playerX * -1
        End Select

            Player(thePlayer).Moving = MOVING_RUNNING
        Else
            Call SetPlayerDir(thePlayer, dir)
            ' Make sure they aren't walking
            Player(thePlayer).Moving = 0
            Player(thePlayer).xOffset = 0
            Player(thePlayer).yOffset = 0
        End If

    End With

    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
End Sub

Private Sub HandleMapNpcDataXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim X As Long
    Dim Y As Long
    Dim dir As Long, dirBySpell As Byte
    Dim npcX As Long, npcY As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    i = Buffer.ReadLong

    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    dir = Buffer.ReadLong
    dirBySpell = Buffer.ReadByte

    With MapNpc(i)
        If dirBySpell > 0 Then
            Select Case dirBySpell
            Case DIR_UP + 1
                npcY = .Y - Y
                npcY = npcY * PIC_Y
                .yOffset = npcY

            Case DIR_DOWN + 1
                npcY = Y - .Y
                npcY = npcY * PIC_Y
                .yOffset = npcY * -1

            Case DIR_LEFT + 1
                npcX = .X - X
                npcX = npcX * PIC_X
                .xOffset = npcX

            Case DIR_RIGHT + 1
                npcX = X - .X
                npcX = npcX * PIC_X
                .xOffset = npcX * -1
            End Select
            
            .Impacted = True
            .ImpactedDir = dirBySpell - 1
        Else
            .dir = dir
        End If

        .X = X
        .Y = Y
    End With
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = getTime
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, NeedMap As Byte, Buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, mapNum As Long
    
    GettingMap = True
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To Player_HighIndex
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).sprite = 0
        Blood(i).timer = 0
    Next

    ' Get map num
    mapNum = Buffer.ReadLong
    MapDataCRC = Buffer.ReadLong
    MapTileCRC = Buffer.ReadLong
    
    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(mapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(mapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
    
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    
    ' load the map if we don't need it
    If NeedMap = 0 Then
        LoadMap mapNum
        applyingMap = False
        CacheNewMapSounds
    End If
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, mapNum As Long, i As Long, X As Long, Y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'zlib
    Buffer.DecompressBuffer
    
    mapNum = Buffer.ReadLong
    
    With Map.MapData
        .Name = Buffer.ReadString
        .Music = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .Left = Buffer.ReadLong
        .Right = Buffer.ReadLong
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        
        .Weather = Buffer.ReadLong
        .WeatherIntensity = Buffer.ReadLong
        
        .Fog = Buffer.ReadLong
        .FogSpeed = Buffer.ReadLong
        .FogOpacity = Buffer.ReadLong
        
        .Red = Buffer.ReadLong
        .Green = Buffer.ReadLong
        .Blue = Buffer.ReadLong
        .alpha = Buffer.ReadLong
        
        .BossNpc = Buffer.ReadLong
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Buffer.ReadLong
        Next
    End With
    
    ReDim Map.TileData.Tile(0 To Map.MapData.MaxX, 0 To Map.MapData.MaxY)

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.TileData.Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Layer(i).tileSet = Buffer.ReadLong
                Map.TileData.Tile(X, Y).Autotile(i) = Buffer.ReadByte
            Next
            Map.TileData.Tile(X, Y).Type = Buffer.ReadByte
            Map.TileData.Tile(X, Y).Data1 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data2 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data3 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data4 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).Data5 = Buffer.ReadLong
            Map.TileData.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    CacheNewMapSounds
    Buffer.Flush: Set Buffer = Nothing
    ' Save the map
    Call SaveMap(mapNum)
    GetMapCRC32 mapNum
    AddText "Downloaded new map.", BrightGreen

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    applyingMap = False

End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .playerName = Buffer.ReadString
            .Num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            tmpLong = Buffer.ReadLong

            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If

        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim musicFile As String

    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i

    Action_HighIndex = 1

    ' player music
    If InGame Then
        musicFile = Trim$(Map.MapData.Music)

        If Not musicFile = "None." Then
            Play_Music musicFile
        Else
            Stop_Music
        End If
    End If

    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1

        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    ' now cache the positions
    initAutotiles
    CurrentWeather = Map.MapData.Weather
    CurrentWeatherIntensity = Map.MapData.WeatherIntensity
    CurrentFog = Map.MapData.Fog
    CurrentFogSpeed = Map.MapData.FogSpeed
    CurrentFogOpacity = Map.MapData.FogOpacity
    CurrentTintR = Map.MapData.Red
    CurrentTintG = Map.MapData.Green
    CurrentTintB = Map.MapData.Blue
    CurrentTintA = Map.MapData.alpha
    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .Num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        tmpLong = Buffer.ReadLong

        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If
        
        .Gravity = -10

    End With

End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadByte
    TempTile(X, Y).DoorOpen = n

    ' re-cache rendering
    If Not GettingMap Then cacheRenderState X, Y, MapLayer.Mask
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = getTime
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, message As String, Color As Long, tmpType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    CreateActionMsg message, Color, tmpType, X, Y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, sprite As Long, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    ' randomise sprite
    sprite = Rand(1, BloodCount)

    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE

        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If

    Next

    ' carry on with the set
    BloodIndex = BloodIndex + 1

    If BloodIndex >= MAX_BYTE Then BloodIndex = 1

    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .sprite = sprite
        .timer = getTime
    End With

End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Slot = Buffer.ReadLong
    SpellCD(Slot) = getTime
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Access As Long, Name As String, message As String, colour As Long, header As String, PK As Long, saycolour As Long
    Dim Channel As Byte, colStr As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
    ' Check access level
    colour = White

    If Access > 0 Then colour = Pink
    If PK > 0 Then colour = BrightRed
    
    ' find channel
    Channel = 0
    Select Case header
        Case "[Map] "
            Channel = ChatChannel.chMap
        Case "[Global] "
            Channel = ChatChannel.chGlobal
    End Select
    
    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)
    ' add to the chat box
    AddText ColourChar & GetColStr(colour) & header & Name & ": " & ColourChar & GetColStr(Grey) & message, Grey, , Channel
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    shopNum = Buffer.ReadLong
    OpenShop shopNum
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    StunDuration = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_BANK
        Bank.Item(i).Num = Buffer.ReadLong
        Bank.Item(i).Value = Buffer.ReadLong
    Next

    InBank = True
    Buffer.Flush: Set Buffer = Nothing
    
    If Not Windows(GetWindowIndex("winBank")).Window.visible Then
        ShowWindow GetWindowIndex("winBank"), , False
    End If
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InTrade = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
    
    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dataType As Byte, i As Long, yourWorth As Long, theirWorth As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    dataType = Buffer.ReadByte

    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        yourWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).text = yourWorth & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        theirWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).text = theirWorth & "g"
    End If

    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeStatus As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    tradeStatus = Buffer.ReadByte
    Buffer.Flush: Set Buffer = Nothing

    Select Case tradeStatus
        Case 0 ' clear
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
        Case 1 ' they've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Other player has accepted."
        Case 2 ' you've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Waiting for other player to accept."
        Case 3 ' no room
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Not enough inventory space."
    End Select
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player_HighIndex = Buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    PlayMapSound X, Y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Index_Offer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Index_Offer = FindOpenOfferSlot
    
    If Index_Offer <> 0 Then
        inOfferInvite(Index_Offer) = Buffer.ReadString
        inOfferType(Index_Offer) = Offers.Offer_Type_Trade
    End If
    Buffer.Flush: Set Buffer = Nothing
    
    Call UpdateWindowOffer(Index_Offer)
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Top As Long
    Dim Index_Offer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Index_Offer = FindOpenOfferSlot
    
    If Index_Offer <> 0 Then
        inOfferInvite(Index_Offer) = Buffer.ReadString
        inOfferType(Index_Offer) = Offers.Offer_Type_Party
    End If
    Buffer.Flush: Set Buffer = Nothing
    
    Call UpdateWindowOffer(Index_Offer)
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, inParty As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    inParty = Buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = Buffer.ReadLong

    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
    Next

    Party.MemberCount = Buffer.ReadLong
    
    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long
    Dim Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' which player?
    playerNum = Buffer.ReadLong

    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
    Next

    ' update the party interface
    UpdatePartyBars
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, NpcNum As Long, mT As String, o(1 To 4) As String, i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    mT = Buffer.ReadString
    For i = 1 To 4
        o(i) = Buffer.ReadString
    Next

    Buffer.Flush: Set Buffer = Nothing

    ' if npcNum is 0, exit the chat system
    If NpcNum = 0 Then
        inChat = False
        HideWindow GetWindowIndex("winNpcChat")
        Exit Sub
    End If

    ' set chat going
    OpenNpcChat NpcNum, mT, o
End Sub

Private Sub HandleStartTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    'inTutorial = True
    ' set the first message
    'SetTutorialState 1
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, TargetType As Long, target As Long, message As String, colour As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    message = Buffer.ReadString
    colour = Buffer.ReadLong
    AddChatBubble target, TargetType, message, colour
    Buffer.Flush: Set Buffer = Nothing
End Sub

Private Sub HandlePlayerChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, winNum As Long, conNum As Long, isSlotEmpty(1 To MAX_CHARS) As Boolean, X As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    For i = 1 To MAX_CHARS
        CharName(i) = Trim$(Buffer.ReadString)
        CharSprite(i) = Buffer.ReadLong
        CharAccess(i) = Buffer.ReadLong
        CharClass(i) = Buffer.ReadLong
        ' set as empty or not
        If Not Len(Trim$(CharName(i))) > 0 Then isSlotEmpty(i) = True
    Next

    Buffer.Flush: Set Buffer = Nothing
    
    HideWindows
    ShowWindow GetWindowIndex("winCharacters")
    
    ' set GUI window up
    winNum = GetWindowIndex("winCharacters")
    For i = 1 To MAX_CHARS
        conNum = GetControlIndex("winCharacters", "lblCharName_" & i)
        With Windows(winNum).Controls(conNum)
            If Not isSlotEmpty(i) Then
                .text = CharName(i)
            Else
                .text = "Blank Slot"
            End If
        End With
        ' hide/show buttons
        If isSlotEmpty(i) Then
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
        Else
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
        End If
    Next
End Sub

Private Sub HandlePlayerVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_BYTE
        Player(MyIndex).Variable(i) = Buffer.ReadLong
    Next
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub HandleProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileSlot As Long, i As Long, Angle As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    ' read bytes from data()
    Buffer.WriteBytes Data()

    ' recieve projectile data
    ProjectileSlot = Buffer.ReadLong
    LastProjectile = Buffer.ReadLong
    
    ReDim Preserve MapProjectile(1 To LastProjectile)
        
    ' populate the values
    With MapProjectile(ProjectileSlot)
        .Owner = Buffer.ReadLong
        .OwnerType = Buffer.ReadLong
        .Direction = Buffer.ReadLong
        .Graphic = Buffer.ReadLong
        .IsAoE = Buffer.ReadByte
        .Rotate = Buffer.ReadLong
        .RotateSpeed = Buffer.ReadLong
        .Speed = Buffer.ReadLong
        .Duration = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .tx = Buffer.ReadLong
        .ty = Buffer.ReadLong
        .IsDirectional = Buffer.ReadByte
        
        If .Speed >= 5000 Then
            .Duration = Tick + .Duration
        End If
        
        For i = 1 To 4
            .ProjectileOffset(i).X = Buffer.ReadLong
            .ProjectileOffset(i).Y = Buffer.ReadLong
        Next
        
    End With

End Sub

Sub HandleClearProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ProjectileSlot As Long, i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ProjectileSlot = Buffer.ReadLong
    ClearProjectile ProjectileSlot
    
    Buffer.Flush
    Set Buffer = Nothing
End Sub
