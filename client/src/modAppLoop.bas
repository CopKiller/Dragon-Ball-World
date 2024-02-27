Attribute VB_Name = "modAppLoop"
Option Explicit

'Loop
'Public Const TICKS_PER_SECOND As Long = 60
'Public Const SKIP_TICKS = 1000 / TICKS_PER_SECOND
'Public Const MAX_FRAME_SKIP = 5

' Timers
Public Tick As Single
Public ElapsedTime As Single
Public FrameTime As Long
Public TickFPS As Long
Public FPS As Long
Public Walktimer As Long
Public tmr25 As Long
Public tmr45 As Long
Public tmr100 As Long
Public tmr10000 As Long
Public maptimer As Long
Public chattmr As Long
Public targettmr As Long
Public fogtmr As Long
Public bartmr As Long
Public AppLooptmr As Long
Public Thread As Boolean

Public AppRunning As Boolean

Public GameState As Byte              '//Controls the current state of the game (In-Game, In-Menu, In-Loading)

'//Game State
Public Enum GameStateEnum
    inMenu = 1
    InLogin
    InLoad
    InGame
End Enum


Public Sub AppLoop()

    Do While AppRunning = True
        Tick = getTime                                  ' Set the inital tick
        ElapsedTime = Tick - FrameTime                  ' Set the time difference for time-based movement
        FrameTime = Tick                                ' Set the time second loop time to the first.

        ' Mute everything but still keep everything playing
        If frmMain.WindowState = vbMinimized Then
            Stop_Music
        End If

        Select Case GameState
        Case GameStateEnum.inMenu, GameStateEnum.InLogin, GameStateEnum.InLoad
            MenuLoop
        Case GameStateEnum.InGame
            GameLoop
        End Select
    Loop

    Call DestroyGame
End Sub

Public Sub GameLoop()
    Dim i As Long, X As Long, Y As Long
    Dim barDifference As Long

    If Thread = False Then
        AppLooptmr = Tick + 25
    End If

    ' handle input
    If GetForegroundWindow() = frmMain.hWnd Then
        HandleMouseInput
    End If

    ' * Check surface timers *
    ' Sprites
    If tmr10000 < Tick Then
        ' check ping
        Call GetPing
        tmr10000 = Tick + 10000
    End If

    For i = 1 To LastProjectile
        Call ProcessProjectile(i)
    Next i

    If tmr25 < Tick Then
        Call CheckKeys    ' Check to make sure they aren't trying to auto do anything

        If GetForegroundWindow() = frmMain.hWnd Then
            Call CheckInputKeys    ' Check which keys were pressed
        End If

        ' check if we need to end the CD icon
        If CountSpellicon > 0 Then
            For i = 1 To MAX_PLAYER_SPELLS
                If PlayerSpells(i).Spell > 0 Then
                    If SpellCD(i) > 0 Then
                        If SpellCD(i) + (Spell(PlayerSpells(i).Spell).CDTime * 1000) < Tick Then
                            SpellCD(i) = 0
                        End If
                    End If
                End If
            Next
        End If

        ' check if we need to unlock the player's spell casting restriction
        If SpellBuffer > 0 Then
            If SpellBuffertimer + (Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000) < Tick Then
                SpellBuffer = 0
                SpellBuffertimer = 0
                ClearPlayerFrame MyIndex

                Player(MyIndex).ConjureAnimProjectileType = ProjectileTypeEnum.None
                Player(MyIndex).ConjureAnimProjectileNum = 0
            End If
        End If

        If CanMoveNow Then
            Call ProcessPlayerActions
        End If

        For i = 1 To MAX_BYTE
            CheckAnimInstance i
        Next

        ' appear tile logic
        AppearTileFadeLogic
        CheckAppearTiles

        tmr25 = Tick + 25
    End If

    ' targetting
    If targettmr < Tick Then
        If tabDown Then
            FindNearestTarget
        End If

        targettmr = Tick + 50
    End If

    ' chat timer
    If chattmr < Tick Then
        ' scrolling
        If ChatButtonUp Then
            ScrollChatBox 0
        End If

        If ChatButtonDown Then
            ScrollChatBox 1
        End If

        ' remove messages
        If chatLastRemove + CHAT_DIFFERENCE_TIMER < getTime Then
            ' remove timed out messages from chat
            For i = Chat_HighIndex To 1 Step -1
                If Len(Chat(i).text) > 0 Then
                    If Chat(i).visible Then
                        If Chat(i).timer + CHAT_TIMER < Tick Then
                            Chat(i).visible = False
                            chatLastRemove = getTime
                            Exit For
                        End If
                    End If
                End If
            Next
        End If

        chattmr = Tick + 50
    End If

    If tmr45 <= Tick Then
        For i = 1 To LastProjectile
            Call ProcessProjectileCurAnimation(i)
        Next

        tmr45 = Tick + 45
    End If

    ' fog scrolling
    If fogtmr < Tick Then
        If CurrentFogSpeed > 0 Then
            ' move
            fogOffsetX = fogOffsetX - 1
            fogOffsetY = fogOffsetY - 1

            ' reset
            If fogOffsetX < -256 Then fogOffsetX = 0
            If fogOffsetY < -256 Then fogOffsetY = 0

            ' reset timer
            fogtmr = Tick + 255 - CurrentFogSpeed
        End If
    End If

    ' elastic bars
    If bartmr < Tick Then
        SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
        SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
        SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).Num > 0 Then
                SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
            End If
        Next

        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
            End If
        Next

        ' reset timer
        bartmr = Tick + 10
    End If

    ' Animations!
    If maptimer < Tick Then

        If Not IsConnected Then GameState = GameStateEnum.inMenu

        ' animate waterfalls
        Select Case waterfallFrame

        Case 0
            waterfallFrame = 1

        Case 1
            waterfallFrame = 2

        Case 2
            waterfallFrame = 0
        End Select

        ' animate autotiles
        Select Case autoTileFrame

        Case 0
            autoTileFrame = 1

        Case 1
            autoTileFrame = 2

        Case 2
            autoTileFrame = 0
        End Select

        ' animate textbox
        If chatShowLine = "|" Then
            chatShowLine = vbNullString
        Else
            chatShowLine = "|"
        End If

        ' re-set timer
        maptimer = Tick + 500
    End If

    Call ProcessWeather

    ' Process input before rendering, otherwise input will be behind by 1 frame
    If Walktimer < Tick Then

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call ProcessMovement(i)
            End If

        Next i

        ' Process npc movements (actually move them)
        For i = 1 To Npc_HighIndex

            If Map.MapData.Npc(i) > 0 Then
                Call ProcessNpcMovement(i)
            End If

        Next i

        Walktimer = Tick + 30    ' edit this value to change WalkTimer
    End If

    ' *********************
    ' ** Render Graphics **
    ' *********************
    If Thread = False Then
        Call Render_Graphics
        Call UpdateSounds

        ' Lock fps
        If Not Options.FPSLock Then
            Do While getTime < Tick + 15
                GoPeekMessage
                Sleep 1
            Loop

        End If

        ' Calculate fps
            If TickFPS < Tick Then
                GameFPS = FPS
                TickFPS = Tick + 1000
                FPS = 0
            Else
                FPS = FPS + 1
            End If
    End If

    GoPeekMessage

    If Thread And AppLooptmr > Tick Then
        Thread = False
        Exit Sub
    End If

    If GameState = GameStateEnum.inMenu Then
        If isLogging Then
            isLogging = False
            GettingMap = True
            Stop_Music
            Play_Music MenuMusic
        Else
            ' Shutdown the game
            Call SetStatus("Destroying game data.")
            AppRunning = False
        End If
    End If
End Sub

Private Sub MenuLoop()
    Dim FPS As Long, tmr500 As Single, fadetmr As Long
    ' handle input
    If GetForegroundWindow() = frmMain.hWnd Then
        HandleMouseInput
    End If

    If Thread = False Then
        AppLooptmr = Tick + 25
    End If

    ' Animations!
    If tmr500 < Tick Then
        ' animate textbox
        If chatShowLine = "|" Then
            chatShowLine = vbNullString
        Else
            chatShowLine = "|"
        End If

        ' re-set timer
        tmr500 = Tick + 500
    End If

    ' trailer
    If videoPlaying Then VideoLoop

    ' fading
    If fadetmr < Tick Then
        If Not videoPlaying Then
            If fadeAlpha > 5 Then
                ' lower fade
                fadeAlpha = fadeAlpha - 5
            Else
                fadeAlpha = 0
            End If
        End If
        fadetmr = Tick + 1
    End If

    ' *********************
    ' ** Render Graphics **
    ' *********************
    If Thread = False Then
        Call Render_Menu
        Call UpdateSounds

        ' Lock fps
        If Not Options.FPSLock Then
            Do While getTime < Tick + 15
                GoPeekMessage
                Sleep 1
            Loop

        End If

        ' Calculate fps
            If TickFPS < Tick Then
                GameFPS = FPS
                TickFPS = Tick + 1000
                FPS = 0
            Else
                FPS = FPS + 1
            End If
    End If

    GoPeekMessage

    If Thread And AppLooptmr > Tick Then
        Thread = False
        Exit Sub
    End If
End Sub
