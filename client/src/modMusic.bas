Attribute VB_Name = "modMusic"
Option Explicit

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"
Public Const Sound_Thunder As String = "Thunder.wav"

Public INIT_MUSIC As Boolean
Public INIT_SOUND As Boolean

Public lastButtonSound As Long
Public lastNpcChatsound As Long

Private Type SoundRec
    Handle As Long
    X As Long
    Y As Long
    Loop As Long
    Channel As Long
End Type

Public CurrentSound As Long
Public Const Max_Sounds As Long = 30
Public Sounds(1 To Max_Sounds) As SoundRec

Public curSong As String
Private songHandle As Long
Private streamHandle As Long
Public ResetSounds As Boolean


Public Function InitFmod() As Boolean
Dim result As Boolean

    On Error GoTo ErrorHandler
    
    If App.LogMode = 0 Then Exit Function
    
    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)
    If Not result Then GoTo ErrorHandler
    
    ' return positive
    InitFmod = True
    INIT_MUSIC = True
    INIT_SOUND = True
    ReDim MapSounds(0)
    Exit Function
    
ErrorHandler:
    InitFmod = False
    INIT_MUSIC = False
    INIT_SOUND = False
End Function

Public Sub Destroy_Music()
    If INIT_MUSIC = False And INIT_SOUND = False Then Exit Sub
    ' destroy music engine
    Stop_Music
    FSOUND_Close
    INIT_MUSIC = False
    curSong = vbNullString
End Sub

Public Sub Play_Music(ByVal song As String)
    If Not INIT_MUSIC Then Exit Sub
    
    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub
    
    ' does it exist?
    If Not FileExist(App.Path & MUSIC_PATH & song) Then Exit Sub
    
    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    
    ' stop the existing music
    Stop_Music
    
    ' find the extension
    Select Case Right$(song, 4)
        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.Path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, 150
            
        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.Path & MUSIC_PATH & song, FSOUND_LOOP_normal, 0, 0)
            ' play it
            FSOUND_Stream_Play 0, streamHandle
            ' set volume
            FSOUND_SetVolume streamHandle, 150
        Case Else
            Exit Sub
    End Select
    
    ' new current song
    curSong = song
End Sub

Public Sub Stop_Music()
    If Not streamHandle = 0 Then
        ' stop stream
        FSOUND_Stream_Stop streamHandle
        ' destroy
        FSOUND_Stream_Close streamHandle
        streamHandle = 0
    End If
    
    If Not songHandle = 0 Then
        ' stop song
        FMUSIC_StopSong songHandle
        ' destroy
        FMUSIC_FreeSong songHandle
        songHandle = 0
    End If
    
    ' no music
    curSong = vbNullString
End Sub

Public Sub StopAllSounds()
    Dim i As Long
    For i = 1 To Max_Sounds
        If Sounds(i).Handle > 0 Then
            FSOUND_StopSound Sounds(i).Channel
            FSOUND_Sample_Free Sounds(i).Handle
        End If
    Next
    
    RemoveAllMapSounds
End Sub

Public Sub Play_Sound(ByVal sound As String, X As Long, Y As Long)
    Dim volume As Byte
    If Not INIT_SOUND Then Exit Sub
    If Options.sound = 0 Then Exit Sub
    
    If Not FileExist(App.Path & SOUND_PATH & sound) Then Exit Sub
    
    ' load the sound
    Sounds(CurrentSound).Handle = LoadSound(sound)
    If Sounds(CurrentSound).Handle > 0 Then
        Sounds(CurrentSound).X = X
        Sounds(CurrentSound).Y = Y
        volume = CalculateSoundVolume(Sounds(CurrentSound).X, Sounds(CurrentSound).Y)
        Sounds(CurrentSound).Channel = FSOUND_PlaySound(FSOUND_FREE, Sounds(CurrentSound).Handle)
        FSOUND_SetVolume Sounds(CurrentSound).Channel, volume
    Else
        CurrentSound = CurrentSound - 1
    End If
End Sub

Public Function LoadSound(ByVal sound As String, Optional MapSound As Boolean = False) As Long
Dim i As Long
    If INIT_SOUND = False Then Exit Function
    If Options.sound = 0 Then Exit Function
    
    If Not MapSound Then
        CurrentSound = CurrentSound + 1
        If CurrentSound > Max_Sounds Then
            FSOUND_Sample_Free Sounds(1).Handle
            CurrentSound = 1
            ResetSounds = True
            For i = 2 To Max_Sounds
                FSOUND_Sample_Free Sounds(i).Handle
                Sounds(i).Handle = 0
            Next
        End If
    End If
    
    If MapSound Then
        LoadSound = FSOUND_Sample_Load(FSOUND_FREE, App.Path & SOUND_PATH & sound, FSOUND_LOOP_normal, 0, 0)
    Else
        LoadSound = FSOUND_Sample_Load(FSOUND_FREE, App.Path & SOUND_PATH & sound, FSOUND_normal, 0, 0)
    End If
End Function

Function CalculateSoundVolume(SoundX As Long, SoundY As Long) As Byte
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, distance As Double, volume As Long
    If GameState <> GameStateEnum.InGame Then CalculateSoundVolume = 150: Exit Function
    If INIT_SOUND = False Then Exit Function
    If Options.sound = 0 Then Exit Function
    
    If SoundX > -1 Or SoundY > -1 Then
        X1 = (Player(MyIndex).X * 32) + Player(MyIndex).xOffset
        Y1 = (Player(MyIndex).Y * 32) + Player(MyIndex).yOffset
        X2 = (SoundX * 32) + 16
        Y2 = (SoundY * 32) + 16
        If ((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) < 0 Then
            distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) * -1)
        Else
            distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))
        End If
        If distance >= 256 Then
            CalculateSoundVolume = 0
        Else
            CalculateSoundVolume = 150 - ((distance / 256) * 150)
            If CalculateSoundVolume = 0 Then CalculateSoundVolume = 1
        End If
    Else
        CalculateSoundVolume = 150
    End If
    
End Function

Sub StopSound(soundindex As Long)
    FSOUND_StopSound Sounds(soundindex).Channel
End Sub

Sub RemoveAllMapSounds()
Dim i As Long
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            MapSounds(i).InUse = False
            FSOUND_StopSound MapSounds(i).Channel
            FSOUND_Sample_Free MapSounds(i).SoundHandle
        Next
        MapSoundCount = 0
        ReDim MapSounds(0)
    End If
End Sub

Sub UpdateSounds()
If INIT_SOUND = False Then Exit Sub
If Options.sound = 0 Then Exit Sub
Dim i As Long, X As Long
    For i = 1 To Max_Sounds
        If Sounds(i).Handle <> 0 Then
            Call FSOUND_SetVolume(Sounds(i).Channel, CalculateSoundVolume(Sounds(i).X, Sounds(i).Y))
        End If
    Next
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            FSOUND_SetVolume MapSounds(i).Channel, CalculateSoundVolume(MapSounds(i).X, MapSounds(i).Y)
        Next
    End If
End Sub

Sub CacheNewMapSounds()
If INIT_SOUND = False Then Exit Sub
If Options.sound = 0 Then Exit Sub
    Dim i As Long, X As Long, Y As Long
    RemoveAllMapSounds
    If Trim$(Map.MapData.Music) <> Null Then
        MapSoundCount = MapSoundCount + 1
        ReDim Preserve MapSounds(MapSoundCount)
        MapSounds(MapSoundCount).SoundHandle = LoadSound(Trim$(Map.MapData.Music), True)
        FSOUND_SetLoopMode MapSounds(MapSoundCount).SoundHandle, FSOUND_LOOP_normal
        MapSounds(MapSoundCount).X = -1
        MapSounds(MapSoundCount).Y = -1
        MapSounds(MapSoundCount).Channel = FSOUND_PlaySound(FSOUND_FREE, MapSounds(MapSoundCount).SoundHandle)
        FSOUND_SetVolume MapSounds(MapSoundCount).Channel, CalculateSoundVolume(MapSounds(MapSoundCount).X, MapSounds(MapSoundCount).Y)
    End If
    For X = 0 To Map.MapData.maxX
        For Y = 0 To Map.MapData.maxY
            If Map.TileData.Tile(X, Y).Type = TILE_TYPE_SOUND Then
                MapSoundCount = MapSoundCount + 1
                ReDim Preserve MapSounds(MapSoundCount)
                MapSounds(MapSoundCount).SoundHandle = LoadSound(Trim$(Map.TileData.Tile(X, Y).Data4), True)
                FSOUND_SetLoopMode MapSounds(MapSoundCount).SoundHandle, FSOUND_LOOP_normal
                MapSounds(MapSoundCount).X = X
                MapSounds(MapSoundCount).Y = Y
                MapSounds(MapSoundCount).Channel = FSOUND_PlaySound(FSOUND_FREE, MapSounds(MapSoundCount).SoundHandle)
                FSOUND_SetVolume MapSounds(MapSoundCount).Channel, CalculateSoundVolume(MapSounds(MapSoundCount).X, MapSounds(MapSoundCount).Y)
            End If
        Next
    Next

End Sub


