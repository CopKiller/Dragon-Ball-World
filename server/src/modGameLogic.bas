Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, X, Y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).playerName = playerName
            MapItem(MapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).X = X
            MapItem(MapNum, i).Y = Y
            MapItem(MapNum, i).Bound = isSB
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For X = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).TileData.Tile(X, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).TileData.Tile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).TileData.Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).TileData.Tile(X, Y).Data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).TileData.Tile(X, Y).Data1, Map(MapNum).TileData.Tile(X, Y).Data2, MapNum, X, Y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    npcNum = Map(MapNum).MapData.Npc(mapNpcNum)

    If npcNum > 0 Then
    
        With MapNpc(MapNum).Npc(mapNpcNum)
            .Num = npcNum
            .Target = 0
            .TargetType = 0 ' clear
            .Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
            .Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
            .Dir = Int(Rnd * 4)
            .spellBuffer.Spell = 0
            .spellBuffer.Timer = 0
            .spellBuffer.Target = 0
            .spellBuffer.tType = 0
        
            'Check if theres a spawn tile for the specific npc
            For X = 0 To Map(MapNum).MapData.MaxX
                For Y = 0 To Map(MapNum).MapData.MaxY
                    If Map(MapNum).TileData.Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                        If Map(MapNum).TileData.Tile(X, Y).Data1 = mapNpcNum Then
                            .X = X
                            .Y = Y
                            .Dir = Map(MapNum).TileData.Tile(X, Y).Data2
                            Spawned = True
                            Exit For
                        End If
                    End If
                Next Y
            Next X
            
            If Not Spawned Then
        
                ' Well try 100 times to randomly place the sprite
                For i = 1 To 100
                    X = Random(0, Map(MapNum).MapData.MaxX)
                    Y = Random(0, Map(MapNum).MapData.MaxY)
        
                    If X > Map(MapNum).MapData.MaxX Then X = Map(MapNum).MapData.MaxX
                    If Y > Map(MapNum).MapData.MaxY Then Y = Map(MapNum).MapData.MaxY
        
                    ' Check if the tile is walkable
                    If NpcTileIsOpen(MapNum, X, Y) Then
                        .X = X
                        .Y = Y
                        Spawned = True
                        Exit For
                    End If
        
                Next
                
            End If
    
            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
    
                For X = 0 To Map(MapNum).MapData.MaxX
                    For Y = 0 To Map(MapNum).MapData.MaxY
    
                        If NpcTileIsOpen(MapNum, X, Y) Then
                            .X = X
                            .Y = Y
                            Spawned = True
                        End If
    
                    Next
                Next
    
            End If
    
            ' If we suceeded in spawning then send it to everyone
            If Spawned Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSpawnNpc
                Buffer.WriteLong mapNpcNum
                Buffer.WriteLong .Num
                Buffer.WriteLong .X
                Buffer.WriteLong .Y
                Buffer.WriteLong .Dir
                
                SendDataToMap MapNum, Buffer.ToArray()
                Buffer.Flush: Set Buffer = Nothing
            End If
            
            SendMapNpcVitals MapNum, mapNpcNum
        End With
    End If
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).X = X Then
                If MapNpc(MapNum).Npc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).TileData.Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).TileData.Tile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).TileData.Tile(X, Y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim N As Long
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Function
    End If

    X = MapNpc(MapNum).Npc(mapNpcNum).X
    Y = MapNpc(MapNum).Npc(mapNpcNum).Y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If Y > 0 Then
                N = Map(MapNum).TileData.Tile(X, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MapData.MaxY Then
                N = Map(MapNum).TileData.Tile(X, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If X > 0 Then
                N = Map(MapNum).TileData.Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MapData.MaxX Then
                N = Map(MapNum).TileData.Tile(X + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If Y > 0 And X > 0 Then
                N = Map(MapNum).TileData.Tile(X - 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If Y > 0 And X < Map(MapNum).MapData.MaxX Then
                N = Map(MapNum).TileData.Tile(X + 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_DOWN_LEFT

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MapData.MaxY And X > 0 Then
                N = Map(MapNum).TileData.Tile(X - 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
'#######################################################################################################################
'#######################################################################################################################
        Case DIR_DOWN_RIGHT

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MapData.MaxY And X < Map(MapNum).MapData.MaxX Then
                N = Map(MapNum).TileData.Tile(X + 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).TileData.Tile(MapNpc(MapNum).Npc(mapNpcNum).X, MapNpc(MapNum).Npc(mapNpcNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(mapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1
        Case DIR_DOWN
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1
        Case DIR_LEFT
            MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1
        Case DIR_UP_LEFT
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1: MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1
        Case DIR_UP_RIGHT
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y - 1: MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1
        Case DIR_DOWN_LEFT
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1: MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X - 1
        Case DIR_DOWN_RIGHT
            MapNpc(MapNum).Npc(mapNpcNum).Y = MapNpc(MapNum).Npc(mapNpcNum).Y + 1: MapNpc(MapNum).Npc(mapNpcNum).X = MapNpc(MapNum).Npc(mapNpcNum).X + 1
    End Select

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcMove
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).X
    Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Y
    Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
    Buffer.WriteLong movement
    
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    
    SendDataToMap MapNum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim N As Long
    N = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            N = N + 1
        End If

    Next

    GetTotalMapPlayers = N
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim Y As Long
    Dim X As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)

    For X = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY
            TempTile(MapNum).DoorOpen(X, Y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim X As Long, Y As Long, Resource_Count As Long
    Resource_Count = 0

    For X = 0 To Map(MapNum).MapData.MaxX
        For Y = 0 To Map(MapNum).MapData.MaxY

            If Map(MapNum).TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).TileData.Tile(X, Y).Data1).health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    oldBound = Player(Index).Inv(oldSlot).Bound
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    newBound = Player(Index).Inv(newSlot).Bound
    
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    Player(Index).Inv(newSlot).Bound = oldBound
    
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    Player(Index).Inv(oldSlot).Bound = newBound
    
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = Player(Index).Spell(oldSlot).Spell
    NewNum = Player(Index).Spell(newSlot).Spell
    OldUses = Player(Index).Spell(oldSlot).Uses
    NewUses = Player(Index).Spell(newSlot).Uses
    
    Player(Index).Spell(oldSlot).Spell = NewNum
    Player(Index).Spell(oldSlot).Uses = NewUses
    Player(Index).Spell(newSlot).Spell = OldNum
    Player(Index).Spell(newSlot).Uses = OldUses
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0, , True
        PlayerMsg Index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim partynum As Long, i As Long

    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' find out how many members we have
        Party_CountMembers partynum
        ' make sure there's more than 2 people
        If Party(partynum).MemberCount > 2 Then
            ' check if leader
            If Party(partynum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) > 0 And Party(partynum).Member(i) <> Index Then
                        Party(partynum).Leader = Party(partynum).Member(i)
                        PartyMsg partynum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = Index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg partynum, GetPlayerName(Index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = Index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partynum
            ' only 2 people, disband
            PartyMsg partynum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(partynum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty partynum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, "This player has an outstanding party invitation already.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        partynum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(partynum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, Index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, "Invitation sent.", Green
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, Index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, "Invitation sent.", Green
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long, X As Long

    If Index = 0 Then Exit Sub
    
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If TempPlayer(targetPlayer).inParty > 0 Then
        PlayerMsg Index, Trim$(GetPlayerName(targetPlayer)) & " is already in a party.", BrightRed
        PlayerMsg targetPlayer, "You're already in a party.", BrightRed
        Exit Sub
    End If

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        partynum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) = 0 Then
                'add to the party
                Party(partynum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partynum
                ' send everyone's data to everyone
                SendPlayerData_Party partynum
                ' send update to all - including new player
                SendPartyUpdate partynum
                ' Send party vitals to everyone again
                For X = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(X) > 0 Then
                        SendPartyVitals partynum, Party(partynum).Member(X)
                    End If
                Next
                ' let everyone know they've joined
                PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partynum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg Index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partynum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partynum).MemberCount = 2
        Party(partynum).Leader = Index
        Party(partynum).Member(1) = Index
        Party(partynum).Member(2) = targetPlayer
        SendPlayerData_Party partynum
        SendPartyUpdate partynum
        SendPartyVitals partynum, Index
        SendPartyVitals partynum, targetPlayer
        ' let them know it's created
        PartyMsg partynum, "Party created.", BrightGreen
        PartyMsg partynum, GetPlayerName(Index) & " has joined the party.", Pink
        PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = partynum
        TempPlayer(targetPlayer).inParty = partynum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal targetPlayer As Long)
    If Not IsConnected(Index) Or Not IsPlaying(Index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(Index).TradeRequest = 0
        Exit Sub
    End If
    
    PlayerMsg Index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partynum As Long)
Dim i As Long, highIndex As Long, X As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partynum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partynum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(partynum).Member(X) = Party(partynum).Member(X + 1)
                    Party(partynum).Member(X + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partynum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partynum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partynum
End Sub

Public Sub Party_ShareExp(ByVal partynum As Long, ByVal exp As Long, ByVal Index As Long, Optional ByVal enemyLevel As Long = 0)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    If Party(partynum).MemberCount <= 0 Then Exit Sub

    ' check if it's worth sharing
    If Not exp >= Party(partynum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, exp, enemyLevel
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partynum).MemberCount
    leftOver = exp Mod Party(partynum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partynum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare, enemyLevel
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partynum).Member(RAND(1, Party(partynum).MemberCount))
    ' give the exp
    If leftOver > 0 Then GivePlayerEXP tmpIndex, leftOver, enemyLevel
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal exp As Long, Optional ByVal enemyLevel As Long = 0)
Dim multiplier As Long, partynum As Long, expBonus As Long
    ' no exp
    If exp = 0 Then Exit Sub
    ' rte9
    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    ' make sure we're not max level
    If Not GetPlayerLevel(Index) >= MAX_LEVELS Then
        ' check for exp deduction
        If enemyLevel > 0 Then
            ' exp deduction
            If enemyLevel <= GetPlayerLevel(Index) - 3 Then
                ' 3 levels lower, exit out
                Exit Sub
            ElseIf enemyLevel <= GetPlayerLevel(Index) - 2 Then
                ' half exp if enemy is 2 levels lower
                exp = exp / 2
            End If
        End If
        ' check if in party
        partynum = TempPlayer(Index).inParty
        If partynum > 0 Then
            If Party(partynum).MemberCount > 1 Then
                multiplier = Party(partynum).MemberCount - 1
                ' multiply the exp
                expBonus = (exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                exp = exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerExp(Index, GetPlayerExp(Index) + exp)
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "+" & exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp Index
    Else
        Call SetPlayerExp(Index, 0)
        SendEXP Index
    End If
End Sub

Public Sub Unique_Item(ByVal Index As Long, ByVal ItemNum As Long)
Dim ClassNum As Long, i As Long

    Select Case Item(ItemNum).Data1
        Case 1 ' Reset Stats
            ClassNum = GetPlayerClass(Index)
            If ClassNum <= 0 Or ClassNum > Max_Classes Then Exit Sub
            ' re-set the actual stats to class defaults
            For i = 1 To Stats.Stat_Count - 1
                SetPlayerStat Index, i, Class(ClassNum).Stat(i)
            Next
            ' give player their points back
            SetPlayerPOINTS Index, (GetPlayerLevel(Index) - 1) * 3
            ' take item
            TakeInvItem Index, ItemNum, 1
            ' let them know we've done it
            PlayerMsg Index, "Your stats have been reset.", BrightGreen
            ' send them their new stats
            SendPlayerData Index
        Case Else ' Exit out otherwise
            Exit Sub
    End Select
End Sub

Public Function hasProficiency(ByVal Index As Long, ByVal proficiency As Long) As Boolean
    Select Case proficiency
        Case 0 ' None
            hasProficiency = True
            Exit Function
        Case 1 ' Heavy
            If GetPlayerClass(Index) = 1 Then
                hasProficiency = True
                Exit Function
            End If
        Case 2 ' Light
            If GetPlayerClass(Index) = 2 Or GetPlayerClass(Index) = 3 Then
                hasProficiency = True
                Exit Function
            End If
    End Select
    hasProficiency = False
End Function

Public Sub CheckProjectile(ByVal i As Long)
    Dim Angle As Long, X As Long, Y As Long, N As Long
    Dim Attacker As Long, spellNum As Long
    Dim BaseDamage As Long, Damage As Long

    If i < 0 Or i > MAX_PROJECTILE_MAP Then Exit Sub
    If MapProjectile(i).OwnerType = TARGET_TYPE_PLAYER Then
        If Not IsPlaying(MapProjectile(i).Owner) Then: Call ClearProjectile(i): Exit Sub
    ElseIf MapProjectile(i).OwnerType = TARGET_TYPE_NPC Then
        If MapNpc(MapProjectile(i).MapNum).Npc(MapProjectile(i).Owner).Num = 0 Then: Call ClearProjectile(i): Exit Sub
    End If

    Attacker = MapProjectile(i).Owner
    spellNum = MapProjectile(i).spellNum
    BaseDamage = Spell(spellNum).Vital
    Damage = BaseDamage + Int(GetPlayerStat(Attacker, Intelligence) / 3)

    ' ****** Create Particle ******
    With MapProjectile(i)
        If .Graphic > 0 Then
            If .Speed < 5000 Then

                ' ****** Update Position ******
                Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tX, .tY)
                .X = .X + (Sin(Angle) * ElapsedTime * (.Speed / 1000))
                .Y = .Y - (Cos(Angle) * ElapsedTime * (.Speed / 1000))

                If Spell(spellNum).IsAoE Then
                    Select Case MapProjectile(i).direction
                    Case DIR_UP
                        .xTargetAoE = .X - (Int(Spell(MapProjectile(i).spellNum).DirectionAoE(DIR_UP + 1).X / 2) * PIC_X)
                        .yTargetAoE = .Y
                    Case DIR_DOWN
                        .xTargetAoE = .X - (Int(Spell(MapProjectile(i).spellNum).DirectionAoE(DIR_DOWN + 1).X / 2) * PIC_X)
                        .yTargetAoE = .Y
                    Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                        .xTargetAoE = .X
                        .yTargetAoE = .Y - (Int(Spell(MapProjectile(i).spellNum).DirectionAoE(DIR_LEFT + 1).Y / 2) * PIC_Y)
                    Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                        .xTargetAoE = .X
                        .yTargetAoE = .Y - (Int(Spell(MapProjectile(i).spellNum).DirectionAoE(DIR_RIGHT + 1).Y / 2) * PIC_Y)
                    End Select
                End If
            End If
        End If
    End With

    ' ****** Erase Projectile ******    Seperate Loop For Erasing
    If MapProjectile(i).OwnerType = TARGET_TYPE_PLAYER Then
        ' VERIFICA TYLE_BLOCK e TYLE_RESOURCE
        For X = 0 To Map(GetPlayerMap(Attacker)).MapData.MaxX
            For Y = 0 To Map(GetPlayerMap(Attacker)).MapData.MaxY
                If Map(GetPlayerMap(Attacker)).TileData.Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(Attacker)).TileData.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                    If Abs(MapProjectile(i).X - (X * PIC_X)) < 20 Then
                        If Abs(MapProjectile(i).Y - (Y * PIC_Y)) < 20 Then
                            Call ClearProjectile(i)
                            Exit Sub
                        End If
                    End If
                End If
            Next Y
        Next X

        If Not Spell(MapProjectile(i).spellNum).IsAoE Then
            ' VERIFICA PLAYER NO CAMINHO
            For N = 1 To Player_HighIndex
                If IsPlaying(N) Then
                    If N <> Attacker Then
                        If Abs(MapProjectile(i).X - (GetPlayerX(N) * PIC_X)) < 20 Then
                            If Abs(MapProjectile(i).Y - (GetPlayerY(N) * PIC_Y)) < 20 Then
                                If CanPlayerAttackPlayer(Attacker, N, True) Then
                                    If MapProjectile(i).Speed <> 6000 Then
                                        If Spell(spellNum).Projectile.ImpactRange > 0 Then
                                            Call MakeImpact(N, Spell(spellNum).Projectile.ImpactRange, TARGET_TYPE_PLAYER, GetPlayerMap(Attacker), Attacker, False)
                                        End If
                                        PlayerAttackPlayer Attacker, N, Damage, spellNum
                                        Call ClearProjectile(i)
                                        Exit Sub
                                    Else
                                        If tick > MapProjectile(i).Duration Then
                                            If Spell(spellNum).Projectile.ImpactRange > 0 Then
                                                Call MakeImpact(N, Spell(spellNum).Projectile.ImpactRange, TARGET_TYPE_PLAYER, GetPlayerMap(Attacker), Attacker, False)
                                            End If
                                            PlayerAttackPlayer Attacker, N, Damage, spellNum
                                            MapProjectile(i).Duration = tick + 1000
                                        End If
                                    End If
                                Else
                                    If MapProjectile(i).Speed <> 6000 Then
                                        Call ClearProjectile(i)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            ' VERIFICA NPC NO CAMINHO
            For N = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(Attacker)).Npc(N).Num <> 0 Then
                    If Abs(MapProjectile(i).X - (MapNpc(GetPlayerMap(Attacker)).Npc(N).X * PIC_X)) < 20 Then
                        If Abs(MapProjectile(i).Y - (MapNpc(GetPlayerMap(Attacker)).Npc(N).Y * PIC_Y)) < 20 Then
                            If CanPlayerAttackNpc(Attacker, N, True) Then
                                If MapProjectile(i).Speed <> 6000 Then
                                    If Spell(spellNum).Projectile.ImpactRange > 0 Then
                                        Call MakeImpact(N, Spell(spellNum).Projectile.ImpactRange, TARGET_TYPE_NPC, GetPlayerMap(Attacker), Attacker, False)
                                    End If
                                    PlayerAttackNpc Attacker, N, Damage, spellNum
                                    Call ClearProjectile(i)
                                    Exit Sub
                                Else
                                    If tick > MapProjectile(i).Duration Then
                                        If Spell(spellNum).Projectile.ImpactRange > 0 Then
                                            Call MakeImpact(N, Spell(spellNum).Projectile.ImpactRange, TARGET_TYPE_NPC, GetPlayerMap(Attacker), Attacker, False)
                                        End If
                                        PlayerAttackNpc Attacker, N, Damage, spellNum
                                        MapProjectile(i).Duration = tick + 1000
                                    End If
                                End If
                            Else
                                If MapProjectile(i).Speed <> 6000 Then
                                    Call ClearProjectile(i)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Else    ' SE É DANO EM AREA If Not Spell(MapProjectile(i).spellNum).IsAoE Then
            ' VERIFICA NPC NO CAMINHO
            For N = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(Attacker)).Npc(N).Num <> 0 Then
                    If Abs((MapNpc(GetPlayerMap(Attacker)).Npc(N).X * PIC_X) - MapProjectile(i).X) < (20 * Spell(MapProjectile(i).spellNum).DirectionAoE(MapProjectile(i).direction + 1).X) Then
                        If Abs((MapNpc(GetPlayerMap(Attacker)).Npc(N).Y * PIC_Y) - MapProjectile(i).Y) < (20 * Spell(MapProjectile(i).spellNum).DirectionAoE(MapProjectile(i).direction + 1).Y) Then
                            If CanPlayerAttackNpc(Attacker, N, True) Then
                                If Spell(MapProjectile(i).spellNum).Projectile.RecuringDamage Then
                                    If tick > MapProjectile(i).AttackTimer(N) Then
                                        If Spell(spellNum).Projectile.ImpactRange > 0 Then
                                            Call MakeImpact(N, Spell(spellNum).Projectile.ImpactRange, TARGET_TYPE_NPC, GetPlayerMap(Attacker), Attacker, False)
                                        End If
                                        PlayerAttackNpc Attacker, N, Damage, spellNum
                                        MapProjectile(i).AttackTimer(N) = tick + MapProjectile(i).Speed
                                        'Call PlayerMsg(Attacker, "ID: " & N, White)
                                    End If
                                Else
                                    ' FALTA IMPLEMENTAR
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If

        ' VERIFICA SE CHEGOU AO ALVO
        If Abs(MapProjectile(i).X - MapProjectile(i).tX) < 20 Then
            If Abs(MapProjectile(i).Y - MapProjectile(i).tY) < 20 Then
                If MapProjectile(i).Speed <> 6000 Then
                    Call ClearProjectile(i)
                    Exit Sub
                End If
            End If
        End If

        ' VERIFICAR SE É UMA TRAP E O TEMPO DE SPAWN ACABOU
        If MapProjectile(i).Speed >= 5000 Then
            If tick >= MapProjectile(i).Duration Then
                Call ClearProjectile(i)
                Exit Sub
            End If
        End If

    End If
End Sub

Function SecondsToHMS(ByRef Segundos As Long) As String
    Dim HR As Long, ms As Long, Ss As Long, MM As Long
    Dim Total As Long, Count As Long

    If Segundos = 0 Then
        SecondsToHMS = "0s "
        Exit Function
    End If

    HR = (Segundos \ 3600)
    MM = (Segundos \ 60)
    Ss = Segundos
    'ms = (Segundos * 10)

    ' Pega o total de segundos pra trabalharmos melhor na variavel!
    Total = Segundos

    ' Verifica se tem mais de 1 hora em segundos!
    If HR > 0 Then
        '// Horas
        Do While (Total >= 3600)
            Total = Total - 3600
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = Count & "h "
            Count = 0
        End If
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf MM > 0 Then
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "m "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "s "
            Count = 0
        End If
    ElseIf Ss > 0 Then
        ' Joga na função esse segundo.
        SecondsToHMS = Ss & "s "
        Total = Total - Ss
    End If
End Function
