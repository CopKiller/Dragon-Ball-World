Attribute VB_Name = "Map_Database"
Option Explicit

' **********
' ** Maps **
' **********
Public Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String, f As Long, x As Long, y As Long, i As Long
    
    ' save map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    
    ' if it exists then kill the ini
    If FileExist(filename) Then Kill filename
    
    ' General
    With Map(MapNum).MapData
        PutVar filename, "General", "Name", .Name
        PutVar filename, "General", "Music", .Music
        PutVar filename, "General", "Moral", Val(.Moral)
        PutVar filename, "General", "Up", Val(.Up)
        PutVar filename, "General", "Down", Val(.Down)
        PutVar filename, "General", "Left", Val(.left)
        PutVar filename, "General", "Right", Val(.Right)
        PutVar filename, "General", "BootMap", Val(.BootMap)
        PutVar filename, "General", "BootX", Val(.BootX)
        PutVar filename, "General", "BootY", Val(.BootY)
        PutVar filename, "General", "MaxX", Val(.MaxX)
        PutVar filename, "General", "MaxY", Val(.MaxY)
        
        PutVar filename, "General", "Weather", Val(.Weather)
        PutVar filename, "General", "WeatherIntensity", Val(.WeatherIntensity)
        
        PutVar filename, "General", "Fog", Val(.Fog)
        PutVar filename, "General", "FogSpeed", Val(.FogSpeed)
        PutVar filename, "General", "FogOpacity", Val(.FogOpacity)
        
        PutVar filename, "General", "Red", Val(.Red)
        PutVar filename, "General", "Green", Val(.Green)
        PutVar filename, "General", "Blue", Val(.Blue)
        PutVar filename, "General", "Alpha", Val(.Alpha)
        
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With
    
    ' dump tile data
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    f = FreeFile
    
    With Map(MapNum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Put #f, , .TileData.Tile(x, y).Type
                    Put #f, , .TileData.Tile(x, y).Data1
                    Put #f, , .TileData.Tile(x, y).Data2
                    Put #f, , .TileData.Tile(x, y).Data3
                    Put #f, , .TileData.Tile(x, y).Data4
                    Put #f, , .TileData.Tile(x, y).Data5
                    Put #f, , .TileData.Tile(x, y).Autotile
                    Put #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Put #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Put #f, , .TileData.Tile(x, y).Layer(i).x
                        Put #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With

    DoEvents
End Sub

Public Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Public Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist(App.Path & "\Data\maps\map" & i & ".dat") Or Not FileExist(App.Path & "\Data\maps\map" & i & ".ini") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Public Sub LoadMap(MapNum As Long)
    Dim filename As String, i As Long, f As Long, x As Long, y As Long
    
    ' load map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    
    ' General
    With Map(MapNum).MapData
        .Name = GetVar(filename, "General", "Name")
        .Music = GetVar(filename, "General", "Music")
        .Moral = Val(GetVar(filename, "General", "Moral"))
        .Up = Val(GetVar(filename, "General", "Up"))
        .Down = Val(GetVar(filename, "General", "Down"))
        .left = Val(GetVar(filename, "General", "Left"))
        .Right = Val(GetVar(filename, "General", "Right"))
        .BootMap = Val(GetVar(filename, "General", "BootMap"))
        .BootX = Val(GetVar(filename, "General", "BootX"))
        .BootY = Val(GetVar(filename, "General", "BootY"))
        .MaxX = Val(GetVar(filename, "General", "MaxX"))
        .MaxY = Val(GetVar(filename, "General", "MaxY"))
        
        .Weather = Val(GetVar(filename, "General", "Weather"))
        .WeatherIntensity = Val(GetVar(filename, "General", "WeatherIntensity"))
        
        .Fog = Val(GetVar(filename, "General", "Fog"))
        .FogSpeed = Val(GetVar(filename, "General", "FogSpeed"))
        .FogOpacity = Val(GetVar(filename, "General", "FogOpacity"))
        
        .Red = Val(GetVar(filename, "General", "Red"))
        .Green = Val(GetVar(filename, "General", "Green"))
        .Blue = Val(GetVar(filename, "General", "Blue"))
        .Alpha = Val(GetVar(filename, "General", "Alpha"))
        
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With
        
    ' dump tile data
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    f = FreeFile
    
    ' redim the map
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY) As TileRec
    
    With Map(MapNum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Get #f, , .TileData.Tile(x, y).Type
                    Get #f, , .TileData.Tile(x, y).Data1
                    Get #f, , .TileData.Tile(x, y).Data2
                    Get #f, , .TileData.Tile(x, y).Data3
                    Get #f, , .TileData.Tile(x, y).Data4
                    Get #f, , .TileData.Tile(x, y).Data5
                    Get #f, , .TileData.Tile(x, y).Autotile
                    Get #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Get #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Get #f, , .TileData.Tile(x, y).Layer(i).x
                        Get #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
End Sub

Public Sub LoadMaps()
    Dim filename As String, MapNum As Long

    Call CheckMaps

    For MapNum = 1 To MAX_MAPS
        LoadMap MapNum
        ClearTempTile MapNum
        CacheResources MapNum
        DoEvents
    Next
End Sub

Public Sub ClearMap(ByVal MapNum As Long)
    Map(MapNum) = EmptyMap
    Map(MapNum).MapData.Name = vbNullString
    Map(MapNum).MapData.MaxX = MAX_MAPX
    Map(MapNum).MapData.MaxY = MAX_MAPY
    ReDim Map(MapNum).TileData.Tile(0 To Map(MapNum).MapData.MaxX, 0 To Map(MapNum).MapData.MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Public Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

Public Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    MapItem(MapNum, index) = EmptyMapItem
    MapItem(MapNum, index).playerName = vbNullString
End Sub

Public Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Public Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    MapNpc(MapNum) = EmptyMapNpc
End Sub

Public Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Public Sub GetMapCRC32(MapNum As Long)
    Dim Data() As Byte, filename As String, f As Long
    ' map data
    filename = App.Path & "\data\maps\map" & MapNum & ".ini"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        Data = Space$(LOF(f))
        Get #f, , Data
        Close #f
        MapCRC32(MapNum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    If FileExist(filename) Then
        f = FreeFile
        Open filename For Binary As #f
        Data = Space$(LOF(f))
        Get #f, , Data
        Close #f
        MapCRC32(MapNum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(MapNum).MapTileCRC = 0
    End If
End Sub

Public Sub ClearProjectile(ByVal IndexProjectile As Long)
    Dim i As Long
    
    If IndexProjectile < 0 Or IndexProjectile > MAX_PROJECTILE_MAP Then Exit Sub
    MapProjectile(IndexProjectile) = EmptyMapProjectile
    
End Sub
