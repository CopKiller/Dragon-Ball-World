Attribute VB_Name = "Player_Frames"
Option Explicit

Public Enum ProjectileTypeEnum
    None = 0
    KiBall
    GekiDama
    
    ProjectileTypeCount
End Enum

Function GetPlayerFrame(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerFrame = TempPlayer(index).PlayerFrame
End Function

Sub SetPlayerFrame(ByVal index As Long, ByVal frameValue As Long)

    If index > MAX_PLAYERS Then Exit Sub
    TempPlayer(index).PlayerFrame = frameValue
End Sub

Sub SendPlayerFrameToMap(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerFrame
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerFrame(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

End Sub

Sub SendPlayerFrameToMapBut(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerFrame
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerFrame(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

End Sub

Sub ClearPlayerFrameToMap(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    SetPlayerFrame index, 0
    
    Buffer.WriteLong SPlayerFrame
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerFrame(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub ClearPlayerFrameToMapBut(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    SetPlayerFrame index, 0
    
    Buffer.WriteLong SPlayerFrame
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerFrame(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerConjureProjectileCustomToMapBut(ByVal index As Long, _
                                              ByVal projectileType As ProjectileTypeEnum, _
                                              ByVal projectileNum As Long)
    Dim Buffer As clsBuffer
    
    TempPlayer(index).ProjectileCustomType = projectileType
    TempPlayer(index).ProjectileCustomNum = projectileNum

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerConjureProjectileCustom
    Buffer.WriteLong index
    Buffer.WriteLong TempPlayer(index).ProjectileCustomType
    Buffer.WriteLong TempPlayer(index).ProjectileCustomNum
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing

End Sub
