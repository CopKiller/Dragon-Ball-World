Attribute VB_Name = "modMapPng"
Option Explicit

Public TexGround(1 To MAX_MAPS) As Long
Public TexFringe(1 To MAX_MAPS) As Long

Public Const Path_Ground As String = "\data files\graphics\Maps\Ground\"
Public Const Path_Fringe As String = "\data files\graphics\Maps\Fringe\"

Public Sub DrawGround(ByVal MapNum As Integer)
    Dim Width As Integer, Height As Integer
    
    If TexGround(MapNum) = 0 Then Exit Sub

    ' Caso não exista um path pro mapa!
    If mTexture(TexGround(MapNum)).Path = vbNullString Then Exit Sub

    Width = mTexture(TexGround(MapNum)).Width
    Height = mTexture(TexGround(MapNum)).Height

    RenderTexture TexGround(MapNum), ConvertMapX(0), ConvertMapY(0), 0, 0, Width, Height, Width, Height

    'Render box
    'EngineDrawSquare (GlobalX \ PIC_X) * PIC_X, (GlobalY \ PIC_Y) * PIC_Y, 32, 32, DX8Colour(White, 200), 1
End Sub

Public Sub DrawFringe(ByVal MapNum As Integer)
    Dim Width As Integer, Height As Integer
    
    If TexFringe(MapNum) = 0 Then Exit Sub

    ' Caso não exista um path pro mapa!
    If mTexture(TexFringe(MapNum)).Path = vbNullString Then Exit Sub

    Width = mTexture(TexFringe(MapNum)).Width
    Height = mTexture(TexFringe(MapNum)).Height

    RenderTexture TexFringe(MapNum), ConvertMapX(0), ConvertMapY(0), 0, 0, Width, Height, Width, Height
End Sub

Public Sub InitMapPng()
    Dim i As Long
    
    For i = 1 To MAX_MAPS
        TexGround(i) = LoadTextureFile(App.Path & Path_Ground & i)
        TexFringe(i) = LoadTextureFile(App.Path & Path_Fringe & i)
        GoPeekMessage
    Next i
End Sub
