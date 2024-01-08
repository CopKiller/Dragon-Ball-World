Attribute VB_Name = "Client_Frames"
Option Explicit

Private animTmr(1 To MAX_PLAYERS) As Long
Private animConjure(1 To MAX_PLAYERS) As Byte
Private widthAnim(1 To MAX_PLAYERS) As Long
Private heightAnim(1 To MAX_PLAYERS) As Long

Public Enum ProjectileTypeEnum
    None = 0
    KiBall
    GekiDama
    
    ProjectileTypeCount
End Enum

Function GetPlayerFrame(ByVal index As Long) As Long

    If index > Player_HighIndex Then Exit Function
    GetPlayerFrame = Player(index).playerFrame
End Function

Sub SetPlayerFrame(ByVal index As Long, ByVal frameValue As Long)

    If index > Player_HighIndex Then Exit Sub
    Player(index).playerFrame = frameValue
End Sub

Sub ClearPlayerFrame(ByVal index As Long)
    If Player(index).playerFrame > 0 Then
        Player(index).playerFrame = 0
    End If
End Sub

Sub ResetProjectileAnimation(ByVal index As Long)
    animTmr(index) = 0
    animConjure(index) = 0
    widthAnim(index) = 0
    heightAnim(index) = 0
End Sub

Sub DrawProjectileAnimation(ByVal index As Long)
    Dim sRECT As RECT, Anim As Long
    Dim X As Long, Y As Long, projectileNum As Long
    Dim originalWidth As Long, originalHeight As Long
    Dim quantFrames As Byte
    
    quantFrames = 12
    
    'Private animTmr(1 To MAX_PLAYERS) As Long
    'Private animConjure(1 To MAX_PLAYERS) As Byte
    'Private widthAnim(1 To MAX_PLAYERS) As Long
    'Private heightAnim(1 To MAX_PLAYERS) As Long

    If Player(index).ProjectileCustomType = ProjectileTypeEnum.GekiDama Then
        projectileNum = Player(index).ProjectileCustomNum
        If projectileNum > 0 Then
        
            originalWidth = (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            originalHeight = (mTexture(TextureProjectile(projectileNum)).RealHeight)

            If animTmr(index) < getTime Then
                If animConjure(index) < quantFrames Then
                    animConjure(index) = animConjure(index) + 1
                Else
                    animConjure(index) = 1
                End If
                
                widthAnim(index) = widthAnim(index) + (originalWidth / quantFrames)
                heightAnim(index) = heightAnim(index) + (originalHeight / quantFrames)
                
                animTmr(index) = getTime + 80
            End If
            
            If widthAnim(index) > originalWidth Then
                widthAnim(index) = originalWidth
            End If
            If heightAnim(index) > originalHeight Then
                heightAnim(index) = originalHeight
            End If
            
            Anim = animConjure(index)
            
            X = GetPlayerX(index)
            Y = GetPlayerY(index) - 2
            
            X = X * PIC_X
            Y = Y * PIC_Y
            
            X = X - (widthAnim(index) / 2) + 16
            Y = Y - (heightAnim(index) / 2)
            
            sRECT.Top = 0
            sRECT.Bottom = sRECT.Top + (mTexture(TextureProjectile(projectileNum)).RealHeight)
            sRECT.Left = Anim * (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            sRECT.Right = sRECT.Left + (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            
            Call RenderTexture(TextureProjectile(projectileNum), ConvertMapX(X), ConvertMapY(Y), sRECT.Left, sRECT.Top, widthAnim(index), heightAnim(index), sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top)
        End If
    End If
End Sub
