Attribute VB_Name = "Client_Frames"
Option Explicit

Private animTmr(1 To MAX_PLAYERS) As Long
Private animConjure(1 To MAX_PLAYERS) As Byte
Private widthAnim(1 To MAX_PLAYERS) As Long
Private heightAnim(1 To MAX_PLAYERS) As Long

Public Enum ProjectileTypeEnum
    None = 0
    KiBall
    GenkiDama
    IsTrap
    
    ProjectileTypeCount
End Enum

Function GetPlayerFrame(ByVal Index As Long) As Long

    If Index > Player_HighIndex Then Exit Function
    GetPlayerFrame = Player(Index).playerFrame
End Function

Sub SetPlayerFrame(ByVal Index As Long, ByVal frameValue As Long)

    If Index > Player_HighIndex Then Exit Sub
    Player(Index).playerFrame = frameValue
End Sub

Sub ClearPlayerFrame(ByVal Index As Long)
    If Player(Index).playerFrame > 0 Then
        Player(Index).playerFrame = 0
    End If
End Sub

Sub ResetProjectileAnimation(ByVal Index As Long)
    animTmr(Index) = 0
    animConjure(Index) = 0
    widthAnim(Index) = 0
    heightAnim(Index) = 0
End Sub

Sub DrawProjectileAnimation(ByVal Index As Long)
    Dim sRECT As RECT, Anim As Long
    Dim X As Long, Y As Long, projectileNum As Long
    Dim originalWidth As Long, originalHeight As Long
    Dim quantFrames As Byte
    
    quantFrames = 12

    If Player(Index).ConjureAnimProjectileType = ProjectileTypeEnum.GenkiDama Then
        projectileNum = Player(Index).ConjureAnimProjectileNum
        If projectileNum > 0 Then
        
            originalWidth = (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            originalHeight = (mTexture(TextureProjectile(projectileNum)).RealHeight)

            If animTmr(Index) < getTime Then
                If animConjure(Index) < quantFrames Then
                    animConjure(Index) = animConjure(Index) + 1
                Else
                    animConjure(Index) = 1
                End If
                
                widthAnim(Index) = widthAnim(Index) + (originalWidth / quantFrames)
                heightAnim(Index) = heightAnim(Index) + (originalHeight / quantFrames)
                
                animTmr(Index) = getTime + 80
            End If
            
            If widthAnim(Index) > originalWidth Then
                widthAnim(Index) = originalWidth
            End If
            If heightAnim(Index) > originalHeight Then
                heightAnim(Index) = originalHeight
            End If
            
            Anim = animConjure(Index)
            
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 2
            
            X = X * PIC_X
            Y = Y * PIC_Y
            
            X = X - (widthAnim(Index) / 2) + 16
            Y = Y - (heightAnim(Index) / 2)
            
            sRECT.Top = 0
            sRECT.Bottom = sRECT.Top + (mTexture(TextureProjectile(projectileNum)).RealHeight)
            sRECT.Left = Anim * (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            sRECT.Right = sRECT.Left + (mTexture(TextureProjectile(projectileNum)).RealWidth / quantFrames)
            
            Call RenderTexture(TextureProjectile(projectileNum), ConvertMapX(X), ConvertMapY(Y), sRECT.Left, sRECT.Top, widthAnim(Index), heightAnim(Index), sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top)
        End If
    End If
End Sub
