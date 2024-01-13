Attribute VB_Name = "modProjectile"
Option Explicit

Public MapProjectile() As MapProjectileRec
Public EmptyMapProjectile As MapProjectileRec

Public Type XYRec
    X As Double
    Y As Double
End Type

Public Type ProjectileDataRec
    Graphic As Long
    RecuringDamage As Boolean
    Speed As Long
    Rotation As Integer
    Ammo As Long
    Duration As Long
    AnimOnHit As Long
    ProjectileOffset(1 To 4) As XYRec
    ImpactRange As Byte
    projectileType As Byte
End Type

Public Type MapProjectileRec
    Owner As Long
    OwnerType As Byte
    Graphic As Long
    
    Speed As Long
    RotateSpeed As Byte
    Rotate As Single
    Duration As Long
    ProjectileOffset(1 To 4) As XYRec
    Direction As Byte
    X As Long
    Y As Long
    xOffset As Long
    yOffset As Long
    tx As Long
    ty As Long
    spellnum As Long
    'Cliente Apenas
    IsAoE As Boolean
    IsDirectional As Boolean
    curAnim As Long
    EndTime As Long
End Type

Public Sub DrawProjectile(ByVal i As Long)
    Dim SpriteTop As Byte
    Dim Xo As Long, Yo As Long
    Dim sRECT As RECT, Anim As Long
    Dim textureWidth As Long, textureHeight As Long

    ' Check for subscript out of range RTE9
    With MapProjectile(i)
        If i < 0 Or i > MAX_PROJECTILE_MAP Then Exit Sub

        If .Graphic <= 0 Or .Graphic > CountProjectile Then
            Exit Sub
        End If

        ' Atribuição às variaveis
        Anim = .curAnim

        With sRECT
            If MapProjectile(i).IsDirectional Then
                Select Case MapProjectile(i).Direction
                Case DIR_UP
                    SpriteTop = 0
                Case DIR_DOWN
                    SpriteTop = 1
                Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                    SpriteTop = 2
                Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                    SpriteTop = 3
                End Select

                .Top = SpriteTop * (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealHeight / 4)
                .Bottom = .Top + (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealHeight / 4)
            Else
                .Top = 0
                .Bottom = .Top + (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealHeight)
            End If
            .Left = Anim * (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealWidth / 12)
            .Right = .Left + (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealWidth / 12)
        End With

        ' Atribuição às variaveis
        textureWidth = (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealWidth / 12)
        textureHeight = (mTexture(TextureProjectile(MapProjectile(i).Graphic)).RealHeight)
        Xo = .X + 16 - (textureWidth / 2)
        Yo = .Y - (textureHeight / 2)

        ' Acrécimo do offset à posição
        Select Case .Direction
            ' Up
        Case DIR_UP
            Xo = Xo + .ProjectileOffset(DIR_UP + 1).X
            Yo = Yo + .ProjectileOffset(DIR_UP + 1).Y

            ' Down
        Case DIR_DOWN
            Xo = Xo + .ProjectileOffset(DIR_DOWN + 1).X
            Yo = Yo + .ProjectileOffset(DIR_DOWN + 1).Y

            ' Left
        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            Xo = Xo + .ProjectileOffset(DIR_LEFT + 1).X
            Yo = Yo + .ProjectileOffset(DIR_LEFT + 1).Y

            ' Right
        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            Xo = Xo + .ProjectileOffset(DIR_RIGHT + 1).X
            Yo = Yo + .ProjectileOffset(DIR_RIGHT + 1).Y
        End Select

        ' Faz a rotação da imagem ou não. Obs: Tem dois tipos, a que rotaciona o gráfico e a que rotaciona o projétil
        If .Rotate = 0 Then
            Call RenderTexture(TextureProjectile(.Graphic), ConvertMapX(Xo), ConvertMapY(Yo), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top)
        Else
            Call RenderTexture(TextureProjectile(.Graphic), ConvertMapX(Xo), ConvertMapY(Yo), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, , , .Rotate)
        End If
    End With
End Sub

Public Sub ProcessProjectile(ByVal i As Long)
    Dim projectileType As Long
    Dim Angle As Long, X As Long, Y As Long, N As Long

    ' Check for subscript out of range RTE9
    With MapProjectile(i)
        If i < 0 Or i > MAX_PROJECTILE_MAP Then Exit Sub

        If .Owner <= 0 Or .Owner > Player_HighIndex Then
            Exit Sub
        End If

        ' Atribuição às variaveis
        projectileType = Spell(.spellnum).Projectile.projectileType

    End With

    ' ****** Create Particle ******
    With MapProjectile(i)
        If projectileType = ProjectileTypeEnum.KiBall Or projectileType = ProjectileTypeEnum.GenkiDama Then
            ' ****** Update Position ******
            Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tx, .ty)
            .X = .X + (Sin(Angle) * ElapsedTime * (.Speed / 1000))
            .Y = .Y - (Cos(Angle) * ElapsedTime * (.Speed / 1000))

            ' ****** Update Rotation ******
            If .RotateSpeed > 0 Then
                .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                Do While .Rotate > 360
                    .Rotate = .Rotate - 360
                Loop
            End If

            If Abs(.X - .tx) < 60 Then
                If Abs(.Y - .ty) < 60 Then
                    If .curAnim < 9 Then
                        .curAnim = 9
                    End If
                End If
            End If
        ElseIf projectileType = ProjectileTypeEnum.IsTrap Then
            If Tick + 120 >= .Duration Then
                .curAnim = 9
            End If
        End If
    End With

    ' Processamento
    With MapProjectile(i)
        If .OwnerType = TARGET_TYPE_PLAYER Then
            If projectileType = ProjectileTypeEnum.KiBall Or projectileType = ProjectileTypeEnum.GenkiDama Then

                ' Faz a limpeza do projétil caso encontre uma entity ou limite do mapa.
                If Not CheckProjectileFrontEntityOrMapLimits(i) Then
                    GoTo Clear
                End If

                ' Faz o processamento de acerto em jogadores no range, caso seja um projectile sem rescuring faz a limpeza
                If Not ProcessProjectileHasPlayerInRange(i) Then
                    GoTo Clear
                End If

                ' Faz o processamento de acerto em npcs no range, caso seja um projectile sem rescuring faz a limpeza
                If Not ProcessProjectileHasNpcInRange(i) Then
                    GoTo Clear
                End If

                ' Verifica se chegou ao fim da rota
                If Not CheckProjectileEndRoute(i) Then
                    GoTo Clear
                End If
            ElseIf projectileType = ProjectileTypeEnum.IsTrap Then

                ' Faz o processamento de jogadores em cima da trap e limpa caso não seja rescuring damage
                If Not ProcessProjectileHasPlayerInRange(i) Then
                    GoTo Clear
                End If

                ' Faz o processamento de npcs em cima da trap e limpa caso não seja rescuring damage
                If Not ProcessProjectileHasNpcInRange(i) Then
                    GoTo Clear
                End If

                ' Verifica se o tempo de duration da trap já expirou
                If Not CheckProjectileEndRoute(i) Then
                    GoTo Clear
                End If
            End If
        End If
    End With

    Exit Sub
Clear:
    Call ClearProjectile(i)
    Exit Sub
End Sub


Private Function CheckProjectileFrontEntityOrMapLimits(ByVal ProjectileIndex As Long) As Boolean

    Dim X As Long, Y As Long, rangeX As Long, rangeY As Long, minX As Long, minY As Long, maxX As Long, maxY As Long
    
    CheckProjectileFrontEntityOrMapLimits = True

    With MapProjectile(ProjectileIndex)

        ' Obter posição do projetil e converter em grid
        X = .X / PIC_X
        Y = .Y / PIC_Y

        ' Limites do mapa antingido pelo posicionamento do projétil.
        If X < 0 Or X > Map.MapData.maxX Then
            CheckProjectileFrontEntityOrMapLimits = False
            Exit Function
        End If
        If Y < 0 Or Y > Map.MapData.maxY Then
            CheckProjectileFrontEntityOrMapLimits = False
            Exit Function
        End If

        ' Obtem o range da spell, caso não seja uma AoE, define por padrão como 32x32 = 1x1
        rangeX = (Spell(.spellnum).DirectionAoE(.Direction + 1).X)
        rangeY = (Spell(.spellnum).DirectionAoE(.Direction + 1).Y)
        If Not Spell(.spellnum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Obter as posições com os ranges máximos
        maxX = X + rangeX
        maxY = Y + rangeY
        ' Obter as posições com os ranges mínimos
        minX = X - rangeX
        minY = Y - rangeY

        ' Verificações feitas do máximo para o mínimo, onde há mais possibilidade de encontrar primeiro a Entity e sair.
        For X = maxX To minX Step -1
            For Y = maxY To minY Step -1
                If X >= 0 And X <= Map.MapData.maxX Then
                    If Y >= 0 And Y <= Map.MapData.maxY Then
                        Select Case Map.TileData.Tile(X, Y).Type
                        Case TILE_TYPE_BLOCKED, TILE_TYPE_RESOURCE
                            CheckProjectileFrontEntityOrMapLimits = False
                            Exit Function
                        End Select
                    End If
                End If

            Next Y
        Next X

    End With
End Function

Private Function ProcessProjectileHasPlayerInRange(ByVal ProjectileIndex As Long) As Boolean
    Dim X As Long, Y As Long, rangeX As Long, rangeY As Long
    Dim i As Long, ImpactRange As Long, AttackerIndex As Long

    ProcessProjectileHasPlayerInRange = True

    With MapProjectile(ProjectileIndex)

        ' Caso seja um projétil com RescuringDamage, a limpeza nunca poderá ser feita por aqui, então saia da função
        If Spell(.spellnum).Projectile.RecuringDamage Then
            Exit Function
        End If

        ' Obter posição do projetil e converter em grid
        X = .X / PIC_X
        Y = .Y / PIC_Y

        rangeX = (Spell(.spellnum).DirectionAoE(.Direction + 1).X)
        rangeY = (Spell(.spellnum).DirectionAoE(.Direction + 1).Y)
        ImpactRange = Spell(.spellnum).Projectile.ImpactRange
        AttackerIndex = .Owner

        If Not Spell(.spellnum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Verificar todos os jogadores online e afunilar com o numero do mapa do atacante
        For i = 1 To Player_HighIndex
            If i <> AttackerIndex Then
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If isInRangeX(rangeX, GetPlayerX(i), X) Then
                            If isInRangeY(rangeY, GetPlayerY(i), Y) Then
                                If Not Spell(.spellnum).Projectile.RecuringDamage Then
                                    ProcessProjectileHasPlayerInRange = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i

    End With
End Function

Private Function ProcessProjectileHasNpcInRange(ByVal ProjectileIndex As Long) As Boolean
    Dim X As Long, Y As Long, rangeX As Long, rangeY As Long
    Dim N As Long, ImpactRange As Long, AttackerIndex As Long

    ProcessProjectileHasNpcInRange = True

    With MapProjectile(ProjectileIndex)
    
        ' Caso seja um projétil com RescuringDamage, a limpeza nunca poderá ser feita por aqui, então saia da função
        If Spell(.spellnum).Projectile.RecuringDamage Then
            Exit Function
        End If
        
        ' Obter posição do projetil e converter em grid
        X = .X / PIC_X
        Y = .Y / PIC_Y

        rangeX = (Spell(.spellnum).DirectionAoE(.Direction + 1).X)
        rangeY = (Spell(.spellnum).DirectionAoE(.Direction + 1).Y)
        ImpactRange = Spell(.spellnum).Projectile.ImpactRange
        AttackerIndex = .Owner

        If Not Spell(.spellnum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Verificar todos os jogadores online e afunilar com o numero do mapa do atacante
        For N = 1 To MAX_MAP_NPCS
            If MapNpc(N).Num <> 0 Then
                If isInRangeX(rangeX, MapNpc(N).X, X) Then
                    If isInRangeY(rangeY, MapNpc(N).Y, Y) Then
                        If Not Spell(.spellnum).Projectile.RecuringDamage Then
                            ProcessProjectileHasNpcInRange = False
                        End If
                    End If
                End If
            End If
        Next N
    End With
End Function

Private Function CheckProjectileEndRoute(ByVal ProjectileIndex As Long) As Boolean
    Dim TargetType As Long, target As Long

    ' Caso seja uma projectile com Rescuring Damage ativado, com o value de .duration > 0
    ' ao chegar no alvo, ela se estabiliza e começa a contagem da variavel .duration, após isso da o clear,
    ' fiz isto pra dar mais sentido ao rescuring ao ter um target como alvo.

    CheckProjectileEndRoute = True

    With MapProjectile(ProjectileIndex)
        ' Verifica se chegou ao fim da rota
        If Spell(.spellnum).Projectile.projectileType = ProjectileTypeEnum.GenkiDama Or Spell(.spellnum).Projectile.projectileType = ProjectileTypeEnum.KiBall Then
            If isInRangeX(0, (MapProjectile(ProjectileIndex).X / PIC_X), (MapProjectile(ProjectileIndex).tx / PIC_X)) Then
                If isInRangeY(0, (MapProjectile(ProjectileIndex).Y / PIC_Y), (MapProjectile(ProjectileIndex).ty / PIC_Y)) Then

                    If Spell(.spellnum).Projectile.RecuringDamage Then
                        If Spell(.spellnum).Projectile.Duration > 0 Then
                            If Tick >= .Duration Then
                                CheckProjectileEndRoute = False
                            End If
                        Else
                            CheckProjectileEndRoute = False
                        End If
                    Else
                        CheckProjectileEndRoute = False
                    End If
                    
                End If
            End If
        ElseIf Spell(.spellnum).Projectile.projectileType = ProjectileTypeEnum.IsTrap Then
            ' Verifica se o tempo de duration da trap já expirou
            If Tick >= .Duration Then
                CheckProjectileEndRoute = False
            End If
        End If
    End With
End Function

Public Sub ProcessProjectileCurAnimation(ByVal i As Long)
    If MapProjectile(i).Graphic > 0 Then
        If MapProjectile(i).curAnim = 8 Then
            MapProjectile(i).curAnim = 3
        ElseIf MapProjectile(i).curAnim < 8 Then
            MapProjectile(i).curAnim = MapProjectile(i).curAnim + 1
        ElseIf MapProjectile(i).curAnim > 8 And MapProjectile(i).curAnim < 11 Then
            MapProjectile(i).curAnim = MapProjectile(i).curAnim + 1
        End If
    End If
End Sub

Public Sub ClearProjectile(ByVal ProjectileSlot As Long)
    If MapProjectile(ProjectileSlot).OwnerType = TARGET_TYPE_PLAYER Then
        SetPlayerFrame MapProjectile(ProjectileSlot).Owner, 0
    End If
    MapProjectile(ProjectileSlot) = EmptyMapProjectile
End Sub

Public Function isInRangeX(ByVal rangeX As Long, ByVal X1 As Long, ByVal X2 As Long) As Boolean
    Dim nVal As Long

    isInRangeX = False
    nVal = Sqr((X1 - X2) ^ 2)
    If nVal <= rangeX Then isInRangeX = True: Exit Function
End Function

Public Function isInRangeY(ByVal rangeY As Long, ByVal Y1 As Long, ByVal Y2 As Long) As Boolean
    Dim nVal As Long

    isInRangeY = False
    nVal = Sqr((Y1 - Y2) ^ 2)
    If nVal <= rangeY Then isInRangeY = True: Exit Function
End Function
