Attribute VB_Name = "modProjectile"
Option Explicit

Public MapProjectile() As MapProjectileRec
Public EmptyMapProjectile As MapProjectileRec

Public Type XYRec
    x As Double
    y As Double
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
    direction As Byte
    x As Long
    y As Long
    xOffset As Long
    yOffset As Long
    tX As Long
    tY As Long
    spellNum As Long
    
    ' Servidor apenas
    Range As Byte
    Damage As Long
    AnimOnHit As Long
    mapnum As Long
    
    AttackTimer(1 To MAX_MAP_NPCS) As Long
    AttackTimerPlayer(1 To MAX_PLAYERS) As Long
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

Private Function CheckProjectileFrontEntityOrMapLimits(ByVal ProjectileIndex As Long) As Boolean

    Dim x As Long, y As Long, rangeX As Long, rangeY As Long, minX As Long, minY As Long, maxX As Long, maxY As Long
    
    CheckProjectileFrontEntityOrMapLimits = True

    With MapProjectile(ProjectileIndex)

        ' Obter posição do projetil e converter em grid
        x = .x / PIC_X
        y = .y / PIC_Y

        ' Limites do mapa antingido pelo posicionamento do projétil.
        If x < 0 Or x > Map(.mapnum).MapData.maxX Then
            CheckProjectileFrontEntityOrMapLimits = False
            Exit Function
        End If
        If y < 0 Or y > Map(.mapnum).MapData.maxY Then
            CheckProjectileFrontEntityOrMapLimits = False
            Exit Function
        End If

        ' Obtem o range da spell, caso não seja uma AoE, define por padrão como 32x32 = 1x1
        rangeX = (Spell(.spellNum).DirectionAoE(.direction + 1).x)
        rangeY = (Spell(.spellNum).DirectionAoE(.direction + 1).y)
        If Not Spell(.spellNum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Obter as posições com os ranges máximos
        maxX = x + rangeX
        maxY = y + rangeY
        ' Obter as posições com os ranges mínimos
        minX = x - rangeX
        minY = y - rangeY

        ' Verificações feitas do máximo para o mínimo, onde há mais possibilidade de encontrar primeiro a Entity e sair.
        For x = maxX To minX Step -1
            For y = maxY To minY Step -1
                If x >= 0 And x <= Map(.mapnum).MapData.maxX Then
                    If y >= 0 And y <= Map(.mapnum).MapData.maxY Then
                        If Map(.mapnum).TileData.Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map(.mapnum).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                            
                            ' Fazer o envio da animação de hit
                            If .AnimOnHit > 0 And .AnimOnHit <= MAX_ANIMATIONS Then
                                Call SendAnimation(.mapnum, .AnimOnHit, x, y)
                            End If
                            
                            CheckProjectileFrontEntityOrMapLimits = False
                            Exit Function
                        End If
                    End If
                End If
            Next y
        Next x

    End With
End Function

Private Function ProcessProjectileHasPlayerInRange(ByVal Damage As Long, ByVal ProjectileIndex As Long) As Boolean
    Dim x As Long, y As Long, rangeX As Long, rangeY As Long
    Dim i As Long, ImpactRange As Long, AttackerIndex As Long, mapnum As Long

    ProcessProjectileHasPlayerInRange = True

    With MapProjectile(ProjectileIndex)
        ' Obter posição do projetil e converter em grid
        x = .x / PIC_X
        y = .y / PIC_Y

        rangeX = (Spell(.spellNum).DirectionAoE(.direction + 1).x)
        rangeY = (Spell(.spellNum).DirectionAoE(.direction + 1).y)
        ImpactRange = Spell(.spellNum).Projectile.ImpactRange
        AttackerIndex = .Owner
        mapnum = .mapnum

        If Not Spell(.spellNum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Verificar todos os jogadores online e afunilar com o numero do mapa do atacante
        For i = 1 To Player_HighIndex
            If i <> AttackerIndex Then
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = mapnum Then
                        If isInRangeX(rangeX, GetPlayerX(i), x) Then
                            If isInRangeY(rangeY, GetPlayerY(i), y) Then
                                If CanPlayerAttackPlayer(AttackerIndex, i, True) Then
                                    If Spell(.spellNum).Projectile.RecuringDamage Then
                                        If tick > MapProjectile(ProjectileIndex).AttackTimerPlayer(i) Then
                                            ' Fazer o envio da animação de hit
                                            If .AnimOnHit > 0 And .AnimOnHit <= MAX_ANIMATIONS Then
                                                Call SendAnimation(.mapnum, .AnimOnHit, x, y)
                                            End If

                                            ' Causar o impacto
                                            If ImpactRange > 0 Then
                                                Call MakeImpact(i, ImpactRange, TARGET_TYPE_PLAYER, mapnum, AttackerIndex, False)
                                            End If

                                            PlayerAttackPlayer AttackerIndex, i, Damage, .spellNum
                                            MapProjectile(ProjectileIndex).AttackTimerPlayer(i) = tick + MapProjectile(ProjectileIndex).Speed
                                        End If
                                    Else
                                        ' Fazer o envio da animação de hit
                                        If .AnimOnHit > 0 And .AnimOnHit <= MAX_ANIMATIONS Then
                                            Call SendAnimation(.mapnum, .AnimOnHit, x, y)
                                        End If

                                        ' Causar o impacto
                                        If ImpactRange > 0 Then
                                            Call MakeImpact(i, ImpactRange, TARGET_TYPE_PLAYER, mapnum, AttackerIndex, False)
                                        End If

                                        PlayerAttackPlayer AttackerIndex, i, Damage, .spellNum
                                        ProcessProjectileHasPlayerInRange = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i

    End With
End Function

Private Function ProcessProjectileHasNpcInRange(ByVal Damage As Long, ByVal ProjectileIndex As Long) As Boolean
    Dim x As Long, y As Long, rangeX As Long, rangeY As Long
    Dim n As Long, ImpactRange As Long, AttackerIndex As Long, mapnum As Long

    ProcessProjectileHasNpcInRange = True

    With MapProjectile(ProjectileIndex)
        ' Obter posição do projetil e converter em grid
        x = .x / PIC_X
        y = .y / PIC_Y

        rangeX = (Spell(.spellNum).DirectionAoE(.direction + 1).x)
        rangeY = (Spell(.spellNum).DirectionAoE(.direction + 1).y)
        ImpactRange = Spell(.spellNum).Projectile.ImpactRange
        AttackerIndex = .Owner
        mapnum = .mapnum

        If Not Spell(.spellNum).IsAoE Then
            rangeX = 0
            rangeY = 0
        End If

        ' Verificar todos os jogadores online e afunilar com o numero do mapa do atacante
        For n = 1 To MAX_MAP_NPCS
            If MapNpc(mapnum).Npc(n).Num <> 0 Then
                If isInRangeX(rangeX, x, MapNpc(mapnum).Npc(n).x) Then
                    If isInRangeY(rangeY, y, MapNpc(mapnum).Npc(n).y) Then
                        If CanPlayerAttackNpc(AttackerIndex, n, True) Then
                            If Spell(.spellNum).Projectile.RecuringDamage Then
                                If tick > .AttackTimer(n) Then
                                    ' Fazer o envio da animação de hit
                                    If .AnimOnHit > 0 And .AnimOnHit <= MAX_ANIMATIONS Then
                                        Call SendAnimation(.mapnum, .AnimOnHit, x, y)
                                    End If

                                    ' Causar o impacto
                                    If ImpactRange > 0 Then
                                        Call MakeImpact(n, ImpactRange, TARGET_TYPE_NPC, mapnum, AttackerIndex, False)
                                    End If
                                    
                                    PlayerAttackNpc AttackerIndex, n, Damage, .spellNum
                                    .AttackTimer(n) = tick + .Speed
                                End If
                            Else
                                ' Fazer o envio da animação de hit
                                If .AnimOnHit > 0 And .AnimOnHit <= MAX_ANIMATIONS Then
                                    Call SendAnimation(.mapnum, .AnimOnHit, x, y)
                                End If

                                ' Causar o impacto
                                If ImpactRange > 0 Then
                                    Call MakeImpact(n, ImpactRange, TARGET_TYPE_NPC, mapnum, AttackerIndex, False)
                                End If

                                PlayerAttackNpc AttackerIndex, n, Damage, .spellNum
                                ProcessProjectileHasNpcInRange = False
                            End If
                        End If
                    End If
                End If
            End If
        Next n
    End With
End Function

Private Function CheckProjectileEndRoute(ByVal ProjectileIndex As Long) As Boolean
    Dim TargetType As Long, Target As Long

    ' Caso seja uma projectile com Rescuring Damage ativado, com o value de .duration > 0
    ' ao chegar no alvo, ela se estabiliza e começa a contagem da variavel .duration, após isso da o clear,
    ' fiz isto pra dar mais sentido ao rescuring ao ter um target como alvo.

    CheckProjectileEndRoute = True

    With MapProjectile(ProjectileIndex)
        ' Verifica se chegou ao fim da rota
        If Spell(.spellNum).Projectile.projectileType = ProjectileTypeEnum.GenkiDama Or Spell(.spellNum).Projectile.projectileType = ProjectileTypeEnum.KiBall Then
            If isInRangeX(0, (MapProjectile(ProjectileIndex).x / PIC_X), (MapProjectile(ProjectileIndex).tX / PIC_X)) Then
                If isInRangeY(0, (MapProjectile(ProjectileIndex).y / PIC_Y), (MapProjectile(ProjectileIndex).tY / PIC_Y)) Then

                    If Spell(.spellNum).Projectile.RecuringDamage Then
                        If Spell(.spellNum).Projectile.Duration > 0 Then
                            If tick >= .Duration Then
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
        ElseIf Spell(.spellNum).Projectile.projectileType = ProjectileTypeEnum.IsTrap Then
            ' Verifica se o tempo de duration da trap já expirou
            If tick >= .Duration Then
                CheckProjectileEndRoute = False
            End If
        End If
    End With
End Function

Public Sub CheckProjectile(ByVal i As Long)
    Dim Angle As Long
    Dim Damage As Long
    Dim projectileType As Long

    ' Verificações necessárias para evitar subscript out of range
    With MapProjectile(i)
        If i < 0 Or i > MAX_PROJECTILE_MAP Then
            Exit Sub
        End If
        If .mapnum < 0 Or .mapnum > MAX_MAPS Then
            GoTo Clear
        End If

        If .OwnerType = TARGET_TYPE_NPC Then
            If .Owner <= 0 Or .Owner > MAX_MAP_NPCS Then
                GoTo Clear
            ElseIf MapNpc(.mapnum).Npc(.Owner).Num = 0 Then
                GoTo Clear
            End If
        ElseIf .OwnerType = TARGET_TYPE_PLAYER Then
            If .Owner <= 0 Or .Owner > Player_HighIndex Then
                GoTo Clear
            ElseIf Not IsPlaying(.Owner) Then
                GoTo Clear
            End If
        Else
            GoTo Clear
        End If

        If .spellNum < 0 Or .spellNum > MAX_SPELLS Then
            GoTo Clear
        End If

        If Spell(.spellNum).Projectile.projectileType <= ProjectileTypeEnum.None Or Spell(.spellNum).Projectile.projectileType >= ProjectileTypeEnum.ProjectileTypeCount Then
            GoTo Clear
        End If
    End With

    With MapProjectile(i)
        ' Atribuição às variáveis...
        Damage = Spell(.spellNum).Vital + Int(GetPlayerStat(.Owner, Intelligence) / 3)
        projectileType = Spell(.spellNum).Projectile.projectileType

        ' Movimentação do projétil baseada no targetX(tX) e targetY(tY) que é definido na criação da projectile
        If .Graphic > 0 Then
            If projectileType = ProjectileTypeEnum.KiBall Or projectileType = ProjectileTypeEnum.GenkiDama Then
                On Error GoTo Clear

                ' ****** Update Position ******
                Angle = DegreeToRadian * Engine_GetAngle(.x, .y, .tX, .tY)
                .x = .x + (Sin(Angle) * ElapsedTime * (.Speed / 1000))
                .y = .y - (Cos(Angle) * ElapsedTime * (.Speed / 1000))
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
                If Not ProcessProjectileHasPlayerInRange(Damage, i) Then
                    GoTo Clear
                End If

                ' Faz o processamento de acerto em npcs no range, caso seja um projectile sem rescuring faz a limpeza
                If Not ProcessProjectileHasNpcInRange(Damage, i) Then
                    GoTo Clear
                End If
                
                ' Verifica se chegou ao fim da rota
                If Not CheckProjectileEndRoute(i) Then
                    GoTo Clear
                End If
            ElseIf projectileType = ProjectileTypeEnum.IsTrap Then
            
                ' Faz o processamento de jogadores em cima da trap e limpa caso não seja rescuring damage
                If Not ProcessProjectileHasPlayerInRange(Damage, i) Then
                    GoTo Clear
                End If
                
                ' Faz o processamento de npcs em cima da trap e limpa caso não seja rescuring damage
                If Not ProcessProjectileHasNpcInRange(Damage, i) Then
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

Public Function isInRangeX(ByVal rangeX As Long, ByVal x1 As Long, ByVal x2 As Long) As Boolean
    Dim nVal As Long

    isInRangeX = False
    nVal = Sqr((x1 - x2) ^ 2)
    If nVal <= rangeX Then isInRangeX = True: Exit Function
End Function

Public Function isInRangeY(ByVal rangeY As Long, ByVal y1 As Long, ByVal y2 As Long) As Boolean
    Dim nVal As Long

    isInRangeY = False
    nVal = Sqr((y1 - y2) ^ 2)
    If nVal <= rangeY Then isInRangeY = True: Exit Function
End Function

