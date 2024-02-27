Attribute VB_Name = "modAnimated"
Option Explicit

'item animated
Private StepItem As Byte
Private Itemtmr As Long

'quest objetives animated
Private StepQuestObj As Byte
Private QuestObjtmr As Long

Public Enum Animated
    AnimTextureItem = 1
    AnimTextureQuestObj
End Enum

Public Sub RenderTexture_Animated(Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal sX As Single, ByVal sY As Single, _
                                  ByVal w As Long, ByVal h As Long, ByVal sW As Single, ByVal sH As Single, ByVal AnimType As Animated, _
                                  Optional ByVal colour As Long = -1, Optional ByVal offset As Boolean = False, _
                                  Optional ByVal degrees As Single = 0, Optional ByVal Shadow As Byte = 0)

    If AnimType = AnimTextureItem Then
        If Itemtmr <= getTime Then
            If StepItem < 4 Then
                StepItem = StepItem + 1
                Itemtmr = getTime + 200
            Else
                StepItem = 0
            End If
        End If

        If StepItem = 1 Then
            Y = Y - 2
        ElseIf StepItem = 2 Then
            Y = Y - 4
        ElseIf StepItem = 3 Then
            Y = Y - 2
        End If
    End If

    If AnimType = AnimTextureQuestObj Then
        If QuestObjtmr <= getTime Then
            If StepQuestObj < 4 Then
                StepQuestObj = StepQuestObj + 1
                QuestObjtmr = getTime + 100
            Else
                StepQuestObj = 0
            End If
        End If

        If StepQuestObj = 1 Then
            Y = Y - 2
        ElseIf StepQuestObj = 2 Then
            Y = Y - 4
        ElseIf StepQuestObj = 3 Then
            Y = Y - 6
        End If
    End If

    RenderTexture Texture, X, Y, sX, sY, w, h, sW, sH, colour, offset, degrees, Shadow
End Sub

Public Function VerifyWindowsIsInCur() As Boolean
    Dim i As Integer
    For i = 1 To windowCount
        With Windows(i)
            '.Window.state = entStates.Normal
            If .Window.visible Then
                If Not .Window.clickThrough Then
                    If GlobalX >= .Window.Left And GlobalX <= .Window.Left + .Window.Width Then
                        If GlobalY >= .Window.Top And GlobalY <= .Window.Top + .Window.Height Then
                            VerifyWindowsIsInCur = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End With
    Next i
End Function


Function SecondsToHMS(ByRef Segundos As Long) As String
    Dim HR As Long, ms As Long, SS As Long, MM As Long
    Dim Total As Long, Count As Long

    If Segundos = 0 Then Exit Function

    HR = (Segundos \ 3600)
    MM = (Segundos \ 60)
    SS = Segundos
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
    ElseIf SS > 0 Then
        ' Joga na função esse segundo.
        SecondsToHMS = SS & "s "
        Total = Total - SS
    End If
End Function
