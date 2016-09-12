Attribute VB_Name = "modFOOD"
Option Explicit

Public Type tPosAndVel
    POS        As geoVector2D
    Vel        As geoVector2D
End Type

Public FOOD()  As tPosAndVel
Public FOODcolor() As Long


Public NFood   As Long

Public Const FoodSize As Double = 9


Public Sub InitFOOD(HowMuch As Long)
    Dim I      As Long

    NFood = HowMuch

    ReDim FOOD(NFood)
    For I = 0 To NFood

        With FOOD(I)
            .POS.x = wMinX + Rnd * (wMaxX - wMinX)
            .POS.y = wMinY + Rnd * (wMaxY - wMinY)
        End With
    Next

End Sub
Private Sub RemoveFood(wF As Long)
    Dim I      As Long

    NFood = NFood - 1
    For I = wF To NFood
        FOOD(I) = FOOD(I + 1)
    Next

End Sub
Private Sub FoodToRNDPosition(wF As Long)
    With FOOD(wF)
        .POS.x = wMinX + Rnd * (wMaxX - wMinX)
        .POS.y = wMinY + Rnd * (wMaxY - wMinY)
    End With
End Sub
Public Sub DrawFOOD()
    Dim I      As Long

    ' vbDRAW.CC.SetSourceColor vbGreen

    For I = 0 To NFood
        With FOOD(I)
            'vbDRAW.CC.Ellipse .Pos.X, .Pos.y, 5, 5
            'vbDRAW.CC.Fill

            If InsideBB(CameraBB, FOOD(I).POS) Then
                vbDRAW.CC.RenderSurfaceContent "FoodIcon", .POS.x - FoodSize, .POS.y - FoodSize, , , CAIRO_FILTER_FAST, 0.75
            End If
        End With
    Next


End Sub

Public Sub FoodMoveAndCheckEaten()
    Dim I      As Long
    Dim J      As Double
    Dim HeadPosition As geoVector2D
    Dim Hvel   As geoVector2D

    Dim D      As Double
    Dim dx     As Double
    Dim dy     As Double
    Dim vD     As Double
    Dim vDx    As Double
    Dim vDy    As Double


    Dim GrabR  As Double


    For I = 0 To NFood
        With FOOD(I)

            For J = 0 To NSnakes
                If InsideBB(Snake(J).getBB, FOOD(I).POS) Then    'when there's a lot of food skip check far away
                    HeadPosition = Snake(J).GetHEADPos

                    'HeadPosition = VectorSUM(HeadPosition, VectorMUL(Snake(J).GetHEADVel, Snake(J).MySIZE * 3))

                    dx = HeadPosition.x - .POS.x
                    dy = HeadPosition.y - .POS.y
                    D = dx * dx + dy * dy

                    GrabR = Snake(J).MySIZE * 10
                    GrabR = GrabR * GrabR
                    If D < GrabR Then
                        vD = 0.01 * Sqr(GrabR) / (Sqr(D) * 1)
                        vDx = dx * vD
                        vDy = dy * vD
                        .Vel.x = .Vel.x + vDx
                        .Vel.y = .Vel.y + vDy
                        Snake(J).TongueOut = 1
                    End If


                    'GrabR = Snake(J).DIAM * 0.7
                    GrabR = Snake(J).DIAM * 0.5 + FoodSize * 0.5
                    GrabR = GrabR * GrabR
                    If D < GrabR Then
                        If J = PLAYER Then
                            If Snake(PLAYER).IsDying = 0 Then
                                ' MultipleSounds.playsound "eatfruit.wav"

                                MultipleSounds.PlaySound SoundPlayerChomp
                            End If
                        Else
                            HeadPosition = Snake(0).GetHEADPos
                            dx = HeadPosition.x - .POS.x
                            dy = HeadPosition.y - .POS.y
                            D = Sqr(dx * dx + dy * dy)
                            MultipleSounds.PlaySound SoundEnemyChomp, ClampLong(-dx * 3, -10000, 10000), ClampLong(-D * 0.8, -10000, 0)
                        End If

                        Snake(J).fLength = Snake(J).fLength + 1
                        '   FoodToRNDPosition I
                        RemoveFood I
                    End If
                End If

            Next


            .POS = VectorSUM(.POS, .Vel)
            .Vel = VectorMUL(.Vel, 0.992)
            If .POS.x < wMinX Then .POS.x = wMinX: .Vel.x = -.Vel.x
            If .POS.y < wMinY Then .POS.y = wMinY: .Vel.y = -.Vel.y
            If .POS.x > wMaxX Then .POS.x = wMaxX: .Vel.x = -.Vel.x
            If .POS.y > wMaxY Then .POS.y = wMaxY: .Vel.y = -.Vel.y


        End With
    Next


End Sub

Public Sub CreateFoodFromDeadSnake(wS As Long)
    Dim I      As Long
    For I = 0 To Snake(wS).Ntokens - 2    '1
        NFood = NFood + 1
        ReDim Preserve FOOD(NFood)
        With FOOD(NFood)
            .POS = Snake(wS).GetTokenPos(I)
            .Vel.x = (Rnd * 2 - 1) * 0.25
            .Vel.y = (Rnd * 2 - 1) * 0.25
        End With
    Next
End Sub


Public Function PointToNearestFood(HeadPos As geoVector2D) As geoVector2D
    Dim I      As Long
    Dim J      As Long
    Dim D      As Double
    Dim dx     As Double
    Dim dy     As Double
    Dim MIND   As Double

    MIND = 1E+32

    For I = 0 To NFood
        dx = FOOD(I).POS.x - HeadPos.x
        dy = FOOD(I).POS.y - HeadPos.y
        D = dx * dx + dy * dy
        If D < MIND Then
            MIND = D
            J = I
        End If
    Next

    PointToNearestFood = VectorNormalize(VectorSUB(FOOD(J).POS, HeadPos))

End Function
