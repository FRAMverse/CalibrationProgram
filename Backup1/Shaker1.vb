Module Shaker1
    Sub ShakDistr()

        ''COMPUTE TOTAL RECOVERIES FOR STOCK IN EACH FISHERY AND CONCENTRATION
        'ReDim Conc(NumFish), PointFish(NumFish), StkFishCatch(NumFish), StkCheck(NumStk, NumFish)
        'For STk = MinStk To NumStk
        '    For Fish = 1 To NumFish
        '        StkCheck(STk, Fish) = 0
        '        For Age = 2 To MaxAge
        '            For TStep = 1 To NumSteps
        '                StkFishCatch(Fish) = StkFishCatch(Fish) + Catcha(STk, Age, Fish, TStep)
        '            Next TStep
        '        Next Age
        '        If TrueCatch(Fish) > 0 Then
        '            Conc(Fish) = StkFishCatch(Fish) / TrueCatch(Fish)
        '        End If
        '    Next Fish

        '    'SORT CONCENTRATIONS

        '    For Fish = 1 To NumFish
        '        PointFish(Fish) = Fish
        '    Next Fish
        '    For Fish = 1 To NumFish - 1
        '        For Check = Fish + 1 To NumFish
        '            If Conc(Check) < Conc(Fish) Then
        '                TempConc = Conc(Check)
        '                TempPoint = PointFish(Check)
        '                Conc(Check) = Conc(Fish)
        '                PointFish(Check) = PointFish(Fish)
        '                Conc(Fish) = TempConc
        '                PointFish(Fish) = TempPoint
        '            End If
        '        Next Check
        '    Next Fish

        '    'SET FLAG FOR INCLUSION OF STOCK WITHIN SHAKER COMPUTATIONS
        '    ' FIRST, LOCATE FISHERY WITH CONCENTRATION >0

        '    FirstFish = 0
        '    Fish = 0
        '    While FirstFish = 0 And Fish <= NumFish
        '        Fish = Fish + 1
        '        If Conc(Fish) > 0 Then
        '            FirstFish = Fish
        '        End If
        '    End While
        '    If FirstFish = 0 Then
        '        MsgBox("Catch in all fisheries is 0")
        '        Stop
        '    End If

        '    ' SECOND, CUTOFF LOWER 30 OF FISHERIES WITH CATCH

        '    NumNonZero = (NumFish - FirstFish + 1)
        '    LowerLim = FirstFish + 0.3 * NumNonZero

        '    ' THIRD, SET STOCK INCLUSION FLAG

        '    For Fish = LowerLim To NumFish
        '        StkCheck(STk, PointFish(Fish)) = 1
        '    Next Fish
        'Next STk

    End Sub
End Module

    