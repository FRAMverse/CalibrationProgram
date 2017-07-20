Module Shaker2
    Sub ShakDistr2()
        '**************************************************************************
        'Subroutine to compute shaker inclusion flag based on exploitation rates.
        '**************************************************************************

        ReDim PointStk(NumStk), StkCheck(NumStk, NumFish)

        'RECONSTRUCT COHORT USING CATCH RECOVERIES ONLY

        CompCohort2()

        'COMPUTE EXPLOITATION RATES FOR EACH FISHERY

        SelAge = 4 'only age 4 fish are evaluated for shaker inclusion (presumably age with highest liklihood of recoveries)
        Print(12, "Shaker Inclusion Flag - No Exploitation in Fishery", vbCrLf)
        For Fish = 1 To NumFish - 1
            ReDim u(NumStk)
            For STk = MinStk To NumStk
                Catchb = 0
                StkCheck(STk, Fish) = 0
                'COMPUTE TOTAL CATCH STOCK OF SELAGE Lumped over time steps

                For TStep = 1 To NumSteps
                    Catchb = Catchb + CWTAll(STk, SelAge, Fish, TStep)
                Next TStep

                'COMPUTE EXPLOITATION RATE

                If Cohort(STk, SelAge) > 0 Then
                    u(STk) = Catchb / (Cohort(STk, SelAge) * AnnSurvRate(SelAge))
                Else
                    u(STk) = 0
                End If
            Next STk

            'SORT EXPLOITATION RATES

            For STk = MinStk To NumStk
                PointStk(STk) = STk
            Next STk
            For STk = MinStk To NumStk - 1
                For Check = STk + 1 To NumStk
                    If u(Check) < u(STk) Then
                        Tempu = u(Check)
                        TempPoint = PointStk(Check)
                        u(Check) = u(STk)
                        PointStk(Check) = PointStk(STk)
                        u(STk) = Tempu
                        PointStk(STk) = TempPoint
                    End If
                Next Check
            Next STk

            'SET FLAG FOR INCLUSION OF STOCK WITHIN SHAKER COMPUTATIONS
            ' FIRST, LOCATE STOCKS WITH EXPLOITATION >0

            FirstStk = 0
            STk = 0
            While FirstStk = 0 And STk < NumStk
                STk = STk + 1
                If u(STk) > 0 Then
                    FirstStk = STk
                End If
            End While
            If FirstStk = 0 Then
                Print(12, "None of the model stocks are exploited in fishery", vbCrLf)
                Print(12, "Fishery =", Fish, vbCrLf)
                ' Print()
            Else

                ' SECOND, CUTOFF LOWER 30 OF STOCKS WITH U > 0

                NumNonZero = (NumStk - FirstStk + 1)
                LowerLim = FirstStk + 0.3 * NumNonZero

                ' THIRD, SET STOCK INCLUSION FLAG

                For STk = LowerLim To NumStk
                    StkCheck(PointStk(STk), Fish) = 1
                    Print(12, STk & "," & Fish & "," & StkCheck(PointStk(STk), Fish) & vbCrLf)
                Next STk
            End If

        Next Fish

    End Sub

End Module
