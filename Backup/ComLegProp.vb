Imports System.Math
Module ComLegProp
    Sub CompLegProp()
        '**************************************************************************
        'Subroutine to compute legal proportion for a specific stock, age, maturity
        ' and fishery.
        '**************************************************************************
        '
        Time1 = (Age - 1) * 12 + PntTime(TStep) 'used to be age-1 but Pete changed growth function parameters
        Mean = L(STk, TermStat) * (1 - EXP(-K(STk, TermStat) * (Time1 - T0(STk, TermStat))))
        SD = CV(STk, Age, TermStat) * Mean
        If (MinSize(FishYear, FishNum, TStep) < Mean - 3 * SD) Then
            LegalProp = 1
        End If
        If (MinSize(FishYear, FishNum, TStep) > Mean + 3 * SD) Then
            LegalProp = 0
        End If
        If ((MinSize(FishYear, FishNum, TStep) >= Mean - 3 * SD) And (MinSize(FishYear, FishNum, TStep) <= Mean + 3 * SD)) Then
            LegalProp = (1 - NormlDistr(MinSize(FishYear, FishNum, TStep), Mean, SD))
        End If
        'a = MinSize(FishYear, Fish, TStep)
        'COMPUTE SUBLEGAL PROPORTION (ADJUST AGES <= MaxAgeEncAdj
        ' FOR NONRECRUITED PROPORTION AND DISTRIBUTION).

        If Age <= MaxAgeEncAdj Then
            SubLegalProp = EncRateAdj(Age, FishNum, TStep) * (1 - LegalProp)
        Else
            SubLegalProp = (1 - LegalProp)
        End If




    End Sub

End Module
