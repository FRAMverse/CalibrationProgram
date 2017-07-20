Module CompCohort2Mod
    Sub CompCohort2()
        '**********************************************************************
        'Subroutine to reconstruct cohort from CWT recoveries.  Cohort
        ' reconstruction proceeds backwards in time beginning with the oldest
        ' age class and last time step.  Note that for this analysis, the
        ' is simply constructed on an annual basis with no differentiation of
        ' terminal and preterminal components.  This cohort analyis is used
        ' only if the shaker distribution flag is set to 2 (distribution based
        ' on exploitation rates).
        '**********************************************************************

        ReDim Cohort(NumStk, MaxAge)
        ReDim AnnSurvRate(MaxAge)

        'COMPUTE SURVIVAL RATE ON AN ANNUAL BASIS

        For Age = 2 To MaxAge
            AnnSurvRate(Age) = 1
            For TStep = NumSteps To 1 Step -1
                AnnSurvRate(Age) = AnnSurvRate(Age) * SurvRate(Age, TStep)
            Next TStep
        Next Age

        'RECONSTRUCT COHORT ON AN ANNUAL BASIS

        For STk = MinStk To NumStk
            For Age = MaxAge To 2 Step -1

                'SUM RECOVERIES OF STOCK OVER ALL TIME STEPS

                Recoveries = 0
                For TStep = NumSteps To 1 Step -1
                    For Fish = 1 To NumFish
                        Recoveries = Recoveries + CWTAll(STk, Age, Fish, TStep)
                    Next Fish
                Next TStep

                'COHORT ANALYSIS FOR PRETERMINAL COHORT

                If Age = MaxAge Then
                    Cohort(STk, Age) = Recoveries / AnnSurvRate(Age)
                Else
                    Cohort(STk, Age) = (Cohort(STk, Age + 1) + Recoveries) / AnnSurvRate(Age)
                End If
            Next Age
        Next STk
    End Sub
End Module
