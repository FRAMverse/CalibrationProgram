Module MatRateMod
    Sub MaturationRate()


        'COMPUTE MATURATION RATES IN RECOVERY YEARS
        For TStep = 1 To NumSteps
            For Age = 2 To MaxAge
                If TotMortTerm(Age, TStep) > 0 Then
                    If TStep <> NumSteps Then
                        MatRate(Stknum, Age, TStep) = TotMortTerm(Age, TStep) / (TotMortTerm(Age, TStep) + (Cohort(Age, TStep + 1) / SurvRate(Age, TStep + 1)))
                    Else
                        If Age <> MaxAge Then
                            MatRate(Stknum, Age, TStep) = TotMortTerm(Age, TStep) / (TotMortTerm(Age, TStep) + (Cohort(Age + 1, 1) / SurvRate(Age + 1, 1)))
                        Else
                            MatRate(Stknum, Age, TStep) = 1
                        End If
                    End If
                Else
                    MatRate(Stknum, Age, TStep) = 0
                End If
            Next Age
        Next TStep
    End Sub

End Module
