Module SaveCohortMod
    Sub SaveCohort()
        '**********************************************************************
        'Subroutine saves age 2 cohort size of each stock for convergence check
        ' at end of next iteration.
        '**********************************************************************

        'SAVE AGE 2 COHORT SIZE FOR COMPARISON WITH NEXT ITERATION
        If Firstpass = True Then
            For STk = MinStk To MaxStk
                Age2Cohort(STk) = Cohort(2, 1)
            Next STk
        Else
            For STk = MinStk To NumStk
                Age2Cohort(STk) = CohortAll(STk, 2, PTerm, 1)
            Next STk
        End If
    End Sub

End Module
