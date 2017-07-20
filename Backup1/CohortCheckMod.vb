Imports System.Math

Module CohortCheckMod
    Sub CohortCheck()
        'Subroutine to determine if cohort size has stabilized in successive
        ' iterations.  Maximum percent difference in cohort size is compared
        ' with the parameter ConvTol to determine if change in cohort size is
        ' acceptable.
        '**********************************************************************

        'LOCATE MAXIMUM CHANGE IN AGE 2 COHORT SIZE BETWEEN ITERATIONS
        OldDif = 0

        If Firstpass = True Then

            For STk = MinStk To MaxStk
                Dif = Math.Abs(Age2Cohort(STk) - Cohort(2, 1)) / Age2Cohort(STk)
                If Dif > OldDif Then
                    OldDif = Dif
                End If
            Next STk
        Else

            For STk = MinStk To NumStk
                Dif = Abs(Age2Cohort(STk) - CohortAll(STk, 2, PTerm, 1)) / Age2Cohort(STk)
                If Dif > OldDif Then
                    OldDif = Dif
                End If
            Next STk
        End If
        'CHECK IF MAXIMUM DIFFERENCE IS GREATER THAN ALLOWABLE TOLERANCE

        If OldDif > CT Then
            ConvFlag = "FALSE"

        Else
            ConvFlag = "TRUE"
        End If

    End Sub

End Module
