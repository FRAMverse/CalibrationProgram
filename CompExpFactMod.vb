Module CompExpFactMod
    Sub CompExpFact()
        '**********************************************************************
        'Subroutine computes the expansion factor for CWT recoveries by
        ' dividing the total (tagged + untagged) observed escapement by the
        ' CWT escapement.
        'Subroutine also computes the total expanded catch in each fishery.
        '**********************************************************************

        'COMPUTE AGE 3+ CWT ESCAPEMENT FOR EACH STOCK AND EXPANSION FACTOR

        'need to read ObsEscapement
        ReDim EscExpFact(NumStk)
        For STk = MinStk To NumStk
            If STk = 23 Then
                STk = 23
            End If

            CWTEscpmnt = 0
            For Age = 3 To MaxAge
                'For TStep = 1 To NumSteps
                '    CWTEscpmnt = CWTEscpmnt + CWTAll(STk, Age, 74, TStep)
                'Next TStep
                ' compute Escapement as Esc + FWNet + FWSport (same as ETRS)
                For TStep = 1 To NumSteps

                    CWTEscpmnt = CWTEscpmnt + CWTAll(STk, Age, NumFish, TStep)

                Next TStep

            Next Age
            'observed escapement includes freshwater net and sport as well as ETRS fisheries (11A, 12H etc.)
            If CWTEscpmnt = 0 Or ObsEscpmnt(STk) < 0 Then
                EscExpFact(STk) = 1
            Else
                'EscExpFact(STk) = 1
                EscExpFact(STk) = ObsEscpmnt(STk) / CWTEscpmnt
                Debug.Print(EscExpFact(STk))
            End If
        Next

    End Sub
End Module
