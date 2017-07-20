Module IncMortMod
    Sub IncMort()
        'CALL SHAKER AND CNR SUBROUTINES
        ReDim Shaker(MaxAge, NumFish, NumSteps)
        'ReDim ShakerAll(NumStk, MaxAge, NumFish, NumSteps)
        ReDim CNR(MaxAge, NumFish, NumSteps)
        Dim Shakerfilenew As String

        Shakerfilenew = "C:\data\calibration\07Qbasic\Shakerfilenew"
        FileOpen(22, Shakerfilenew, OpenMode.Output)
        Print(22, "stock,  age, fish, tstep, shaker" & vbCrLf)

        For TStep = 1 To NumSteps
            For Fish = 1 To NumFish
                FishNum = Fish
                TermStat = 0
                If TermFlag(Fish, TStep) = TermStat Then
                    ReDim PropSubPop(NumStk, MaxAge)
                    CompShakers()
                    'solve issues with CNR calculations
                    'If CNRFlag(Fish, TStep) <> 0 Then
                    '    CompCNR()
                    'End If

                End If
            Next Fish
        Next TStep
        FileClose(22)
    End Sub
End Module
