Module IncMortMod
    Sub IncMort()
        'CALL SHAKER AND CNR SUBROUTINES
        ReDim Shaker(MaxAge, NumFish, NumSteps)
        'ReDim ShakerAll(NumStk, MaxAge, NumFish, NumSteps)
        ReDim CNR(NumStk, MaxAge, NumFish, NumSteps)
        ReDim CNRLeg(NumStk, MaxAge, NumFish, NumSteps)
        ReDim CNRSub(NumStk, MaxAge, NumFish, NumSteps)
        ReDim PropLegCatch(NumStk, MaxAge, 1)
        ReDim PropSubCatch(NumStk, MaxAge, 1)
        ReDim TotalCNR(NumFish, NumSteps)

        Dim Shakerfilenew As String

        Shakerfilenew = filepath & "\Shakerfilenew.txt"
        FileOpen(22, Shakerfilenew, OpenMode.Output)
        Print(22, "stock,  age, fish, tstep, shaker" & vbCrLf)

        For TStep = 1 To NumSteps
            For Fish = 1 To NumFish
                FishNum = Fish
                TermStat = 0
                If TermFlag(Fish, TStep) = TermStat Then
                    ReDim PropSubPop(NumStk, MaxAge)
                    CompShakers()
                    
                    If CNRFlag(BaseType, Yr, Fish, TStep) <> 0 Then

                        CompCNR()

                    End If
                End If
            Next Fish
        Next TStep
        FileClose(22)
    End Sub
End Module
