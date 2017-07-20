Module ExpRateMod
    Sub ExploitationRate()
        ''COMPUTE EXPLOITATION RATES IN RECOVERY YEARS
        BYnew = BYNum - 1980
        STk = Stknum
        BY = BYNum
        Stage = StageNum

        For Age = 2 To MaxAge
            FishYear = BY + Age
            For Fish = 1 To NumFish - 1
                For TStep = 1 To NumSteps
                    If Fish = 15 And STk = 15 And BY = 1995 And TStep = 3 Then
                        TStep = 3
                    End If

                    If AgeCatch(Age, Fish, TStep) <> 0 Then
                        FishNum = Fish
                       
                        Call CompLegProp()
                        If LegalProp > 0 Then
                            If TermFlag(Fish, TStep) = PTerm Then
                                ExRate(BYnew, Stknum, StageNum, Age, Fish, TStep) = AgeCatch(Age, Fish, TStep) / (Cohort(Age, TStep) * LegalProp)

                                mynumber = ExRate(BYnew, Stknum, StageNum, Age, Fish, TStep)
                            Else
                                ExRate(BYnew, Stknum, StageNum, Age, Fish, TStep) = AgeCatch(Age, Fish, TStep) / (TotMortTerm(Age, TStep) * LegalProp)
                            End If

                            If ExRate(BYnew, Stknum, StageNum, Age, Fish, TStep) <> 0 Then
                                ' Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & Fish & "," & TStep & "," & ExRate(BYnew, Stknum, StageNum, Age, Fish, TStep) & vbCrLf)
                            End If

                        End If
                    End If
                Next
            Next
        Next

    
    End Sub
End Module
