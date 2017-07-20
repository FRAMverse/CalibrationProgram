Module Impute2Mod
    Sub Impute2()
        'Just a read in and print out of the exp rates
        Firstpass = False
        coma = ","
        Input(1, NumImputeER)
        Input(1, Junk)
        ' the *** in the line below is a flag to the CHCAL input routines
        ' to read the new stuff
        'PRINT #5, "*** below are imputed exploitation rates"
        'PRINT #5, NumImputeER, coma$, Junk$
        If NumImputeER > 0 Then
            For Iread = 1 To CInt(NumImputeER)
                Input(1, FisheryNum)
                Input(1, junk)
                Print(5, FisheryNum, coma, junk, vbCrLf)
                Input(1, NumOfStks)
                Input(1, junk)
                Print(5, NumOfStks, coma, junk, vbCrLf)
                For Ireadstk = 1 To NumOfStks
                    aline = LineInput(1)
                    Print(5, aline, vbCrLf)
                Next Ireadstk
            Next Iread
        End If

    End Sub
End Module
