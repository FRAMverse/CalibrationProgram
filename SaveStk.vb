Module SaveStk
    Sub SaveStkCheck()
        '**********************************************************************
        'Subroutine writes flag for inclusion of stock within shaker
        ' computations to .CAL file.
        '**********************************************************************

        ReDim fishb(NumFish)
        ReDim stock(NumStk)
        'SET UP DESCRIPTORS

        For STk = MinStk To NumStk
            Stock(STk) = ",Stock " + LTrim(Str(STk)) + ";"
        Next STk
        'StkForm = "\        \"

        For Fish = 1 To NumFish
            fishb(Fish) = "Fishery " + LTrim(Str(Fish)) + ";"
        Next Fish
        'FishForm = "\           \"

        'WRITE FLAG FOR INCLUSION OF STOCK IN SHAKER COMPUTATIONS

        Print(4, "Flags for inclusion of stock in shaker computations", vbCrLf)
        For STk = MinStk To NumStk
            For Fish = 1 To NumFish
                Print(4, StkCheck(STk, Fish).ToString.PadRight(15))
                Print(4, stock(STk))
                Print(4, fishb(Fish), vbCrLf)
            Next Fish
        Next STk
    End Sub
End Module
