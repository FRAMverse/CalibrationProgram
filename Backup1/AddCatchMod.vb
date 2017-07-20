Module AddCatchMod
    Sub Addcatch()
        '**********************************************************************
        'Subroutine to add catch in all fisheries and compute the proportion of
        ' the catch comprised of each stock.
        '**********************************************************************

        ReDim AnnualCatch(NumFish)
        ReDim StockCatch(NumStk, NumFish)
        ReDim StockCatchProp(NumStk, NumFish)
        ReDim TimeCatch(NumFish, NumSteps)
        ReDim TotCatch(NumFish, NumSteps)
        ReDim AgeCatch(MaxAge, NumFish, NumSteps)
        Dim CWTRow As CWTData
        'COMPUTE TOTAL CATCH IN EACH FISHERY

        If Firstpass = False Then

            'COMPUTE TOTAL CATCH IN EACH FISHERY

            For TStep = 1 To NumSteps
                For Fish = 1 To NumFish - 1
                    For STk = MinStk To NumStk
                        For Age = 2 To MaxAge
                            If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep)
                                AnnualCatch(Fish) = AnnualCatch(Fish) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep)
                                StockCatch(STk, Fish) = StockCatch(STk, Fish) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep)
                                StockCatchProp(STk, Fish) = StockCatchProp(STk, Fish) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep)
                                TimeCatch(Fish, TStep) = TimeCatch(Fish, TStep) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep)
                            End If
                        Next Age
                    Next STk
                Next Fish
            Next TStep

            'COMPUTE PROPORTION OF CATCH COMPRISED OF EACH STOCK

            For Fish = 1 To NumFish - 1
                If AnnualCatch(Fish) > 0 Then
                    For STk = MinStk To NumStk
                        StockCatchProp(STk, Fish) = StockCatch(STk, Fish) / AnnualCatch(Fish)
                    Next STk
                End If
            Next Fish

        Else

            For Each CWTRow In CWTList
                If CWTRow.cLookUp = DictionaryKey Then
                    Age = CWTRow.cAge
                    Fish = CWTRow.cFish
                    TStep = CWTRow.cTStep
                    STk = CWTRow.cStk
                    TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + ExpFact * CWTRow.cCatch
                    AnnualCatch(Fish) = AnnualCatch(Fish) + ExpFact * CWTRow.cCatch
                    TimeCatch(Fish, TStep) = TimeCatch(Fish, TStep) + ExpFact * CWTRow.cCatch
                    StockCatch(STk, Fish) = StockCatch(STk, Fish) + ExpFact * CWTRow.cCatch
                    StockCatchProp(STk, Fish) = StockCatchProp(STk, Fish) + ExpFact * CWTRow.cCatch
                    AgeCatch(Age, Fish, TStep) = AgeCatch(Age, Fish, TStep) + ExpFact * CWTRow.cCatch
                    TotalStk(STk) = TotalStk(STk) + CWTRow.cCatch
                End If
            Next

            For Fish = 1 To NumFish - 1
                If AnnualCatch(Fish) > 0 Then
                    StockCatchProp(STk, Fish) = StockCatch(STk, Fish) / AnnualCatch(Fish)
                End If
            Next Fish
        End If

    End Sub
End Module
