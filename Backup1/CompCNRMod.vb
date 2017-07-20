Module CompCNRMod
    Sub CompCNR()
        '**********************************************************************
        'Subroutine computes CNR mortality in fisheries using one of 4 methods
        ' depending upon value of CNRFlag parameter:
        ' CNRFlag = 1 - Computed CNR; CNRInput1: not used, CNRInput2: not used, CNRInput3: SubSelectivity, CNRInput4: LegalSelectivity
        ' CNRFlag = 2 - Ratio of CNR days to normal days; CNRInput1: CNRDays, CNRInput2: CatchDays, CNRInput3: SubSelectivity, CNRInput4: LegalSelectivity
        ' CNRFlag = 3 - External estimate of legal and sublegal encounters during CNR fishery; CNRInput1: Legal Encounters, CNRInput2: Sublelgal Encounters
        ' CNRFlag = 4 - Total Encounters; CNRInput1: Total Encounters
        '**********************************************************************
        ' Only flag 2 or 3 can be used for the first pass (OOB), because program has no info about how to allocate total encounters 
        ' (all stocks) to the individual stock
        ReDim PropLegCatch(NumStk, MaxAge, 1)

        'CHECK TO DETERMINE IF CATCH OCCURRED IN FISHERY

        If TotCatch(Fish, TStep) = 0 Then
            MsgBox("CNR cannot be estimated for fishery with no catch")
            Stop
        End If

        ''TWO POSSIBLE METHODS OF COMPUTING CNR CATCH

        'If OOBStatus = 3 Or Firstpass = False Then
        '    Select Case CNRFlag(Fish, TStep)

        '        Case 2
        '            '...Ratio of CNR days to normal days. 
        '            ' only works for strata with CWT recoveries; Catch > 0
        '            For STk = MinStk To NumStk
        '                For Age = 2 To MaxAge
        '                    For Each CWTRow In CWTList
        '                        If CWTRow.cLookUp = DictionaryKey Then
        '                            If CWTRow.cFish = Fish And CWTRow.cTStep = TStep And CWTRow.cStk = STk And CWTRow.cAge = Age Then
        '                                CNR(Age, Fish, TStep) = ExpFact * CWTRow.cCatch * (CNRInput(1, Fish, TStep) / CNRInput(2, Fish, TStep)) * ShakMortRate(Fish, TStep) * CNRInput(4, Fish, TStep)
        '                                CNR(Age, Fish, TStep) = CNR(Age, Fish, TStep) + Shaker(Age, Fish, TStep) * (CNRInput(1, Fish, TStep) / CNRInput(2, Fish, TStep)) * CNRInput(3, Fish, TStep)
        '                            End If
        '                        End If
        '                    Next
        '                Next Age
        '            Next STk


        '        Case 3
        '            '...External estimate of encounters. Be careful not to exclude computations for stock,age,fish,tstep combos not in CWTList
        '            ' in OOB run this code splits the total legal and sublegal encounter input into ages (of that stock)
        '            'doesn't make sense as CNR input is by fishery, not fishery and stock
        '            For STk = MinStk To NumStk
        '                For Age = 2 To MaxAge

        '                    If CWTRow.cFish = Fish And CWTRow.cTStep = TStep And CWTRow.cStk = STk And CWTRow.cAge = Age Then
        '                        'If PointRAM(STk, Age, Fish, TStep) > 0 Then
        '                        '    FileGet(4, StkMortRec, PointRAM(STk, Age, Fish, TStep))
        '                        'Else
        '                        '    RamRecord = RamRecord + 1
        '                        '    PointRAM(STk, Age, Fish, TStep) = RamRecord
        '                        '    StkMortRec.MyCatch = 0
        '                        '    StkMortRec.Shaker = 0
        '                        'End If
        '                        LegalProp = ExpFact * CWTRow.cCatch / TotCatch(Fish, TStep)
        '                        SubLegProp = PropSubPop(STk, Age)
        '                        CNR(Age, Fish, TStep) = LegalProp * CNRInput(1, Fish, TStep) * ShakMortRate(Fish, TStep) 'shouldn't this be legal release mort rate
        '                        CNR(Age, Fish, TStep) = CNR(Age, Fish, TStep) + SubLegProp * CNRInput(2, Fish, TStep) * ShakMortRate(Fish, TStep)

        '                    End If

        '                Next Age
        '            Next STk
        '    End Select
        'Else


        'End If
    End Sub
End Module
