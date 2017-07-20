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
        

        'CHECK TO DETERMINE IF CATCH OCCURRED IN FISHERY

        If TotCatch(Fish, TStep) = 0 Then
            MsgBox("CNR cannot be estimated for fishery with no catch")
            Stop
        End If

        'TWO POSSIBLE METHODS OF COMPUTING CNR CATCH

        If OOBStatus = 3 Or Firstpass = False Then ' all stocks run


            Select Case CNRFlag(BaseType, Yr, Fish, TStep)

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


                Case 3
                    '            '...External estimate of encounters. Be careful not to exclude computations for stock,age,fish,tstep combos not in CWTList
                    '            ' in OOB run this code splits the total legal and sublegal encounter input into ages (of that stock)
                    '            'doesn't make sense as CNR input is by fishery, not fishery and stock
                    For STk = MinStk To NumStk
                        For Age = 2 To MaxAge
                            If TotCatch(Fish, TStep) > 0 Then
                                PropLegCatch(STk, Age, BaseType) = TotExpCWTAll(STk, Age, Fish, TStep) / TotCatch(Fish, TStep)
                            End If
                            If TotEnc > 0 Then
                                PropSubCatch(STk, Age, BaseType) = (ShakerAll(STk, Age, Fish, TStep) / ShakMortRate(Fish, TStep)) / TotEnc
                            End If
                            If TStep = 3 Then
                                TStep = 3
                            End If

                            '- PS Sport legal size rel mort rate set now to 50% of shaker release rate (10 vs 20)
                            If Fish = 36 Or Fish = 42 Or Fish = 45 Or Fish = 53 Or Fish = 54 Or Fish = 56 Or Fish = 57 Or Fish = 64 Or Fish = 67 Then

                                CNRLeg(STk, Age, Fish, TStep) = PropLegCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep) / 2
                                CNRSub(STk, Age, Fish, TStep) = PropSubCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 2, Fish, TStep) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                            Else

                                CNRLeg(STk, Age, Fish, TStep) = PropLegCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                                CNRSub(STk, Age, Fish, TStep) = PropSubCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 2, Fish, TStep) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                            End If
                            CNR(STk, Age, Fish, TStep) = CNRLeg(STk, Age, Fish, TStep) + CNRSub(STk, Age, Fish, TStep)

                            If CNR(STk, Age, Fish, TStep) > 0 Then
                                Print(22, STk & "," & Age & "," & Fish & "," & TStep & "," & CNRLeg(STk, Age, Fish, TStep) & "," & CNRSub(STk, Age, Fish, TStep) & "," & AnnualCatch(Fish) / TrueCatch(Fish) & "," & "AngelikaNewCNR" & vbCrLf)
                            End If






                        Next Age
                    Next STk

                Case 4 'Total Encounters
                    For STk = MinStk To NumStk
                        For Age = 2 To MaxAge
                            If TotCatch(Fish, TStep) > 0 Then 'ppn of landed catch that is of a given stock and age
                                PropLegCatch(STk, Age, BaseType) = TotExpCWTAll(STk, Age, Fish, TStep) / TotCatch(Fish, TStep)
                            End If
                            If TotEnc > 0 Then 'ppn of sublegal encounters that is of a given stock and age
                                PropSubCatch(STk, Age, BaseType) = (ShakerAll(STk, Age, Fish, TStep) / ShakMortRate(Fish, TStep)) / TotEnc
                            End If
                            If TStep = 3 Then
                                TStep = 3
                            End If

                            'compute ppn legal
                            If TargEncRate(Fish, TStep) <> -1.0! Then
                                PropLeg = 1 / (1 + TargEncRate(Fish, TStep))
                            Else
                                PropLeg = 1
                            End If
                            '- PS Sport legal size rel mort rate set now to 50% of shaker release rate (10 vs 20)
                            If Fish = 36 Or Fish = 42 Or Fish = 45 Or Fish = 53 Or Fish = 54 Or Fish = 56 Or Fish = 57 Or Fish = 64 Or Fish = 67 Then

                                CNRLeg(STk, Age, Fish, TStep) = PropLegCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * PropLeg * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep) / 2
                                CNRSub(STk, Age, Fish, TStep) = PropSubCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * (1 - PropLeg) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                            Else

                                CNRLeg(STk, Age, Fish, TStep) = PropLegCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * PropLeg * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                                CNRSub(STk, Age, Fish, TStep) = PropSubCatch(STk, Age, BaseType) * CNRInput(BaseType, 0, 1, Fish, TStep) * (1 - PropLeg) * AnnualCatch(Fish) / TrueCatch(Fish) * ShakMortRate(Fish, TStep)
                            End If
                            CNR(STk, Age, Fish, TStep) = CNRLeg(STk, Age, Fish, TStep) + CNRSub(STk, Age, Fish, TStep)

                            If CNR(STk, Age, Fish, TStep) > 0 Then
                                Print(22, STk & "," & Age & "," & Fish & "," & TStep & "," & CNRLeg(STk, Age, Fish, TStep) & "," & CNRSub(STk, Age, Fish, TStep) & "," & PropSubCatch(STk, Age, BaseType) & "," & CNRInput(BaseType, 0, 1, Fish, TStep) & "," & (1 - PropLeg) & "," & AnnualCatch(Fish) / TrueCatch(Fish) & "," & ShakMortRate(Fish, TStep) & "," & "AngelikaNewCNR" & vbCrLf)
                            End If


                        Next Age
                    Next STk

            End Select
            'Else


        End If
    End Sub
End Module
