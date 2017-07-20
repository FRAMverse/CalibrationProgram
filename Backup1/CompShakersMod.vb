Module CompShakersMod
    Sub CompShakers()
        'Assigns stock age composition to sublegal encounters using legal exploitaton rates
        'If there isn't legal ER for an age, ER from next higher age will be assigned, etc. 
        'Sublegal Encounters are computed as Landed Catch * S/L Ratios (TargEncRate)
       

        Dim SublegalStock(NumStk, MaxAge) As Double
        Dim SumSublegalStock As Double
        Dim SublegalPPN(NumStk, MaxAge) As Double
        Dim SublegalStkAge(NumStk, MaxAge) As Double
        Dim SubEncRate(NumStk, MaxAge) As Double
        Dim StkProp(NumStk) As Double 'proportion of sublegal encounter of a stock in a fishery
        Dim StkSum(NumStk) As Double 'sublegal encounters of a stock in a fishery summed over ages
        Dim AgeProp(NumStk, MaxAge) As Double
        Dim AnnualCatchMod(NumFish) As Double 'annual fishery catch of stocks with a sublegal population-prevents stock assignments of sublegal
        Dim FishSublegalPop As Double ' total number of sublegals in a fishery summed over all the stocks encountered in fishery
        Dim FindRate As Double = 0
        Dim LegalRate(MaxAge) As Double
        Dim SubLegalRate(MaxAge) As Double
        Dim SaveAge As Integer
        Dim newage As Integer
        Dim StkinFishTS(NumStk) As Double
        Dim LegProp(NumStk, MaxAge) As Double
        Dim NewStockCatchProp(NumStk) As Double
        Dim StkSubLegalPop(NumStk) As Double
        Dim NewTotSublegalPop As Double

       
        TermStat = 0

        'COMPUTE SHAKER MORTALITY IF CATCH OCCURRED IN FISHERY
        FishSublegalPop = 0
        If TotCatch(Fish, TStep) > 0 Then
            

            'COMPUTE TOTAL NUMBER OF ENCOUNTERS
            'in OOB run, this sets sublegal encounters to the catch of a stock in a fishery and time step unless TargetEncRate is set 
            If TargEncRate(Fish, TStep) <> -1.0! Then
                TotEnc = TimeCatch(Fish, TStep) * TargEncRate(Fish, TStep)
            Else
                TotEnc = 0
            End If

            'COMPUTE SUBLEGAL POPULATION FOR STOCK
            ' FOR FISHERIES WITH SUBLEGAL ENCOUNTERS
            ' USE ONLY STOCKS WITH CATCH IN THIS FISHERY

            If TotEnc > 0 Then
                For STk = MinStk To MaxStk
                    


                    If StockCatchProp(STk, Fish) > 0 Then ' catch of this stock in this fishery = 1 for OOB
                        TotSubLegalPop = 0

                        If Firstpass = True Then
                            For Age = 2 To MaxAge


                                FishYear = BYNum + Age

                                CompLegProp()

                                PropSubPop(STk, Age) = Cohort(Age, TStep) * SubLegalProp

                                TotSubLegalPop = TotSubLegalPop + PropSubPop(STk, Age)
                            Next Age

                            'COMPUTE PROPORTION OF SUBLEGAL POPULATION IN EACH AGE CLASS
                            'FOR THIS STOCK

                            For Age = 2 To MaxAge
                                If TotSubLegalPop > 0 Then
                                    PropSubPop(STk, Age) = PropSubPop(STk, Age) / TotSubLegalPop
                                Else
                                    PropSubPop(STk, Age) = 0
                                End If
                            Next Age

                            'COMPUTE SHAKERS FOR THIS STOCK

                            For Age = 2 To MaxAge

                                Shaker(Age, Fish, TStep) = TotEnc * PropSubPop(STk, Age) * ShakMortRate(Fish, TStep) * StockCatchProp(STk, Fish)
                                'Print(13, BY & "," & STk & "," & Stage & "," & Age & "," & Fish & "," & TStep & "," & Shaker(Age, Fish, TStep) & vbCrLf)
                            Next Age
                        Else 'ALLSTOCKS RUN
                            FishYear = BaseYear
                            For Age = 2 To MaxAge
                                CompLegProp()

                                PropSubPop(STk, Age) = CohortAll(STk, Age, TermStat, TStep) * SubLegalProp
                                TotSubLegalPop = TotSubLegalPop + PropSubPop(STk, Age)

                                FishSublegalPop = FishSublegalPop + TotSubLegalPop 'total sublegalpop summed over stocks in a fishery tstep
                                'StkinFishTS(STk) = CWTAll(STk, Age, Fish, TStep) + StkinFishTS(STk)
                                LegProp(STk, Age) = LegalProp
                            Next Age

                            If TotSubLegalPop > 0 Then
                                For Age = 2 To MaxAge
                                    AgeProp(STk, Age) = PropSubPop(STk, Age) / TotSubLegalPop
                                Next
                            End If
                            



                        End If 'firstpass true
                    End If ' stockCatchProp > 0
                Next STk 'end result for GalenMethod = 1 is # sublegals by stock and age as well as total sublegal for fishery

                ' deal with stocks without sublegal population in a fishery, need to redistribute impacts over remaining stocks





                If Firstpass = False Then


                    For STk = 1 To NumStk
                        For Age = 2 To MaxAge ' sum sublegals for a stock over ages
                            CompLegProp()
                            StkSubLegalPop(STk) = StkSubLegalPop(STk) + CohortAll(STk, Age, PTerm, TStep) * SubLegalProp
                        Next Age
                    Next STk

                    For STk = 1 To NumStk
                        If StkSubLegalPop(STk) > 0 Then 'only add stock with sublegal pop > 0

                            NewTotSublegalPop = NewTotSublegalPop + StockCatchProp(STk, Fish)

                        End If
                    Next STk

                    For STk = 1 To NumStk
                        If StkSubLegalPop(STk) > 0 Then
                            For Age = 2 To MaxAge
                                NewStockCatchProp(STk) = StockCatchProp(STk, Fish) / NewTotSublegalPop
                            Next Age
                        End If
                    Next

                End If
                '________________________________________________________________________________________________________

                If Firstpass = False Then

                    SumSublegalStock = 0
                    For STk = 1 To NumStk

                        'If StkinFishTS(STk) > 0 Then

                        If StockCatchProp(STk, Fish) > 0 Then
                            '############################################# JON VERSION #########################################
                            For Age = 2 To MaxAge
                                'If Fish = 31 And TStep = 3 Then
                                '    TStep = 3
                                'End If
                                ' compute legal exploitation rates for all ages of stock in fishery, timestep

                                If CohortAll(STk, Age, 0, TStep) * LegProp(STk, Age) <> 0 Then
                                    LegalRate(Age) = CWTAll(STk, Age, Fish, TStep) * EscExpFact(STk) / (CohortAll(STk, Age, 0, TStep) * LegProp(STk, Age))

                                Else
                                    LegalRate(Age) = 0
                                End If
                            Next
                            For Age = 2 To MaxAge

                                'if legal exploitation rate is > 0 for age, set it as sublegal exploitation rate
                                If LegalRate(Age) > 0 Then

                                    SubLegalRate(Age) = LegalRate(Age)
                                Else
                                    'if LegalRate(Age) is 0:
                                    '   set SubLegalRate(Age) to 0 if LegalProp(Age) is > 0.5
                                    '   otherwise move to Age+1
                                    '       if LegalRate(Age+1) > 0 then SubLegalRate(Age) = LegalRate(Age+1)
                                    '       if LegalRate(Age+1) = 0 then SubLegalRate(Age) = 0 if LegalProp(Age+1) is > 0.5
                                    '   otherwise move to Age+2, etc...
                                    For newage = Age To MaxAge
                                        If LegalRate(newage) <> 0 Then

                                            SubLegalRate(Age) = LegalRate(newage)
                                            newage = 5
                                        Else
                                            SaveAge = Age
                                            Age = newage

                                            If LegProp(STk, Age) > 0.5 Then
                                                SubLegalRate(SaveAge) = 0
                                                newage = 5
                                            End If
                                            Age = SaveAge
                                            'If newage <> SaveAge Then
                                            '    CompLegProp()
                                            'End If
                                        End If
                                    Next
                                End If

                                SublegalStock(STk, Age) = 0

                                If 1 - LegProp(STk, Age) <> 0 Then
                                    SublegalStock(STk, Age) = PropSubPop(STk, Age) * SubLegalRate(Age)
                                    SumSublegalStock = SumSublegalStock + SublegalStock(STk, Age)
                                End If
                            Next 'age

                        End If
                        'End If 'stock in fish
                    Next 'stock

                    For STk = 1 To NumStk
                        For Age = 2 To MaxAge
                            If SumSublegalStock > 0 Then
                                SublegalPPN(STk, Age) = SublegalStock(STk, Age) / SumSublegalStock
                                ShakerAll(STk, Age, Fish, TStep) = TotEnc * SublegalPPN(STk, Age) * ShakMortRate(Fish, TStep)
                                'ShakerAll(STk, Age, Fish, TStep) = TotEnc * SublegalPPN(STk, Age) 'Encounters
                                If ShakerAll(STk, Age, Fish, TStep) > 0 Then
                                    Print(22, STk & "," & Age & "," & Fish & "," & TStep & "," & "," & ShakerAll(STk, Age, Fish, TStep) & "," & "AngelikaNew" & vbCrLf)
                                End If
                            Else ' when size limit is small this method will not produce any sublegals
                                ShakerAll(STk, Age, Fish, TStep) = TotEnc * AgeProp(STk, Age) * ShakMortRate(Fish, TStep) * NewStockCatchProp(STk)
                                'ShakerAll(STk, Age, Fish, TStep) = TotEnc * AgeProp(STk, Age) * NewStockCatchProp(STk) 'Encounters
                                If ShakerAll(STk, Age, Fish, TStep) > 0 Then
                                    Print(22, STk & "," & Age & "," & Fish & "," & TStep & "," & "," & ShakerAll(STk, Age, Fish, TStep) & "," & "AngelikaNew" & vbCrLf)
                                End If

                            End If
                        Next
                    Next
                End If 'first pass = false
            End If 'Tot Enc > 0
        End If ' Tot Catch > 0
    End Sub
End Module
