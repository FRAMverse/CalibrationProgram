Module ForwardMod
    Sub Forward()
        '**********************************************************************
        'Subroutine is used to adjust recoveries to a different base period.
        '  Exploitation rate scale factors are read from an input file and used
        '  to adjust the base period rates.
        '**********************************************************************

       
        'CHECK EXPLOITATION RATES WITHIN EACH TIME STEP TO MAKE SURE RATES
        ' ARE LESS THAN 1
        'AHB: this is calcualated within a stock/BY/STAGE loop and the resulting CompCatch and CompEsc need to be added to a list
        'in AccessDb
        ReDim ExAdjFact(MaxAge, NumSteps, Term)
        ReDim CompCatch(MaxAge, NumFish, NumSteps)
        ReDim CompEscapement(MaxAge, NumSteps)
        Dim MatCohort As Double


       


        'MIXED FISHERIES
        '    Dim TempExRate As Double
        BY = BYNum
        STk = Stknum
        Stage = StageNum
        BYnew = BY - 1980
        FishYear = BaseYear

        '**************************************************************************************************************************************
        'AHB: this code's only purpose is to deal with time step ERs > 1 (catch > run size when summed over all fisheries). Needs to multiply times PPN Legal to 
        'address different size limits and hence legal proportions when summing over ERs for fisheries with a multitude of size limits; i.e. two fisheries with an ER of 60% 
        'may take less than 100% of run size if 

        For Age = 2 To MaxAge
            For TStep = 1 To NumSteps
                If TStep = 3 And STk = 33 And BY = 1996 And Age = 2 Then
                    TStep = 3
                End If
                TotExRate = 0
                'PRETERMINAL FISHERIES

                For Fish = 1 To NumFish - 1
                    If Fish = 15 Then
                        Fish = 15
                    End If
                    FishNum = Fish
                    If TermFlag(Fish, TStep) = PTerm Then
                        ExAdjFact(Age, TStep, PTerm) = 0

                        Call CompLegProp()

                        If FishScalar(BY + Age, Fish, TStep) > 0 Then
                            If ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep) < 1 Then
                                TempExRate = (ExRate(BYnew, STk, Stage, Age, Fish, TStep) * (1 + OtherMort(Fish, TStep)) / FishScalar(BY + Age, Fish, TStep)) * LegalProp
                                TotExRate = TotExRate + TempExRate
                            Else
                                TempExRate = 0.9 ' no fishery can exceed 100% ER, this is later on accounted for when computing cohort size (see below)
                                TotExRate = TotExRate + TempExRate * (1 + OtherMort(Fish, TStep)) * LegalProp
                                Print(13, "BY" & "," & BYnew + 1980 & "," & "Age " & "," & Age & "," & "Fish" & "," & Fish & "," & "TS" & "," & TStep & "," & "Scalar" & "," & FishScalar(BY + Age, Fish, TStep) & "," & "ER" & "," & ExRate(BYnew, STk, Stage, Age, Fish, TStep) & vbCrLf)
                            End If

                            'THIS CODE PICKS UP FISHERIES WHERE NO RECOVERIES OCCURRED IN THE BASE PERIOD
                            ' AND SCALES THE CATCH TO 1% OF THE OBSERVED VALUE.

                        Else 'AHB:fishery was closed during base period (what about fisheries that were closed in OOB year?)
                            ' designate surrogate fisheries (hence surrogate fisheries can be different for OOB stocks than base period stocks)
                            TempExRate = 0.01 * ExRate(BYnew, STk, Stage, Age, Fish, TStep) * LegalProp
                            TotExRate = TotExRate + TempExRate * (1 + OtherMort(Fish, TStep))
                        End If
                    End If
                Next Fish
                If TotExRate > 1.0! Then 'summed over all fisheries
                    ExAdjFact(Age, TStep, PTerm) = 0.95 / TotExRate
                    MsgBox("Sum of Preterminal Exploitation Rate >1" & " Stk= " & STk & " BY= " & BY & "Stage= " & Stage & " Age= " & Age & " Time= " & TStep)

                Else
                    ExAdjFact(Age, TStep, PTerm) = 1
                End If
                'Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & TStep & "," & "Pterm" & "," & TotExRate & vbCrLf)
            Next TStep
        Next Age

        'MATURE FISHERIES

        For Age = 2 To MaxAge
            For TStep = 1 To NumSteps
                TotExRate = 0
                For Fish = 1 To NumFish - 1
                    FishNum = Fish
                    If TermFlag(Fish, TStep) = Term Then
                        ExAdjFact(Age, TStep, Term) = 0
                        Call CompLegProp()
                        If FishScalar(BY + Age, Fish, TStep) > 0 Then
                            If ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep) < 1 Then
                                TempExRate = (ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep)) * LegalProp
                                TotExRate = TotExRate + TempExRate * (1 + OtherMort(Fish, TStep))
                            Else
                                TempExRate = 0.9
                                TotExRate = TotExRate + TempExRate * (1 + OtherMort(Fish, TStep)) * LegalProp
                                Print(13, "BY" & "," & BYnew + 1980 & "," & "Age " & "," & Age & "," & "Fish" & "," & Fish & "," & "TS" & "," & TStep & "," & "Scalar" & "," & FishScalar(BY + Age, Fish, TStep) & "," & "ER" & "," & ExRate(BYnew, STk, Stage, Age, Fish, TStep) & vbCrLf)
                            End If

                            'THIS CODE PICKS UP FISHERIES WHERE NO RECOVERIES OCCURRED IN THE BASE PERIOD
                            ' AND SCALES THE CATCH TO 1% OF THE OBSERVED VALUE.

                        Else
                            TempExRate = 0.01 * ExRate(BYnew, STk, Stage, Age, Fish, TStep) * LegalProp
                            TotExRate = TotExRate + TempExRate * (1 + OtherMort(Fish, TStep))
                        End If
                    End If
                Next Fish
                If TotExRate > 1.0! Then
                    ExAdjFact(Age, TStep, Term) = 0.95 / TotExRate
                    MsgBox("Sum of Terminal Exploitation Rate >1" & " Stk= " & STk & " BY= " & BY & " Age= " & Age & " Time= " & TStep)

                Else
                    ExAdjFact(Age, TStep, Term) = 1
                End If
                ' Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & TStep & "," & "Term" & "," & TotExRate & vbCrLf)
            Next TStep
        Next Age
        '**************************************************************************************************************************************
        'GET INITIAL COHORT********
        Cohort2 = Cohort(2, 1) / SurvRate(2, 1) 'AHB: just needed to initalize the value for age 2s in order to work within the age loop (Cohort2 within loop is abunandance by age (age2-5) post nat mort)

        For Age = 2 To MaxAge
            For TStep = 1 To NumSteps

                'MIXED FISHERIES

                Cohort2 = Cohort2 * SurvRate(Age, TStep)
                MixedCatch = 0
                For Fish = 1 To NumFish
                    If Fish = 45 And TStep = 3 Then
                        TStep = 3
                    End If
                    FishNum = Fish

                    If TermFlag(Fish, TStep) = PTerm Then
                        CompLegProp() 'should use BP size limit: FishYear = BaseYear


                        If FishScalar(BY + Age, Fish, TStep) > 0 Then
                            If ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep) < 1 Then
                                CompCatch(Age, Fish, TStep) = ExAdjFact(Age, TStep, PTerm) * (ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep)) * Cohort2 * LegalProp
                                MixedCatch = MixedCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                            Else
                                ExRate(BYnew, STk, Stage, Age, Fish, TStep) = 0.9 * ExAdjFact(Age, TStep, PTerm)
                                CompCatch(Age, Fish, TStep) = ExRate(BYnew, STk, Stage, Age, Fish, TStep) * Cohort2 * LegalProp
                                MixedCatch = MixedCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                                MsgBox("Preterminal Exploitation Rate >1" & ";Fishery= " & Fish & " Age= " & Age & " Time= " & TStep)
                                'CompCatch(Age, Fish, TStep) = ExAdjFact(Age, TStep, PTerm) * (ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep)) * Cohort2 * LegalProp


                            End If

                            'THIS CODE PICKS UP FISHERIES WHERE NO RECOVERIES OCCURRED IN THE BASE PERIOD
                            ' AND SCALES THE CATCH TO 1% OF THE OBSERVED VALUE.

                        Else
                            CompCatch(Age, Fish, TStep) = 0.01 * ExAdjFact(Age, TStep, PTerm) * ExRate(BYnew, STk, Stage, Age, Fish, TStep) * Cohort2 * LegalProp
                            MixedCatch = MixedCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                        End If
                        'If CompCatch(Age, Fish, TStep) <> 0 Then
                        '    'Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & Fish & "," & TStep & "," & CompCatch(Age, Fish, TStep) & vbCrLf)
                        'End If

                    End If

                Next Fish

                'COMPUTE POST MIXED FISHERY COHORT

                Cohort2 = Cohort2 - MixedCatch ' AHB: no attempt to deal with sublegal mortality or CNR?

                'COMPUTE MATURE RUN

                MatCohort = Cohort2 * MatRate(STk, Age, TStep)

                'COMPUTE MATURE CATCH

                MatCatch = 0
                For Fish = 1 To NumFish
                    FishNum = Fish
                    If TermFlag(Fish, TStep) = Term Then
                        If FishScalar(BY + Age, Fish, TStep) > 0 Then
                            If ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep) < 1 Then
                                CompLegProp()
                                CompCatch(Age, Fish, TStep) = ExAdjFact(Age, TStep, Term) * (ExRate(BYnew, STk, Stage, Age, Fish, TStep) / FishScalar(BY + Age, Fish, TStep)) * MatCohort * LegalProp
                                MatCatch = MatCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                            Else
                                ExRate(BYnew, STk, Stage, Age, Fish, TStep) = 0.9 * ExAdjFact(Age, TStep, Term)
                                CompCatch(Age, Fish, TStep) = ExRate(BYnew, STk, Stage, Age, Fish, TStep) * MatCohort * LegalProp
                                MatCatch = MatCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                                MsgBox("Terminal Exploitation Rate >1" & ";Fishery= " & Fish & " Age= " & Age & " Time= " & TStep)
                                'Print(13, "BY" & "," & BYnew + 1980 & "," & "Age " & "," & Age & "," & "Fish" & "," & Fish & "," & "TS" & "," & TStep & "," & "Scalar" & "," & FishScalar(BY + Age, Fish, TStep) & "," & "ER" & "," & ExRate(BYnew, STk, Stage, Age, Fish, TStep) & vbCrLf)
                            End If
                        Else
                            Call CompLegProp()
                            CompCatch(Age, Fish, TStep) = 0.01 * ExAdjFact(Age, TStep, Term) * ExRate(BYnew, STk, Stage, Age, Fish, TStep) * MatCohort * LegalProp
                            MatCatch = MatCatch + CompCatch(Age, Fish, TStep) * (1 + OtherMort(Fish, TStep))
                        End If
                        If CompCatch(Age, Fish, TStep) <> 0 Then
                            'Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & Fish & "," & TStep & "," & CompCatch(Age, Fish, TStep) & vbCrLf)
                        End If
                    End If

                Next Fish

                'COMPUTE(ESCAPEMENT)

                CompEscapement(Age, TStep) = MatCohort - MatCatch

                'COMPUTE REMAINING COHORT

                Cohort2 = Cohort2 * (1 - MatRate(STk, Age, TStep))

                If CompEscapement(Age, TStep) <> 0 Then
                    'Print(13, BYnew + 1980 & "," & STk & "," & Stage & "," & Age & "," & "74" & "," & TStep & "," & CompEscapement(Age, TStep) & vbCrLf)
                    imputerecoveriesmain.cCatch = CompEscapement(Age, TStep)
                    imputerecoveriesmain.cFish = 74
                    imputerecoveriesmain.cBY = BY
                    imputerecoveriesmain.cStk = STk
                    imputerecoveriesmain.cAge = Age
                    imputerecoveriesmain.cStage = Stage
                    imputerecoveriesmain.cTStep = TStep

                    ImputeListMain.Add(imputerecoveriesmain)
                    TotLanded += CompEscapement(Age, TStep)
                End If

                For Fish = 1 To NumFish - 1
                    If CompCatch(Age, Fish, TStep) <> 0 Then
                        imputerecoveriesmain.cCatch = CompCatch(Age, Fish, TStep)
                        imputerecoveriesmain.cFish = Fish
                        imputerecoveriesmain.cBY = BY
                        imputerecoveriesmain.cStk = STk
                        imputerecoveriesmain.cAge = Age
                        imputerecoveriesmain.cStage = Stage
                        imputerecoveriesmain.cTStep = TStep

                        ImputeListMain.Add(imputerecoveriesmain)
                        TotLanded += CompCatch(Age, Fish, TStep)
                    End If
                Next

            Next TStep
        Next Age

        TotLandedItem.bBY = BY
        TotLandedItem.bStage = Stage
        TotLandedItem.bStk = STk
        TotLandedItem.bTotalCatch = TotLanded

        TotalLandedList.Add(TotLandedItem)


    End Sub
End Module
