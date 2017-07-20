Module SaveDatMod
    Sub SaveDat()
        '**********************************************************************
        'Subroutine saves results from cohort analysis/calibration for input to
        ' the forward analysis.
        '**********************************************************************


        ReDim AEQ(NumStk, MaxAge, NumSteps + 1)
        ReDim MatRate(NumStk, MaxAge, NumSteps + 1)
        'ReDim Newstock(NumStk)
        Dim outstep As Integer
        Dim MatAge As Integer
        ReDim ExRateAll(NumStk, MaxAge, NumFish, NumSteps + 1)
        ReDim EncRateAllShaker(NumStk, MaxAge, NumFish, NumSteps + 1)

        'SAVE AEQ RATES
        ' First MatRate must be recalculated for AEQ Calculations
        For OutStep = 1 To NumSteps + 1
            If OutStep <> 4 Then
                For STk = MinStk To NumStk
                    

                    For Age = 2 To MaxAge
                        If CohortAll(STk, Age, Term, outstep) > 0 Then
                            If outstep <> NumSteps + 1 Then
                                If outstep = NumSteps Then
                                    If Age = MaxAge Then
                                        MatRate(STk, Age, outstep) = 1
                                    Else
                                        MatRate(STk, Age, outstep) = CohortAll(STk, Age, Term, outstep) / (CohortAll(STk, Age, Term, outstep) + (CohortAll(STk, Age + 1, PTerm, 1) / SurvRate(Age + 1, 1)))
                                    End If
                                Else
                                    MatRate(STk, Age, outstep) = CohortAll(STk, Age, Term, outstep) / (CohortAll(STk, Age, Term, outstep) + (CohortAll(STk, Age, PTerm, outstep + 1) / SurvRate(Age, outstep + 1)))
                                End If
                            Else
                                If Age <> MaxAge Then
                                    MatRate(STk, Age, outstep) = CohortAll(STk, Age, Term, outstep) / (CohortAll(STk, Age, Term, outstep) + (CohortAll(STk, Age + 1, PTerm, 1) / SurvRate(Age + 1, 1)))
                                Else
                                    MatRate(STk, Age, outstep) = 1
                                End If
                            End If
                        Else
                            If Age = MaxAge And outstep = NumSteps Then
                                MatRate(STk, Age, outstep) = 1
                            Else
                                MatRate(STk, Age, outstep) = 0
                            End If
                        End If
                    Next Age
                Next STk
            Else
                For STk = MinStk To NumStk
                    If STk = 24 Then
                        STk = 24
                    End If

                    For Age = 2 To MaxAge - 1
                        MatAge = Age + 1
                        If MatRate(STk, MatAge, 1) > 0 Then
                            MatRate(STk, Age, 4) = MatRate(STk, MatAge, 1)
                        End If
                    Next Age
                Next STk
            End If
        Next OutStep

        '---- Now Calculate AEQ

        '--- Gary Look here  .... Age 4 Step 3 Stock 10

        '- July 2005 AEQ Fix #3  JFP
        '- The Time 4 AEQ's should be the same as Time 1 
        '- because each Age has been matured for the next year
        '- This requires 2 passes because the Age and TimeStep loops 
        '- use Negative STEP (-1)


        For STk = MinStk To NumStk
            For Age = MaxAge To 2 Step -1
                For TStep = (NumSteps + 1) To 1 Step -1
                    If TStep <> NumSteps + 1 Then
                        If TStep = NumSteps Then
                            If Age = MaxAge Then
                                AEQ(STk, Age, TStep) = 1
                            Else
                                '***** GSM - FIXED AEQ CALCULATION
                                AEQ(STk, Age, TStep) = (MatRate(STk, Age, TStep) + (1 - MatRate(STk, Age, TStep)) * SurvRate(Age + 1, 1)) * AEQ(STk, Age, TStep + 1)
                            End If
                        Else
                            '***** GSM FIXED AEQ CALCULATION
                            AEQ(STk, Age, TStep) = (MatRate(STk, Age, TStep) + (1 - MatRate(STk, Age, TStep)) * SurvRate(Age, TStep + 1)) * AEQ(STk, Age, TStep + 1)
                        End If
                    Else
                        If Age = MaxAge Then
                            AEQ(STk, Age, TStep) = 1.0!
                        Else
                            '***** GSM FIXED AEQ CALCULATION
                            AEQ(STk, Age, TStep) = AEQ(STk, Age + 1, 1)
                        End If
                    End If
                Next TStep
            Next Age
        Next STk

        '- Now Fix Time 4 problem (ie Time 4 = Time 1 AEQ Values)
        For STk = MinStk To NumStk
            For Age = MaxAge To 2 Step -1
                TStep = (NumSteps + 1)
                AEQ(STk, Age, TStep) = AEQ(STk, Age, 1)
            Next Age
        Next STk


       

       

        'SAVE STOCK INCLUSION FLAGS FOR SHAKER COMPUATIONS

        

        'SAVE COHORT SIZE

        ReDim StartCohort(NumStk, MaxAge)
        For STk = MinStk To NumStk
            For Age = 2 To MaxAge
                StartCohort(STk, Age) = CohortAll(STk, Age, PTerm, 1) / SurvRate(Age, 1)
              
            Next Age
        Next STk


        ' SAVE MODEL STOCK PROPORTION 
        ReDim ModelStkPPN(NumFish)
        For Fish = 1 To NumFish

            If TrueCatch(Fish) > 0 Then
                ModelStkPPN(Fish) = AnnualCatch(Fish) / TrueCatch(Fish) 'TrueCatch = BasePeriodCatch

            Else
                ModelStkPPN(Fish) = 1.0
            End If

            'SET PROPORTION TO A SMALL NUMBER FOR FISHERIES WITH CATCH BUT
            ' NONE REPRESENTED IN THE MODEL (REQUIRED FOR REPORT GENERATION)

            If TrueCatch(Fish) > 0 And ModelStkPPN(Fish) < 0.0001 Then
                ModelStkPPN(Fish) = 0.0001
            End If
        Next Fish


        'SAVE MATURATION RATE IF RATE > 0


        'For stocks maturing in time 1 repeat maturation rate in time 4 (of the younger age)
        For STk = MinStk To NumStk
            For Age = 2 To MaxAge - 1
                MatAge = Age + 1
                If MatRate(STk, MatAge, 1) > 0 Then
                    MatRate(STk, Age, 4) = MatRate(STk, MatAge, 1)
                End If
            Next Age
        Next STk



        'SAVE EXPLOITATION RATES IF RATE > 0


        'Dim response = MsgBox("Do you want to use surrogate exploitation rates for Area 6 Sport and Area 9 T Net as in the original base period?", 4, )


        For outstep = 1 To NumSteps + 1
            If outstep <> 4 Then

                For STk = MinStk To NumStk
                    For Fish = 1 To NumFish
                        If STk = 24 And Fish = 30 Then
                            STk = 24
                        End If

                        'below code is now obsolete. Fishery replacements are handled in "ImputeRecoveries" table
                        'If response = MsgBoxResult.Yes Then
                        'replace fishery values for fishery 54, 55 and 63 with values from different fishery
                        'Select Case Fish
                        '    Case 54
                        '        FishNum = 42
                        '    Case 55
                        '        If outstep = 3 Then
                        '            FishNum = 53
                        '        Else
                        '            FishNum = Fish
                        '        End If
                        '    Case Else
                        FishNum = Fish
                        'End Select
                        'Else
                        'FishNum = Fish
                        'End If

                        For Age = 2 To MaxAge
                            If Fish = 47 And STk = 9 And outstep = 2 And Age = 3 Then
                                Age = 3
                            End If

                            TermStat = TermFlag(FishNum, outstep)
                            If CWTAll(STk, Age, FishNum, outstep) > 0 Or ShakerAll(STk, Age, FishNum, outstep) > 0 Then
                                TStep = outstep
                                FishYear = BaseYear
                                Call CompLegProp()
                                If LegalProp > 0 And CWTAll(STk, Age, FishNum, outstep) > 0 Then
                                    If TermFlag(FishNum, outstep) = PTerm Then
                                        ExRateAll(STk, Age, Fish, outstep) = TotExpCWTAll(STk, Age, FishNum, outstep) / (CohortAll(STk, Age, PTerm, outstep) * LegalProp)
                                    Else
                                        ExRateAll(STk, Age, Fish, outstep) = TotExpCWTAll(STk, Age, FishNum, outstep) / (CohortAll(STk, Age, Term, outstep) * LegalProp)
                                    End If
                                End If


                            End If

                            'FOR SHAKERS, EXPLOITATION RATE IS IN TERMS OF ENCOUNTERS

                            If ShakerAll(STk, Age, FishNum, outstep) > 0 Then
                                If STk = 35 And Age = 3 And Fish = 34 And TStep = 2 Then
                                    TStep = 2
                                End If

                                Encounters = ShakerAll(STk, Age, FishNum, outstep) / ShakMortRate(FishNum, outstep)
                                If TermFlag(FishNum, outstep) = PTerm Then
                                    FishCohort = CohortAll(STk, Age, PTerm, outstep)
                                    EncRateAllShaker(STk, Age, Fish, outstep) = Encounters / (CohortAll(STk, Age, PTerm, outstep) * SubLegalProp)
                                Else
                                    FishCohort = CohortAll(STk, Age, PTerm, outstep)
                                    EncRateAllShaker(STk, Age, Fish, outstep) = Encounters / (CohortAll(STk, Age, Term, outstep) * SubLegalProp)
                                End If

                            End If
                        Next Age
                    Next Fish
                Next STk
            Else 'deals with time 4

                For STk = MinStk To NumStk
                    For Fish = 1 To NumFish
                        'If response = MsgBoxResult.Yes Then
                        '    'replace fishery values for fishery 54, 55 and 63 with values from different fishery
                        '    Select Case Fish
                        '        Case 54
                        '            FishNum = 42
                        '        Case 55
                        '            If outstep = 3 Then
                        '                FishNum = 53
                        '            Else
                        '                FishNum = Fish
                        '            End If
                        '        Case Else
                        '            FishNum = Fish
                        '    End Select
                        'Else
                        FishNum = Fish
                        'End If

                        For Age = 3 To MaxAge
                            ExRateAll(STk, Age - 1, Fish, 4) = ExRateAll(STk, Age, Fish, 1)
                            EncRateAllShaker(STk, Age - 1, Fish, 4) = EncRateAllShaker(STk, Age, Fish, 1)
                        Next Age
                    Next Fish
                Next STk
            End If
        Next outstep

       




        FileClose()

    End Sub
End Module
