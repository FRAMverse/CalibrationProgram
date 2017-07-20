Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO

Module CompCohortMod
    Sub CompCohort()
        '**********************************************************************
        'Subroutine to reconstruct cohort from CWT recoveries and estimated
        ' shaker and CNR mortality.  Cohort reconstruction proceeds backwards
        ' in time beginning with the oldest age class and last time step.
        '**********************************************************************
        'SUM TOTAL MORTALITY OF STOCK DURING TIME STEP over ages and fisheries
        'find unique stock, BY, stage combinations
        'Dim Cohortfile As String
        'Cohortfile = "C:\data\calibration\07Qbasic\Cohortfile"
        'FileOpen(15, Cohortfile, OpenMode.Output)
        'Print(15, "Cohort,  Stk, age, pterm, tstep" & vbCrLf)




        If Firstpass = True Then 'OOB run
            ReDim Cohort(MaxAge, NumSteps)
            ReDim CohortTerm(MaxAge, NumSteps)
            ReDim TotMortTerm(MaxAge, NumSteps)
            ReDim TotMortPterm(MaxAge, NumSteps)
            ReDim TotalShakerTerm(MaxAge, NumSteps)
            ReDim TotalShakerPTerm(MaxAge, NumSteps)


            For Each CWTRow In CWTList

                If CWTRow.cLookUp = DictionaryKey Then
                    BY = CWTRow.cBY
                    STk = CWTRow.cStk
                    Stage = CWTRow.cStage
                    Age = CWTRow.cAge
                    TStep = CWTRow.cTStep
                    Fish = CWTRow.cFish

                    If TermFlag(Fish, TStep) = 1 Or Fish = 74 Then
                        TotMortTerm(Age, TStep) = TotMortTerm(Age, TStep) + CWTRow.cCatch * ExpFact * (1 + OtherMort(Fish, TStep)) + CNR(Age, Fish, TStep)
                    Else
                        TotMortPterm(Age, TStep) = TotMortPterm(Age, TStep) + CWTRow.cCatch * ExpFact * (1 + OtherMort(Fish, TStep)) + CNR(Age, Fish, TStep)
                    End If
                End If
            Next
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish - 1
                    For TStep = 1 To NumSteps
                        If TermFlag(Fish, TStep) = 1 Or Fish = 74 Then
                            TotalShakerTerm(Age, TStep) = TotalShakerTerm(Age, TStep) + Shaker(Age, Fish, TStep)
                        Else
                            TotalShakerPTerm(Age, TStep) = TotalShakerPTerm(Age, TStep) + Shaker(Age, Fish, TStep)
                        End If
                    Next
                Next
            Next



            For Age = MaxAge To 2 Step -1
                For TStep = NumSteps To 1 Step -1
                    TotMortTerm(Age, TStep) = TotMortTerm(Age, TStep) + TotalShakerTerm(Age, TStep)
                    TotMortPterm(Age, TStep) = TotMortPterm(Age, TStep) + TotalShakerPTerm(Age, TStep)
                Next
            Next
            'cohort = run size after natural mortality
            For Age = MaxAge To 2 Step -1
                For TStep = NumSteps To 1 Step -1
                    If TStep = NumSteps Then
                        If Age = MaxAge Then
                            Cohort(Age, TStep) = TotMortTerm(Age, TStep) + TotMortPterm(Age, TStep)
                        Else
                            Cohort(Age, TStep) = (Cohort(Age + 1, 1) / SurvRate(Age + 1, 1)) + TotMortTerm(Age, TStep) + TotMortPterm(Age, TStep)
                        End If
                    Else
                        Cohort(Age, TStep) = (Cohort(Age, TStep + 1) / SurvRate(Age, TStep + 1)) + TotMortTerm(Age, TStep) + TotMortPterm(Age, TStep)
                    End If

                Next
            Next
        Else 'Allstocks run
            ReDim CohortAll(NumStk, MaxAge, Term, NumSteps)

            For STk = MinStk To NumStk
                If STk = 24 Then
                    STk = 24
                End If
                For Age = MaxAge To 2 Step -1
                    For TStep = NumSteps To 1 Step -1
                       


                        'SUM TOTAL MORTALITY OF STOCK DURING TIME STEP

                        ReDim TotMort(Term)
                        TotMort(Term) = EscExpFact(STk) * CWTAll(STk, Age, 74, TStep)
                        For Fish = 1 To NumFish - 1
                            'If CWTAll(STk, Age, Fish, TStep) > 0 Then
                            If TermFlag(Fish, TStep) = PTerm Then
                                TotMort(PTerm) = TotMort(PTerm) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep) * (1 + OtherMort(Fish, TStep)) + ShakerAll(STk, Age, Fish, TStep) '+ StkMortRec.CNR
                            Else
                                TotMort(Term) = TotMort(Term) + EscExpFact(STk) * CWTAll(STk, Age, Fish, TStep) * (1 + OtherMort(Fish, TStep)) + ShakerAll(STk, Age, Fish, TStep) '+ StkMortRec.CNR
                            End If
                            'End If
                        Next Fish

                        'TERMINAL CohortAll: TOTAL MORTALITY AND ESCAPEMENT OF FISH IN DESIGNATED
                        ' TERMINAL FISHERIES DURING TIMESTEP

                        CohortAll(STk, Age, Term, TStep) = TotMort(Term)

                        'IMMATURE CohortAll: IF END OF YEAR, GET CohortAll ABUNDANCE FROM PREVIOUS YEAR, ONE AGE OLDER
                        ' OTHERWISE USE CohortAll ABUNDANCE FORM SUCCEEDING TIME STEP.
                        ' NOTE THAT CohortAll SIZE COMPUTED IS AFTER NATURAL MORTALITY.

                        If TStep = NumSteps Then
                            If Age = MaxAge Then
                                CohortAll(STk, Age, PTerm, TStep) = CohortAll(STk, Age, Term, TStep) + TotMort(PTerm)
                            Else
                                CohortAll(STk, Age, PTerm, TStep) = (CohortAll(STk, Age + 1, PTerm, 1) / SurvRate(Age + 1, 1)) + CohortAll(STk, Age, Term, TStep) + TotMort(PTerm)
                            End If
                        Else
                            CohortAll(STk, Age, PTerm, TStep) = (CohortAll(STk, Age, PTerm, TStep + 1) / SurvRate(Age, TStep + 1)) + CohortAll(STk, Age, Term, TStep) + TotMort(PTerm)
                        End If
                        'Print(15, CohortAll(STk, Age, PTerm, TStep) & "," & STk & "," & Age & "," & PTerm & "," & TStep & "," & "new", vbCrLf)
                    Next TStep
                Next Age
            Next STk



        End If

    End Sub
End Module
