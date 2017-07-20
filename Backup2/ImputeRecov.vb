Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Imports System.Math
Module ImputeRecov
    ' Impute recoveries
    Sub Impute()
        Dim IFish As Integer
        Dim ITStep As Integer
        Dim SFish As Integer
        Dim STStep As Integer
        Dim Imputecatch As String
        Dim FishTSCatchOrig(NumFish, NumSteps) As Double
        Dim FishTSCatchNew(NumFish, NumSteps) As Double
        Dim CWTAllOrig(NumStk, MaxAge, NumFish, NumSteps) As Double
        Dim AdjCheck(NumFish, NumSteps) As Boolean
        Imputecatch = filepath & "\" & "ImputeCatch.txt"
        FileOpen(15, Imputecatch, OpenMode.Output)
        Print(15, "Stock,  Age, Fish, TStep, CWT" & vbCrLf)

        '*******************************************************************************************************************************
        'AHB 8/21/2015
        '
        'Type 0 for surrogates needed for OOB stocks
        'Type 1 combines recoveries for a fishery from different time steps. This is done before surrogates are processed. recipeint time step is deleted
        'Type 111 combines recoveries for a fishery from different time steps. Recipient time step is included in combo.
        'Type 2 flag reassigns or combines CWTs to new time step or fishery based on records in ImputeRecoveries table'Type'= original concept
        'Type 3 zero out fishery and time step
        'Type 9 applies exploitation rates from old base period (processed in ImputeOldBPERs(), not here)
       
        If OOBStatus = 1 Or Firstpass = True Then
            BaseType = 0
        Else
            BaseType = 1
        End If 'during the last base period calibration OOB stocks had different imputed fisheries than Allstocks

        If BaseType = 0 Then
            'this code needs to be reviewed to make sure it works with all the new BaseType options!!!! AHB 8/21/15
            Dim newlist As New List(Of CWTData)
            Dim newlist2 As New List(Of CWTData)
            Dim imputerecoveries As CWTData
            For Each ImputeItem In ImputeList
                If BaseType = ImputeItem.cBaseType Then
                    'find all the records that meet criteria set in findRecord
                    sublist = CWTList.FindAll(AddressOf findRecord)
                    For Each RecordCWT In sublist
                        imputerecoveries.cCatch = Math.Round(RecordCWT.cCatch / 1000, 4)
                        imputerecoveries.cFish = ImputeItem.cRecipientFish
                        imputerecoveries.cBY = RecordCWT.cBY
                        imputerecoveries.cStk = RecordCWT.cStk
                        imputerecoveries.cAge = RecordCWT.cAge
                        imputerecoveries.cStage = RecordCWT.cStage
                        imputerecoveries.cTStep = ImputeItem.cRecipientTStep
                        imputerecoveries.cLookUp = RecordCWT.cLookUp
                        newlist.Add(imputerecoveries)
                    Next
                    'newlist.AddRange(sublist) 'add sublist to newlist for each row in ImputeItem
                End If
            Next

            Dim Deletelist As New List(Of CWTData)
            For Each ImputeItem In ImputeList
                If BaseType = ImputeItem.cBaseType Then
                    'delete records that are going to be replaced with the source fishery
                    'Deletelist = CWTList.FindAll(AddressOf FindDelRec)
                    CWTList.RemoveAll(AddressOf FindDelRec)
                End If
            Next

            CWTList.AddRange(newlist)
        End If
        '******************************************************
        'impute for all stocks file using array

        'store original fishery catches
        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish - 1
                    For TStep = 1 To NumSteps
                        FishTSCatchOrig(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                        CWTAllOrig(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                    Next
                Next
            Next
        Next

        'add CWTs from source fish/ts to imputed fish/ts 
        'if BaseType = 1, delete source value; if BaseType = 11, divide sum by 1000; if BaseType = 111 add time steps preserve source value
        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 1 Or ImputeItem.cBaseType = 11 Or ImputeItem.cBaseType = 111 Then
                IFish = ImputeItem.cRecipientFish
                ITStep = ImputeItem.cRecipientTStep
                SFish = ImputeItem.cSurrogateFish
                STStep = ImputeItem.cSurrogateTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        If ImputeItem.cBaseType = 1 Then
                            CWTAll(STk, Age, IFish, ITStep) += CWTAllOrig(STk, Age, SFish, STStep)
                            CWTAll(STk, Age, SFish, STStep) = 0
                        End If

                        'If ImputeItem.cBaseType = 11 Then
                        '    CWTAll(STk, Age, IFish, ITStep) += Math.Round(CWTAllOrig(STk, Age, SFish, STStep) / 1000, 4, MidpointRounding.AwayFromZero)
                        'End If

                        'If ImputeItem.cBaseType = 111 Then
                        '    CWTAll(STk, Age, IFish, ITStep) += CWTAllOrig(STk, Age, SFish, STStep)
                        'End If

                    Next
                Next
            End If
        Next

      

        'zero out CWTs to be replaced then start the surrogate process that also allows for adding fisheries and time steps
        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 2 Or ImputeItem.cBaseType = 3 Then
                IFish = ImputeItem.cRecipientFish
                ITStep = ImputeItem.cRecipientTStep
                SFish = ImputeItem.cSurrogateFish
                STStep = ImputeItem.cSurrogateTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) = 0
                        If ImputeItem.cBaseType = 2 Then
                            CWTAll(STk, Age, IFish, ITStep) += Math.Round(CWTAll(STk, Age, SFish, STStep) / 1000, 4, MidpointRounding.AwayFromZero)
                        End If
                    Next
                Next
            End If
        Next

        'For debugging purposes, recompute FishTSCatchNew and output CWTAll where > 0
        ReDim FishTSCatchNew(NumFish, NumSteps)
        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish
                    For TStep = 1 To NumSteps
                        If CWTAll(STk, Age, Fish, TStep) > 0 Then
                            Print(15, STk & "," & Age & "," & Fish & "," & TStep & "," & CWTAll(STk, Age, Fish, TStep) & vbCrLf)
                        End If
                    Next
                Next
            Next
        Next STk

        FileClose(15)
    End Sub
    '########## Begin JC Update; 9/25/2015 ##########
    'subroutine to compute CWTs for fisheries using old BPERs as surrogates after cohort sizes are available or change with each iteration
    '!!! method needs to be applied to the entire fishery (all time steps) !!!!

    Sub ImputeOldBPERs()

        Dim SFish As Integer
        Dim STStep As Integer
        Dim FishTSCatchOrig(NumFish, NumSteps) As Double
        Dim FishTSCatchNew(NumFish, NumSteps) As Double
        ReDim BPERFisheries(NumFish)

        Dim ExpCWTfile As String
        ExpCWTfile = filepath & "ExpCWTfile.txt"
        FileOpen(3, ExpCWTfile, OpenMode.Output)

        'zero out arrays &  Estimate CWT catch for fisheries using Old BPER as surrogate

        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 9 Then
                Fish = ImputeItem.cRecipientFish
                FishNum = Fish
                If Fish = 67 Then
                    Fish = 67
                End If
                TStep = ImputeItem.cRecipientTStep
                SFish = ImputeItem.cSurrogateFish
                STStep = ImputeItem.cSurrogateTStep
                FishYear = 2010
                TermStat = TermFlag(FishNum, TStep)

                TimeCatch(Fish, TStep) = 0
                TotCatch(Fish, TStep) = 0
                AnnualCatch(Fish) = 0
                For STk = MinStk To NumStk
                    StockCatch(STk, Fish) = 0
                    StockCatchProp(STk, Fish) = 0
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, Fish, TStep) = 0
                        TotExpCWTAll(STk, Age, Fish, TStep) = 0
                    Next
                Next STk

                BPERFisheries(Fish) = 1

                'Estimate recoveries using old BPERs
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CompLegProp()
                        CWTAll(STk, Age, Fish, TStep) = CohortAll(STk, Age, TermStat, TStep) * SurrogateFishBP_ER(STk, Age, SFish, STStep) * LegalProp
                        'this is equivalent to CWTAll expanded for escapement (PEF)
                    Next
                Next
            End If
        Next

        'ADD CATCH
        'COMPUTE TOTAL CATCH IN EACH FISHERY
        For Fish = 1 To NumFish

            If BPERFisheries(Fish) = 1 Then

                For TStep = 1 To NumSteps
                    For STk = MinStk To NumStk
                        For Age = 2 To MaxAge
                            If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + CWTAll(STk, Age, Fish, TStep)
                                AnnualCatch(Fish) = AnnualCatch(Fish) + CWTAll(STk, Age, Fish, TStep)
                                StockCatch(STk, Fish) = StockCatch(STk, Fish) + CWTAll(STk, Age, Fish, TStep)
                                StockCatchProp(STk, Fish) = StockCatchProp(STk, Fish) + CWTAll(STk, Age, Fish, TStep)
                                TimeCatch(Fish, TStep) = TimeCatch(Fish, TStep) + CWTAll(STk, Age, Fish, TStep)
                            End If
                        Next Age
                    Next STk
                Next TStep

            End If
        Next


        'ADJUST CATCH
        For Fish = 1 To NumFish
            If BPERFisheries(Fish) = 1 Then
                Select Case CatchFlag(Fish) 'located in cal file to the right of base period fishery catches or in BasePeriodCatch table of Calibration Support db

                    ' ADJUST MODEL CATCH TO ESTIMATE CATCH
                    Case 1
                        RecAdjFactor(Fish) = TrueCatch(Fish) / AnnualCatch(Fish) 'true catch = base period catch 
                        For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        TotExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)
                                    End If
                                Next Age
                            Next STk
                        Next TStep

                        ' ADJUST MODEL CATCH IF GREATER THAN ESTIMATED CATCH

                    Case 2
                        RecAdjFactor(Fish) = TrueCatch(Fish) / AnnualCatch(Fish)
                        If RecAdjFactor(Fish) < 1 Then
                            For TStep = 1 To NumSteps
                                For STk = MinStk To NumStk
                                    For Age = 2 To MaxAge
                                        If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                            TotExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)
                                        End If
                                    Next Age
                                Next STk
                            Next TStep
                        Else
                            RecAdjFactor(Fish) = 99

                            For TStep = 1 To NumSteps
                                For STk = MinStk To NumStk
                                    For Age = 2 To MaxAge
                                        If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                            TotExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                        End If
                                    Next Age
                                Next STk
                            Next TStep
                        End If

                        ' ADJUST MODEL CATCH TO ESTIMATE CATCH
                    Case 3

                        RecAdjFactor(Fish) = TrueCatch(Fish) * ExternalModelStockProportion(Fish) / AnnualCatch(Fish) 'true catch = base period catch 


                        For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        TotExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)

                                    End If
                                Next Age
                            Next STk
                        Next TStep


                        'NO ADJUSTMENT
                    Case Else
                        For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        TotExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                    End If
                                Next Age
                            Next STk
                        Next TStep
                        RecAdjFactor(Fish) = 99
                End Select
            End If
        Next


        'ADD CATCH
        'COMPUTE TOTAL CATCH IN EACH FISHERY
        For Fish = 1 To NumFish
            If BPERFisheries(Fish) = 1 Then
                AnnualCatch(Fish) = 0

                For TStep = 1 To NumSteps
                    TimeCatch(Fish, TStep) = 0
                    TotCatch(Fish, TStep) = 0
                Next
                For STk = MinStk To NumStk
                    StockCatch(STk, Fish) = 0
                    StockCatchProp(STk, Fish) = 0
                Next STk

                For TStep = 1 To NumSteps
                    For STk = MinStk To NumStk
                        For Age = 2 To MaxAge
                            If TotExpCWTAll(STk, Age, Fish, TStep) > 0 Then
                                TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + TotExpCWTAll(STk, Age, Fish, TStep)
                                AnnualCatch(Fish) = AnnualCatch(Fish) + TotExpCWTAll(STk, Age, Fish, TStep)
                                StockCatch(STk, Fish) = StockCatch(STk, Fish) + TotExpCWTAll(STk, Age, Fish, TStep)
                                StockCatchProp(STk, Fish) = StockCatchProp(STk, Fish) + TotExpCWTAll(STk, Age, Fish, TStep)
                                TimeCatch(Fish, TStep) = TimeCatch(Fish, TStep) + TotExpCWTAll(STk, Age, Fish, TStep)
                            End If
                        Next Age
                    Next STk
                Next TStep


                'COMPUTE PROPORTION OF CATCH COMPRISED OF EACH STOCK
                If AnnualCatch(Fish) > 0 Then
                    For STk = MinStk To NumStk
                        StockCatchProp(STk, Fish) = StockCatch(STk, Fish) / AnnualCatch(Fish)
                    Next STk
                End If

            End If
        Next Fish
        Print(3, "Stock" & "," & "Age" & "," & "Fish" & "," & "TStep" & "," & "Catch" & vbCrLf)
        For TStep = 1 To NumSteps
            For Fish = 1 To NumFish - 1
                If Fish = 4 And TStep = 3 Then
                    TStep = 3
                End If
                For STk = MinStk To NumStk
                    For Age = 2 To MaxAge
                        If TotExpCWTAll(STk, Age, Fish, TStep) > 0 Then
                            Print(3, STk & "," & Age & "," & Fish & "," & TStep & "," & TotExpCWTAll(STk, Age, Fish, TStep) & vbCrLf)
                        End If
                    Next Age
                Next STk
            Next Fish
        Next TStep
        FileClose(3)

        

    End Sub '########## End JC Update; 9/25/2015 ##########

    'Function findRecord(ByVal b As CWTData) As Boolean
    Function findRecord(ByVal b As CWTData) As Boolean
        If (b.cFish = ImputeItem.cSurrogateFish And b.cTStep = ImputeItem.cSurrogateTStep) Then

            Return True
        Else
            Return False
        End If
    End Function
    Function FindDelRec(ByVal d As CWTData) As Boolean 'find records to delete 
        n = n + 1
        If (d.cFish = ImputeItem.cRecipientFish And d.cTStep = ImputeItem.cRecipientTStep) Then
            Return True
        Else
            Return False
        End If

    End Function
End Module
