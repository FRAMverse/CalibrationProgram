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
        'Type 1 combines recoveries for a fishery from different time steps. This is done before surrogates are processed. Source time step is deleted
        'Type 11 combines recoveries for a fishery from different time steps. Source time step is included in combo. Fisheries are rescaled to original size.
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
                        imputerecoveries.cFish = ImputeItem.cImputedFish
                        imputerecoveries.cBY = RecordCWT.cBY
                        imputerecoveries.cStk = RecordCWT.cStk
                        imputerecoveries.cAge = RecordCWT.cAge
                        imputerecoveries.cStage = RecordCWT.cStage
                        imputerecoveries.cTStep = ImputeItem.cImputedTStep
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

        'move CWTs from source fish/ts into imputed fish/ts for BaseTypes 1 and 11
        'if BaseType = 1, delete source value; if BaseType = 11, retain source value
        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 1 Or ImputeItem.cBaseType = 11 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) += CWTAllOrig(STk, Age, SFish, STStep)
                        If ImputeItem.cBaseType = 1 Then
                            CWTAll(STk, Age, SFish, STStep) = 0
                        End If
                    Next
                Next
            End If
        Next

        'calculate new total catch by fish/ts
        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish - 1
                    For TStep = 1 To NumSteps
                        FishTSCatchNew(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                    Next
                Next
            Next
        Next

        'scale summed catches back to original catch
        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 11 Then
                Fish = ImputeItem.cImputedFish
                TStep = ImputeItem.cImputedTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * (FishTSCatchOrig(Fish, TStep) / FishTSCatchNew(Fish, TStep))
                    Next
                Next
            End If
        Next

        'zero out CWTs to be replaced then start the surrogate process that also allows for adding fisheries and time steps
        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 2 Or ImputeItem.cBaseType = 3 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
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
                        'If CWTAll(STk, Age, Fish, TStep) > 0 Then
                        FishTSCatchNew(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                        Print(15, STk & "," & Age & "," & Fish & "," & TStep & "," & CWTAll(STk, Age, Fish, TStep) & vbCrLf)
                        'End If
                    Next
                Next
            Next
        Next STk

        FileClose(15)
    End Sub
    '########## Begin JC Update; 9/25/2015 ##########
    'subroutine to compute CWTs for fisheries using old BPERs as surrogates after cohort sizes are available or change with each iteration
    Sub ImputeOldBPERs()

        Dim SFish As Integer
        Dim STStep As Integer
        Dim FishTSCatchOrig(NumFish, NumSteps) As Double
        Dim FishTSCatchNew(NumFish, NumSteps) As Double

        'store original fishery catches
        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish - 1
                    For TStep = 1 To NumSteps
                        FishTSCatchOrig(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                    Next
                Next
            Next
        Next

        'Estimate CWT catch for fisheries using Old BPER as surrogate

        For Each ImputeItem In ImputeList
            If ImputeItem.cBaseType = 9 Then
                FishNum = ImputeItem.cImputedFish
                TStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                FishYear = 2010
                TermStat = TermFlag(FishNum, TStep)
                'Estimate recoveries using old BPERs
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CompLegProp()
                        CWTAll(STk, Age, FishNum, TStep) = CohortAll(STk, Age, TermStat, TStep) * SurrogateFishBP_ER(STk, Age, SFish, STStep) * LegalProp
                        FishTSCatchNew(FishNum, TStep) += CWTAll(STk, Age, FishNum, TStep)
                    Next
                Next
                'Rescale fishery to original size (if original = 0, scale to a total CWT recovery of 1)
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        If FishTSCatchOrig(FishNum, TStep) > 0 Then
                            CWTAll(STk, Age, FishNum, TStep) = CWTAll(STk, Age, FishNum, TStep) * (FishTSCatchOrig(FishNum, TStep) / FishTSCatchNew(FishNum, TStep))
                        Else
                            CWTAll(STk, Age, FishNum, TStep) = CWTAll(STk, Age, FishNum, TStep) / FishTSCatchNew(FishNum, TStep)
                        End If
                    Next
                Next
            End If
        Next
       
        'For debugging purposes, recompute FishTSCatchNew and output CWTAll where > 0
        Dim Imputecatch As String
        Imputecatch = filepath & "ImputeCatch.txt"
        FileOpen(15, Imputecatch, OpenMode.Output)
        Print(15, "Stock,  Age, Fish, TStep, CWT" & vbCrLf)
        ReDim FishTSCatchNew(NumFish, NumSteps)
        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish
                    For TStep = 1 To NumSteps
                        'If CWTAll(STk, Age, Fish, TStep) > 0 Then
                        FishTSCatchNew(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                        Print(15, STk & "," & Age & "," & Fish & "," & TStep & "," & CWTAll(STk, Age, Fish, TStep) & vbCrLf)
                        'End If
                    Next
                Next
            Next
        Next STk
        FileClose(15)

    End Sub '########## End JC Update; 9/25/2015 ##########

    'Function findRecord(ByVal b As CWTData) As Boolean
    Function findRecord(ByVal b As CWTData) As Boolean
        If (b.cFish = ImputeItem.cSourceFish And b.cTStep = ImputeItem.cSourceTStep) Then

            Return True
        Else
            Return False
        End If
    End Function
    Function FindDelRec(ByVal d As CWTData) As Boolean 'find records to delete 
        n = n + 1
        If (d.cFish = ImputeItem.cImputedFish And d.cTStep = ImputeItem.cImputedTStep) Then
            Return True
        Else
            Return False
        End If

    End Function
End Module
