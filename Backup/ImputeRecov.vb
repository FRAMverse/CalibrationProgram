Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Imports System.Math
Module ImputeRecov
    ' Impute recoveries
    Sub Impute()

        ' ''######################################### Begin JC 8/14/15 update #########################################
        ' ''Code to reassign CWTs to new time step based on records in ImputeRecoveries table with 'Type' < 0

        ' ''

        Dim IFish As Integer
        Dim ITStep As Integer
        Dim SFish As Integer
        Dim STStep As Integer
        Dim Imputecatch As String
        Dim FishTSCatchOrig(NumFish, NumSteps) As Double
        Dim FishTSCatchNew(NumFish, NumSteps) As Double
        Dim CWTAllOrig(NumStk, MaxAge, NumFish, NumSteps) As Double
        Imputecatch = "C:\data\calibration\07Qbasic\imputecatch"
        FileOpen(15, Imputecatch, OpenMode.Output)
        Print(15, "Stock,  Age, Fish, TStep, CWT" & vbCrLf)

        ''For Each ImputeItem In ImputeList
        ''    If ImputeItem.cBaseType < 0 Then

        ''        IFish = ImputeItem.cImputedFish
        ''        ITStep = ImputeItem.cImputedTStep
        ''        SFish = ImputeItem.cSourceFish
        ''        STStep = ImputeItem.cSourceTStep

        ''            Case -1, -99
        ''        For STk = 1 To NumStk
        ''            For Age = 2 To MaxAge
        ''                CWTAll(STk, Age, SFish, STStep) = CWTAll(STk, Age, SFish, STStep) + CWTAll(STk, Age, IFish, ITStep)
        ''                CWTAll(STk, Age, IFish, ITStep) = 0
        ''            Next
        ''        Next

        ''        For STk = 1 To NumStk
        ''            For Age = 2 To MaxAge
        ''                CWTAll(STk, Age, SFish, STStep) = CWTAll(STk, Age, SFish, STStep) + CWTAll(STk, Age, IFish, ITStep)
        ''                CWTAll(STk, Age, IFish, ITStep) = 0
        ''            Next
        ''        Next

        ''    End If
        ''Next
        ' ''########################################## End JC 8/14/15 update ##########################################

        ''If OOBStatus = 1 Or Firstpass = True Then
        ''    BaseType = 0
        ''Else
        ''    BaseType = 1
        ''End If 'during the last baser period calibration OOB stocks had different imputed fisheries than Allstocks
        ''If BaseType = 0 Then
        ''    'this code needs to be reviewed to make sure it works with all the new BaseType options!!!! AHB 8/21/15
        ''    Dim newlist As New List(Of CWTData)
        ''    Dim newlist2 As New List(Of CWTData)
        ''    Dim imputerecoveries As CWTData
        ''    For Each ImputeItem In ImputeList
        ''        If BaseType = ImputeItem.cBaseType Then
        ''            'find all the records that meet criteria set in findRecord
        ''            sublist = CWTList.FindAll(AddressOf findRecord)
        ''            For Each RecordCWT In sublist
        ''                imputerecoveries.cCatch = Math.Round(RecordCWT.cCatch / 1000, 4)
        ''                imputerecoveries.cFish = ImputeItem.cImputedFish
        ''                imputerecoveries.cBY = RecordCWT.cBY
        ''                imputerecoveries.cStk = RecordCWT.cStk
        ''                imputerecoveries.cAge = RecordCWT.cAge
        ''                imputerecoveries.cStage = RecordCWT.cStage
        ''                imputerecoveries.cTStep = ImputeItem.cImputedTStep
        ''                imputerecoveries.cLookUp = RecordCWT.cLookUp
        ''                newlist.Add(imputerecoveries)
        ''            Next
        ''            'newlist.AddRange(sublist) 'add sublist to newlist for each row in ImputeItem
        ''        End If
        ''    Next

        ''    Dim Deletelist As New List(Of CWTData)
        ''    For Each ImputeItem In ImputeList
        ''        If BaseType = ImputeItem.cBaseType Then
        ''            'delete records that are going to be replaced with the source fishery
        ''            'Deletelist = CWTList.FindAll(AddressOf FindDelRec)
        ''            CWTList.RemoveAll(AddressOf FindDelRec)
        ''        End If
        ''    Next

        ''    CWTList.AddRange(newlist)
        ''    'For Each RecordCWT In CWTList
        ''    '    'Print(15, RecordCWT.cStk & "," & RecordCWT.cBY & "," & RecordCWT.cStage & "," & RecordCWT.cAge & "," & RecordCWT.cFish _
        ''    '    & "," & RecordCWT.cTStep & "," & RecordCWT.cCatch & vbCrLf)
        ''    'Next
        ''    'FileClose(15)
        ''Else 'impute for all stocks file using array
        ''    'Dim IFish As Integer
        ''    'Dim ITStep As Integer
        ''    'Dim SFish As Integer
        ''    'Dim STStep As Integer
        ''    For Each ImputeItem In ImputeList
        ''        If BaseType = Abs(ImputeItem.cBaseType) Then 'JC 8/14/15 update: added Abs() to accomodate new BaseType codes
        ''            IFish = ImputeItem.cImputedFish
        ''            ITStep = ImputeItem.cImputedTStep
        ''            SFish = ImputeItem.cSourceFish
        ''            STStep = ImputeItem.cSourceTStep
        ''            For STk = 1 To NumStk
        ''                For Age = 2 To MaxAge

        ''                    CWTAll(STk, Age, IFish, ITStep) = Math.Round(CWTAll(STk, Age, SFish, STStep) / 1000, 4, MidpointRounding.AwayFromZero)
        ''                    'CWTAll(STk, Age, IFish, ITStep) = CWTAll(STk, Age, SFish, STStep) / 1000
        ''                    'Print(15, STk & "," & "BY" & "," & "Stage" & "," & Age & "," & IFish _
        ''                    ' & "," & ITStep & "," & CWTAll(STk, Age, IFish, ITStep) & vbCrLf)
        ''                Next
        ''            Next
        ''        End If
        ''    Next
        ''End if
        '*******************************************************************************************************************************
        'AHB 8/21/2015
        '
        'Type 0 for surrogates needed for OOB stocks
        'Type 1 combines recoveries for a fishery from different time steps. This is done before surrogates are processed. Source time step is deleted
        'Type 11 combines recoveries for a fishery from different time steps. Source time step is included in combo. Fisheries are rescaled to original size.
        'Type 2 flag reassigns or combines CWTs to new time step or fishery based on records in ImputeRecoveries table'Type'= original concept
        'type 3 zero out fishery and time step

        'during the last baser period calibration OOB stocks had different imputed fisheries than Allstocks
        
        If OOBStatus = 1 Or Firstpass = True Then
            BaseType = 0
        Else
            BaseType = 1
        End If 'during the last baser period calibration OOB stocks had different imputed fisheries than Allstocks

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

        For Each ImputeItem In ImputeList
            'first add CWTs from designated time steps & delete source value
            If ImputeItem.cBaseType = 1 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) += CWTAll(STk, Age, SFish, STStep)
                        CWTAll(STk, Age, SFish, STStep) = 0
                    Next
                Next
            End If
        Next
        For Each ImputeItem In ImputeList
            'first add CWTs from designated time steps & maintain source value
            'this necessitates scaling to original number of recoverie in order to compute model stock ppn
            If ImputeItem.cBaseType = 11 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) += CWTAllOrig(STk, Age, SFish, STStep)
                    Next
                Next
            End If
        Next

        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish - 1
                    For TStep = 1 To NumSteps
                        FishTSCatchNew(Fish, TStep) += CWTAll(STk, Age, Fish, TStep)
                    Next
                Next
            Next
        Next
        For Each ImputeItem In ImputeList
            'scale summed catches back to original catch
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




        For Each ImputeItem In ImputeList
            'zero out CWTs to be replaced
            If ImputeItem.cBaseType = 2 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) = 0
                    Next
                Next
            End If
        Next


        For Each ImputeItem In ImputeList
            'then start the surrogate process that also allows for adding fisheries and time steps
            If ImputeItem.cBaseType = 2 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep
                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) += Math.Round(CWTAll(STk, Age, SFish, STStep) / 1000, 4, MidpointRounding.AwayFromZero)
                    Next
                Next
            End If
        Next

        For Each ImputeItem In ImputeList
            'finally, zero out fisheries that are slated to be eliminated
            If ImputeItem.cBaseType = 3 Then
                IFish = ImputeItem.cImputedFish
                ITStep = ImputeItem.cImputedTStep
                SFish = ImputeItem.cSourceFish
                STStep = ImputeItem.cSourceTStep

                For STk = 1 To NumStk
                    For Age = 2 To MaxAge
                        CWTAll(STk, Age, IFish, ITStep) = 0
                    Next
                Next
            End If
        Next

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
