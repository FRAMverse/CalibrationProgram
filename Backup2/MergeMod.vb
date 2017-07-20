Imports System.Collections.Generic
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Module MergeMod
    Public Sub Merge()
        'PROGRAM TO MERGE SEVERAL NONBASE PERIOD SIMULATED DATASETS
        'AND RESPLIT TREATY/NONTREATY PROPORTIONS

        'Update CalibrationSupport File

        '*** Note that stock 2 Nooksack Early is hardwired not to resplit the
        '    Nooksack/Samish T/NT because of concerns about the differential
        '    impacts of the fleets.

        'Percent Treaty section below needs to be updated after each pass of calibration
        'with output from ChTrty00.bas, Kristin Nason 1/00
        '**********************************************************************************
        'Rule: Merge data with same stock number

        'ReDim PropTreaty(NumFish + 2, NumSteps + 1)
        ReDim MonthNames(NumSteps + 1)
        ReDim FishName(NumFish + 2)
        ReDim Rmax(200)
        ReDim MergedCWT(NumStk, MaxAge, NumFish, NumSteps)


        ''read in weights and other info from database
        Dim objConnection As New OleDbConnection(SConStr)
        Dim objdataadapter As New OleDbDataAdapter()
        Dim objdataset As New DataSet()
        Dim objdatacommand As New OleDbCommand
        Dim objdatacommand2 As New OleDbCommand

        Dim MergedCatch As String
        mergedcatch = "C:\data\calibration\07Qbasic\mergedcatch"
        FileOpen(14, MergedCatch, OpenMode.Output)
        Print(14, "Stock, Age, Fish, TStep, catch" & vbCrLf)



        objConnection.Open()


        
        'Set and/or read in weights
        SetWeight.ShowDialog()

        'Compute weighted stock recoveries

        For Each TotLandedItem In TotalLandedList
            BY = TotLandedItem.bBY
            Stage = TotLandedItem.bStage
            STk = TotLandedItem.bStk
            BYSTKCatch = TotLandedItem.bTotalCatch
            'find the brood year with the greatest number of recoveries for each stock
            Dim Rmax As Double = 0
            SumCWT = 0
            For Each TotLandedItem2 In TotalLandedList
                If TotLandedItem2.bStk = STk Then
                    SumCWT = SumCWT + TotLandedItem2.bTotalCatch
                    If TotLandedItem2.bTotalCatch > Rmax Then
                        Rmax = TotLandedItem2.bTotalCatch
                    End If
                End If
            Next
            'find the weight and flag info
            objdatacommand = New OleDbCommand("SELECT Weight FROM Weighting Where StockID = " & STk & "and BroodYear = " & BY & "and Stage = " & Stage, objConnection)
            Dim ds As OleDbDataReader = objdatacommand.ExecuteReader()
            If ds.HasRows Then
                While ds.Read
                    BYWeight = ds.Item("Weight")
                End While
            Else
                MsgBox("Stock " & STk & " Brood Year " & BY & " is not included in the Weighting table.")
            End If

            objdatacommand2 = New OleDbCommand("SELECT Flag FROM Weighting Where StockID = " & STk & "and BroodYear = " & BY & "and Stage = " & Stage, objConnection)
            Dim dv As OleDbDataReader = objdatacommand2.ExecuteReader()
            If dv.HasRows Then
                While dv.Read
                    BYFlag = dv.Item("Flag")
                End While
            Else
                MsgBox("Stock " & STk & " Brood Year " & BY & " is not included in the Weighting table.")
            End If

            If BYWeight = 1 Then
                Expander = BYWeight
            Else
                Expander = ((SumCWT - BYSTKCatch) / BYSTKCatch) * BYWeight
            End If


            For Each imputerecoveriesmain In ImputeListMain
                If imputerecoveriesmain.cStk = STk And imputerecoveriesmain.cBY = BY And imputerecoveriesmain.cStage = Stage Then
                    Age = imputerecoveriesmain.cAge
                    Fish = imputerecoveriesmain.cFish
                    TStep = imputerecoveriesmain.cTStep
                    If BYFlag = 0 Then
                        MergedCWT(STk, Age, Fish, TStep) = MergedCWT(STk, Age, Fish, TStep) + imputerecoveriesmain.cCatch * Rmax / BYSTKCatch

                    ElseIf BYFlag = 1 Then

                        MergedCWT(STk, Age, Fish, TStep) = MergedCWT(STk, Age, Fish, TStep) + imputerecoveriesmain.cCatch * Expander

                    Else 'BYflag = 2
                        MergedCWT(STk, Age, Fish, TStep) = MergedCWT(STk, Age, Fish, TStep) + imputerecoveriesmain.cCatch * BYWeight

                    End If
                End If
            Next
        Next

        'Not redistributing catch to treaty/non-treaty. This results in using the same rules as in all stocks run for surrogate fisheries,
        'where the recoveries of a NT/T pair are reported under NT and then split 50/50 (due to surrogate assignments and modeling 50% of total catch for 
        'each fishery in a pair)
        'REDISTRIBUTE CATCH IN TREATY/NONTREATY FISHERIES



        '    For Fish = 1 To NumFish
        '        If FishFlag(Fish) = 1 Then
        '            If (Stocknumber = 2 Or Stocknumber = 3) And Fish = 39 Then
        '                MsgBox("Not redistributing Nooksack/Samish net T/NT catch for Nooksack Early")
        '            Else
        '                For Age = 2 To MaxAge
        '                    For TStep = 1 To NumSteps

        '                        'ADD RECOVERIES IN SET OF FISHERIES

        '                        TotalCatch = OOBCatchW(Stocknumber, Age, Fish, TStep) + OOBCatchW(Stocknumber, Age, Fish + 1, TStep)
        '                        OOBCatchW(Stocknumber, Age, Fish, TStep) = TotalCatch * (1 - PropTreaty(Fish, TStep))
        '                        OOBCatchW(Stocknumber, Age, Fish + 1, TStep) = TotalCatch * PropTreaty(Fish, TStep)
        '                    Next TStep
        '                Next Age
        '            End If
        '        End If
        '    Next Fish
        'Dim TotalCatch As Double = 0
        'For STk = 1 To NumStk

        '    For Fish = 1 To NumFish

        '        If FishFlag(Fish) = 1 Then
        '            If (STk = 2 Or STk = 3) And Fish = 39 Then
        '                MsgBox("Not redistributing Nooksack/Samish net T/NT catch for Nooksack Early")
        '            Else
        '                For Age = 2 To MaxAge
        '                    For TStep = 1 To NumSteps
        '                        TotalCatch = MergedCWT(STk, Age, Fish, TStep) + MergedCWT(STk, Age, Fish + 1, TStep)
        '                        MergedCWT(STk, Age, Fish, TStep) = TotalCatch * (1 - PropTNet(Fish, TStep))
        '                        MergedCWT(STk, Age, Fish + 1, TStep) = TotalCatch * PropTNet(Fish, TStep)
        '                    Next
        '                Next
        '            End If
        '        End If

        '    Next
        'Next



        For STk = 1 To NumStk
            For Age = 2 To MaxAge
                For Fish = 1 To NumFish
                    For TStep = 1 To NumSteps
                        If MergedCWT(STk, Age, Fish, TStep) <> 0 Then
                            Print(14, STk & "," & Age & "," & Fish & "," & TStep & "," & MergedCWT(STk, Age, Fish, TStep) & vbCrLf)
                        End If
                    Next
                Next
            Next
        Next
        objConnection.Close()
        Firstpass = False
        FileClose()

    End Sub

End Module
