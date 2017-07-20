Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Imports System.IO.File
Public Class StartForm
    Public TransDb As New OleDbConnection


    Private Sub StartForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ToolTip1.SetToolTip(RadioButton1, "Only summarizes CWT data for the OOB stocks")
        ToolTip1.SetToolTip(RadioButton2, "First runs OOB stocks and merges results with all stocks then re-runs calcualtions for all stocks")
        ToolTip1.SetToolTip(RadioButton3, "For a new base period that contains all stocks")



    End Sub




    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then 'OOB only
            OOBStatus = 1
            Firstpass = True
            NoExpansions = False
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then 'First OOB then AllStocks
            OOBStatus = 2
            Firstpass = True
            NoExpansions = False
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then 'AllStocks Only
            OOBStatus = 3
            Firstpass = False
            NoExpansions = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunBtn.Click
        mainform.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitBtb.Click
        End
    End Sub

    Private Sub ExportBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportBtn.Click
        Dim iresult As String
        Dim dbConnection9 As New OleDbConnection



        TransferBPName = "NewCalibrationBasePeriodTransfer.mdb"
       
        Try
            FVSdatabasepath = My.Computer.FileSystem.GetFileInfo(FVSdatabasename).DirectoryName
            TransferBPLongName = FVSdatabasepath & "\" & TransferBPName
        Catch ex As Exception
            MsgBox("Please select the TransferFile named 'NewCalibrationBasePeriodTransfer'.")
            OpenFileDialog2.Title = "Select the Transfer File"
            OpenFileDialog2.Filter = "DataBase Files(*.mdb;*.accdb)|*.MDB;*.ACCDB"

            OpenFileDialog2.CheckFileExists = True
            iresult = OpenFileDialog2.ShowDialog()


            'Make sure the user did not click the Cancel button And 
            '    specified a file name for the file to be created.  
            If iresult <> Windows.Forms.DialogResult.Cancel And _
                        OpenFileDialog2.FileName.Length <> 0 Then
                TransferBPLongName = OpenFileDialog2.FileName
                'FVSdatabasepath = Path.GetDirectoryName(TransferBPLongName)
                'TransferBPName = System.IO.Path.GetFileName(OpenFileDialog2.FileName)
            End If
        End Try

        ' If Exists(FVSdatabasepath & "\" & TransferBPName) Then


        If TransferBPName = "" Then
            TransferBPName = System.IO.Path.GetFileName(OpenFileDialog2.FileName)
        End If
        If FVSdatabasepath = "" Then
            FVSdatabasepath = My.Computer.FileSystem.GetFileInfo(OpenFileDialog2.FileName).DirectoryName
        End If

        Me.Visible = False
        FVS_ModelRunSelection.ShowDialog()

        Me.Cursor = Cursors.WaitCursor
        '- Create Copy of Transfer Database File

NewName:

        TransferBPLongName = ""
        MdbSaveFileDialog.Filter = "DataBase Files(*.mdb;*.accdb)|*.MDB;*.ACCDB"
        MdbSaveFileDialog.FilterIndex = 1
        MdbSaveFileDialog.FileName = "NewCalibrationBasePeriodTransfer.Mdb"
        MdbSaveFileDialog.RestoreDirectory = True
        MsgBox("Please enter the name of the new transfer database for saving.")
        If MdbSaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                TransferBPLongName = MdbSaveFileDialog.FileName
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
        End If



        If TransferBPLongName = "" Then Exit Sub
        If TransferBPLongName = "NewCalibrationBasePeriodTransfer.Mdb" Then
            MsgBox("The file 'NewCalibrationBasePeriodTransfer.Mdb' is Reserved" & vbCrLf & _
                   "Please Choose Different Name for Transfer DataBase" & vbCrLf & _
                   "Prevents Corruption of Database Structure!", MsgBoxStyle.OkOnly)
            GoTo NewName
        End If

        
        If Exists(TransferBPLongName) Then Delete(TransferBPLongName)
        Try
            File.Copy(FVSdatabasepath & "\" & TransferBPName, TransferBPLongName, True)
        Catch
            MsgBox("Can't find NewCalibrationBasePeriodTransfer.MDB file in the same directory as the calibration database!" & vbCrLf & "Please transfer to same directory.", MsgBoxStyle.OkOnly)
            Exit Sub
        End Try

        '- TransferDB Connection String
        TransDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TransferBPLongName

        Me.Cursor = Cursors.WaitCursor
        Call TransferBasePeriodTables()
        Me.Cursor = Cursors.Default
        Me.Visible = True
        'Else
        '
        'End If

        Me.Cursor = Cursors.Default
    End Sub
    Private Sub TransferBasePeriodTables()
        Dim CmdStr As String
        Dim RecNum, NumRecs As Integer

       

        'Transfer BaseID Record
        'Retrieve info from SummaryInfo table
        CmdStr = "SELECT * FROM SummaryInfo WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim BIDcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim BaseIDDA As New System.Data.OleDb.OleDbDataAdapter
        BaseIDDA.SelectCommand = BIDcm
        Dim BIDcb As New OleDb.OleDbCommandBuilder
        BIDcb = New OleDb.OleDbCommandBuilder(BaseIDDA)
        If TransferDataSet.Tables.Contains("BaseID") Then
            TransferDataSet.Tables("BaseID").Clear()
        End If
        BaseIDDA.Fill(TransferDataSet, "SummaryInfo")
        BaseYear = TransferDataSet.Tables("SummaryInfo").Rows(0)(14)
        NumFish = TransferDataSet.Tables("SummaryInfo").Rows(0)(2) - 1
        NumSteps = TransferDataSet.Tables("SummaryInfo").Rows(0)(3) + 1
        'CalibrationDB.Close()
        'Retrieve Info from AEQ Table

        CmdStr = "SELECT * FROM AEQ WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim AEQCohortcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim AEQCohortIDDA As New System.Data.OleDb.OleDbDataAdapter
        AEQCohortIDDA.SelectCommand = AEQCohortcm
        Dim AEQCohortcb As New OleDb.OleDbCommandBuilder
        AEQCohortcb = New OleDb.OleDbCommandBuilder(AEQCohortIDDA)
        If TransferDataSet.Tables.Contains("AEQ") Then
            TransferDataSet.Tables("AEQ").Clear()
        End If
        AEQCohortIDDA.Fill(TransferDataSet, "AEQ")
        Dim NumAEQ As Integer
        NumAEQ = TransferDataSet.Tables("AEQ").Rows.Count




        'populate BaseID info on TrasferDB
        Dim BIDTrans As OleDb.OleDbTransaction
        Dim BID As New OleDbCommand
        TransDb.Open()
        BIDTrans = TransDb.BeginTransaction
        BID.Connection = TransDb
        BID.Transaction = BIDTrans
        RecNum = 0
        'BID.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
        BID.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
        "VALUES(" & SelectedBasePeriodID & "," & _
        Chr(34) & "New Chinook Calibration BasePeriod" & Chr(34) & "," & _
        Chr(34) & "CHINOOK" & Chr(34) & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(1) * 2 & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(2) - 1 & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(3) + 1 & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(4) - 1 & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(4) - 3 & "," & _
        TransferDataSet.Tables("SummaryInfo").Rows(0)(4) & "," & _
        Chr(35) & TransferDataSet.Tables("AEQ").Rows(0)(5) & Chr(35) & "," & _
        Chr(34) & " " & Chr(34) & "," & _
        Chr(34) & "1" & Chr(34) & "," & _
        Chr(34) & "1" & Chr(34) & "," & _
        Chr(34) & "1" & Chr(34) & ")"
        BID.ExecuteNonQuery()
        BIDTrans.Commit()
        TransDb.Close()



        'Base Cohort
        CmdStr = "SELECT * FROM BaseCohort WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim BaseCohortcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim BaseCohortIDDA As New System.Data.OleDb.OleDbDataAdapter
        BaseCohortIDDA.SelectCommand = BaseCohortcm
        Dim BaseCohortcb As New OleDb.OleDbCommandBuilder
        BaseCohortcb = New OleDb.OleDbCommandBuilder(BaseCohortIDDA)
        If TransferDataSet.Tables.Contains("BaseCohort") Then
            TransferDataSet.Tables("BaseCohort").Clear()
        End If
        BaseCohortIDDA.Fill(TransferDataSet, "BaseCohort")
        Dim NumBaseCohort As Integer
        NumBaseCohort = TransferDataSet.Tables("BaseCohort").Rows.Count

        Dim BaseCohortTrans As OleDb.OleDbTransaction
        Dim BaseCohort As New OleDbCommand
        TransDb.Open()
        BaseCohortTrans = TransDb.BeginTransaction
        BaseCohort.Connection = TransDb
        BaseCohort.Transaction = BaseCohortTrans
        NumRecs = TransferDataSet.Tables("BaseCohort").Rows.Count

        ' next two paragraphs result in the base cohort being doubled when marked and unmarked stock scalar = 1
        For RecNum = 0 To NumBaseCohort - 1
            BaseCohort.CommandText = "INSERT INTO BaseCohort (BasePeriodID,StockID,Age,BaseCohortSize) " & _
               "VALUES(" & TransferDataSet.Tables("BaseCohort").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("BaseCohort").Rows(RecNum)(1) * 2 & "," & _
                TransferDataSet.Tables("BaseCohort").Rows(RecNum)(2) & "," & _
               TransferDataSet.Tables("BaseCohort").Rows(RecNum)(3) & ")"
            BaseCohort.ExecuteNonQuery()
        Next

        For RecNum = 0 To NumBaseCohort - 1
            BaseCohort.CommandText = "INSERT INTO BaseCohort (BasePeriodID,StockID,Age,BaseCohortSize) " & _
               "VALUES(" & TransferDataSet.Tables("BaseCohort").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("BaseCohort").Rows(RecNum)(1) * 2 - 1 & "," & _
                TransferDataSet.Tables("BaseCohort").Rows(RecNum)(2) & "," & _
               TransferDataSet.Tables("BaseCohort").Rows(RecNum)(3) & ")"
            BaseCohort.ExecuteNonQuery()
        Next

        BaseCohortTrans.Commit()
        TransDb.Close()

        'AEQ
        Dim AEQTrans As OleDb.OleDbTransaction
        Dim AEQ As New OleDbCommand
        TransDb.Open()
        AEQTrans = TransDb.BeginTransaction
        AEQ.Connection = TransDb
        AEQ.Transaction = AEQTrans
        NumRecs = TransferDataSet.Tables("AEQ").Rows.Count

        For RecNum = 0 To NumAEQ - 1
            AEQ.CommandText = "INSERT INTO AEQ (BasePeriodID,StockID,Age,TimeStep, AEQ) " & _
               "VALUES(" & TransferDataSet.Tables("AEQ").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(1) * 2 - 1 & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(3) & "," & _
               TransferDataSet.Tables("AEQ").Rows(RecNum)(4) & ")"
            AEQ.ExecuteNonQuery()
        Next


        For RecNum = 0 To NumAEQ - 1
            AEQ.CommandText = "INSERT INTO AEQ (BasePeriodID,StockID,Age,TimeStep, AEQ) " & _
               "VALUES(" & TransferDataSet.Tables("AEQ").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(1) * 2 & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("AEQ").Rows(RecNum)(3) & "," & _
               TransferDataSet.Tables("AEQ").Rows(RecNum)(4) & ")"
            AEQ.ExecuteNonQuery()
        Next
        AEQTrans.Commit()

        TransDb.Close()


        'BaseExploitationRate
        CmdStr = "SELECT * FROM BaseExploitationRate WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim BaseExploitationRatecm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim BaseExploitationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
        BaseExploitationRateIDDA.SelectCommand = BaseExploitationRatecm
        Dim BaseExploitationRatecb As New OleDb.OleDbCommandBuilder
        BaseExploitationRatecb = New OleDb.OleDbCommandBuilder(BaseExploitationRateIDDA)
        If TransferDataSet.Tables.Contains("BaseExploitationRate") Then
            TransferDataSet.Tables("BaseExploitationRate").Clear()
        End If
        BaseExploitationRateIDDA.Fill(TransferDataSet, "BaseExploitationRate")
        'Dim NumBaseExploitationRate As Integer


        Dim BaseExploitationRateTrans As OleDb.OleDbTransaction
        Dim BaseExploitationRate As New OleDbCommand
        TransDb.Open()
        BaseExploitationRateTrans = TransDb.BeginTransaction
        BaseExploitationRate.Connection = TransDb
        BaseExploitationRate.Transaction = BaseExploitationRateTrans
        NumRecs = TransferDataSet.Tables("BaseExploitationRate").Rows.Count
        For RecNum = 0 To NumRecs - 1
            BaseExploitationRate.CommandText = "INSERT INTO BaseExploitationRate (BasePeriodID,StockID,Age,FisheryID,TimeStep,ExploitationRate,SublegalEncounterRate) " & _
               "VALUES(" & TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(1) * 2 & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(4) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(5) & "," & _
               TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(6) & ")"

            BaseExploitationRate.ExecuteNonQuery()
        Next
        For RecNum = 0 To NumRecs - 1
            BaseExploitationRate.CommandText = "INSERT INTO BaseExploitationRate (BasePeriodID,StockID,Age,FisheryID,TimeStep,ExploitationRate,SublegalEncounterRate) " & _
               "VALUES(" & TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(1) * 2 - 1 & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(4) & "," & _
                TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(5) & "," & _
               TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(6) & ")"

            BaseExploitationRate.ExecuteNonQuery()
        Next
        BaseExploitationRateTrans.Commit()
        TransDb.Close()

        '        'ChinookBaseEncounterAdjustment
        '        CmdStr = "SELECT * FROM ChinookBaseEncounterAdjustment"
        '        Dim ChinookBaseEncounterAdjustmentcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        '        Dim ChinookBaseEncounterAdjustmentIDDA As New System.Data.OleDb.OleDbDataAdapter
        '        ChinookBaseEncounterAdjustmentIDDA.SelectCommand = ChinookBaseEncounterAdjustmentcm
        '        Dim ChinookBaseEncounterAdjustmentcb As New OleDb.OleDbCommandBuilder
        '        ChinookBaseEncounterAdjustmentcb = New OleDb.OleDbCommandBuilder(ChinookBaseEncounterAdjustmentIDDA)
        '        If TransferDataSet.Tables.Contains("ChinookBaseEncounterAdjustment") Then
        '            TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Clear()
        '        End If
        '        ChinookBaseEncounterAdjustmentIDDA.Fill(TransferDataSet, "ChinookBaseEncounterAdjustment")
        '        Dim NumChinookBaseEncounterAdjustment As Integer
        '        NumChinookBaseEncounterAdjustment = TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows.Count

        '        Dim ChinookBaseEncounterAdjustmentTrans As OleDb.OleDbTransaction
        '        Dim ChinookBaseEncounterAdjustment As New OleDbCommand
        '        TransDB.Open()
        '        ChinookBaseEncounterAdjustmentTrans = TransDB.BeginTransaction
        '        ChinookBaseEncounterAdjustment.Connection = TransDB
        '        ChinookBaseEncounterAdjustment.Transaction = ChinookBaseEncounterAdjustmentTrans
        '        NumRecs = TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows.Count
        '        For RecNum = 0 To NumRecs - 1
        '            ChinookBaseEncounterAdjustment.CommandText = "INSERT INTO ChinookBaseEncounterAdjustment (FisheryID,Time1Adjustment,Time2Adjustment,Time3Adjustment,Time4Adjustment) " & _
        '               "VALUES(" & TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(0) & "," & _
        '                TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(1) & "," & _
        '                TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(2) & "," & _
        '                TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(3) & "," & _
        '               TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(4) & ")"

        '            ChinookBaseEncounterAdjustment.ExecuteNonQuery()
        '        Next
        '        ChinookBaseEncounterAdjustmentTrans.Commit()
        '        TransDB.Close()

        'ChinookBaseSizeLimit()
        ReDim SizeLimit(NumFish, NumSteps)
        CmdStr = "SELECT * FROM SizeLimits WHERE FishingYear = " & BaseYear & " ORDER BY FisheryID, TimeStep;"
        Dim ChinookBaseSizeLimitcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim ChinookBaseSizeLimitIDDA As New System.Data.OleDb.OleDbDataAdapter
        ChinookBaseSizeLimitIDDA.SelectCommand = ChinookBaseSizeLimitcm
        Dim ChinookBaseSizeLimitcb As New OleDb.OleDbCommandBuilder
        ChinookBaseSizeLimitcb = New OleDb.OleDbCommandBuilder(ChinookBaseSizeLimitIDDA)
        If TransferDataSet.Tables.Contains("ChinookBaseSizeLimit") Then
            TransferDataSet.Tables("ChinookBaseSizeLimit").Clear()
        End If
        ChinookBaseSizeLimitIDDA.Fill(TransferDataSet, "ChinookBaseSizeLimit")
        'Dim NumChinookBaseSizeLimit As Integer

        NumRecs = TransferDataSet.Tables("ChinookBaseSizeLimit").Rows.Count
        For RecNum = 0 To NumRecs - 1
            SizeLimit(TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(2), TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(3)) = TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(4)
        Next


        Dim ChinookBaseSizeLimitTrans As OleDb.OleDbTransaction
        Dim ChinookBaseSizeLimit As New OleDbCommand
        TransDb.Open()
        ChinookBaseSizeLimitTrans = TransDb.BeginTransaction
        ChinookBaseSizeLimit.Connection = TransDb
        ChinookBaseSizeLimit.Transaction = ChinookBaseSizeLimitTrans

        For Fish = 1 To NumFish
            ChinookBaseSizeLimit.CommandText = "INSERT INTO ChinookBaseSizeLimit (FisheryID,Time1SizeLimit,Time2SizeLimit,Time3SizeLimit,Time4SizeLimit) " & _
                    "VALUES(" & Fish & "," & _
                            SizeLimit(Fish, 1) & "," & _
                            SizeLimit(Fish, 2) & "," & _
                            SizeLimit(Fish, 3) & "," & _
                            SizeLimit(Fish, 4) & ")"
            ChinookBaseSizeLimit.ExecuteNonQuery()
        Next Fish

        ChinookBaseSizeLimitTrans.Commit()
        TransDb.Close()

        'EncounterRateAdjustment
        CmdStr = "SELECT * FROM EncounterRateAdjustment WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim EncounterRateAdjustmentcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim EncounterRateAdjustmentIDDA As New System.Data.OleDb.OleDbDataAdapter
        EncounterRateAdjustmentIDDA.SelectCommand = EncounterRateAdjustmentcm
        Dim EncounterRateAdjustmentcb As New OleDb.OleDbCommandBuilder
        EncounterRateAdjustmentcb = New OleDb.OleDbCommandBuilder(EncounterRateAdjustmentIDDA)
        If TransferDataSet.Tables.Contains("EncounterRateAdjustment") Then
            TransferDataSet.Tables("EncounterRateAdjustment").Clear()
        End If
        EncounterRateAdjustmentIDDA.Fill(TransferDataSet, "EncounterRateAdjustment")
        Dim NumEncounterRateAdjustment As Integer
        NumEncounterRateAdjustment = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count

        Dim EncounterRateAdjustmentTrans As OleDb.OleDbTransaction
        Dim EncounterRateAdjustment As New OleDbCommand
        TransDb.Open()
        EncounterRateAdjustmentTrans = TransDb.BeginTransaction
        EncounterRateAdjustment.Connection = TransDb
        EncounterRateAdjustment.Transaction = EncounterRateAdjustmentTrans
        NumRecs = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count
        For RecNum = 0 To NumRecs - 1
            EncounterRateAdjustment.CommandText = "INSERT INTO EncounterRateAdjustment (BasePeriodID,Age,FisheryID,TimeStep,EncounterRateAdjustment) " & _
               "VALUES(" & TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(3) & "," & _
            TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(4) & "," & _
               TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(5) & ")"

            EncounterRateAdjustment.ExecuteNonQuery()
        Next
        EncounterRateAdjustmentTrans.Commit()
        TransDb.Close()



        'Fishery
        CmdStr = "SELECT * FROM Fishery"
        Dim Fisherycm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim FisheryIDDA As New System.Data.OleDb.OleDbDataAdapter
        FisheryIDDA.SelectCommand = Fisherycm
        Dim Fisherycb As New OleDb.OleDbCommandBuilder
        Fisherycb = New OleDb.OleDbCommandBuilder(FisheryIDDA)
        If TransferDataSet.Tables.Contains("Fishery") Then
            TransferDataSet.Tables("Fishery").Clear()
        End If
        FisheryIDDA.Fill(TransferDataSet, "Fishery")
        Dim NumFishery As Integer


        Dim FisheryTrans As OleDb.OleDbTransaction
        Dim Fishery As New OleDbCommand
        TransDb.Open()
        FisheryTrans = TransDb.BeginTransaction
        Fishery.Connection = TransDb
        Fishery.Transaction = FisheryTrans
        NumRecs = TransferDataSet.Tables("Fishery").Rows.Count
        For RecNum = 0 To NumRecs - 1
            Fishery.CommandText = "INSERT INTO Fishery (Species,VersionNumber,FisheryID,FisheryName,FisheryTitle) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("Fishery").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Fishery").Rows(RecNum)(2) & "," & _
                Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(3) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(4) & Chr(34) & ")"

            Fishery.ExecuteNonQuery()
        Next
        FisheryTrans.Commit()
        TransDb.Close()

        'FisheryModelStockProportion
        ' ''CmdStr = "SELECT * FROM FisheryModelStockProportion WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        ' ''Dim FisheryModelStockProportioncm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        ' ''Dim FisheryModelStockProportionIDDA As New System.Data.OleDb.OleDbDataAdapter
        ' ''FisheryModelStockProportionIDDA.SelectCommand = FisheryModelStockProportioncm
        ' ''Dim FisheryModelStockProportioncb As New OleDb.OleDbCommandBuilder
        ' ''FisheryModelStockProportioncb = New OleDb.OleDbCommandBuilder(FisheryModelStockProportionIDDA)
        ' ''If TransferDataSet.Tables.Contains("FisheryModelStockProportion") Then
        ' ''    TransferDataSet.Tables("FisheryModelStockProportion").Clear()
        ' ''End If
        ' ''FisheryModelStockProportionIDDA.Fill(TransferDataSet, "FisheryModelStockProportion")
        ' ''Dim NumFisheryModelStockProportion As Integer


        ' ''Dim FisheryModelStockProportionTrans As OleDb.OleDbTransaction
        ' ''Dim FisheryModelStockProportion As New OleDbCommand
        ' ''TransDb.Open()
        ' ''FisheryModelStockProportionTrans = TransDb.BeginTransaction
        ' ''FisheryModelStockProportion.Connection = TransDb
        ' ''FisheryModelStockProportion.Transaction = FisheryModelStockProportionTrans
        ' ''NumRecs = TransferDataSet.Tables("FisheryModelStockProportion").Rows.Count
        ' ''For RecNum = 0 To NumRecs - 1
        ' ''    FisheryModelStockProportion.CommandText = "INSERT INTO FisheryModelStockProportion (BasePeriodID,FisheryID,ModelStockProportion) " & _
        ' ''       "VALUES(" & TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(0) & "," & _
        ' ''        TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(1) & "," & _
        ' ''       TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(2) & ")"

        ' ''    FisheryModelStockProportion.ExecuteNonQuery()
        ' ''Next
        ' ''FisheryModelStockProportionTrans.Commit()
        ' ''TransDb.Close()

        'Growth
        CmdStr = "SELECT * FROM Growth WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim Growthcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim GrowthIDDA As New System.Data.OleDb.OleDbDataAdapter
        GrowthIDDA.SelectCommand = Growthcm
        Dim Growthcb As New OleDb.OleDbCommandBuilder
        Growthcb = New OleDb.OleDbCommandBuilder(GrowthIDDA)
        If TransferDataSet.Tables.Contains("Growth") Then
            TransferDataSet.Tables("Growth").Clear()
        End If
        GrowthIDDA.Fill(TransferDataSet, "Growth")
        Dim NumGrowth As Integer
        NumGrowth = TransferDataSet.Tables("Growth").Rows.Count

        Dim GrowthTrans As OleDb.OleDbTransaction
        Dim Growth As New OleDbCommand
        TransDb.Open()
        GrowthTrans = TransDb.BeginTransaction
        Growth.Connection = TransDb
        Growth.Transaction = GrowthTrans
        NumRecs = TransferDataSet.Tables("Growth").Rows.Count
        For RecNum = 0 To NumRecs - 1
            Growth.CommandText = "INSERT INTO Growth (BasePeriodID,StockID,LImmature,KImmature,TImmature,CV2Immature,CV3Immature,CV4Immature,CV5Immature,LMature,KMature,TMature,CV2Mature,CV3Mature,CV4Mature,CV5Mature) " & _
               "VALUES(" & TransferDataSet.Tables("Growth").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Growth").Rows(RecNum)(2) * 2 - 1 & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(3) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(4) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(5) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(6) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(7) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(8) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(9) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(10) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(11) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(12) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(13) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(14) & "," & _
            TransferDataSet.Tables("Growth").Rows(RecNum)(15) & "," & _
               TransferDataSet.Tables("Growth").Rows(RecNum)(16) & ")"

            Growth.ExecuteNonQuery()
        Next
        For RecNum = 0 To NumRecs - 1
            Growth.CommandText = "INSERT INTO Growth (BasePeriodID,StockID,LImmature,KImmature,TImmature,CV2Immature,CV3Immature,CV4Immature,CV5Immature,LMature,KMature,TMature,CV2Mature,CV3Mature,CV4Mature,CV5Mature) " & _
               "VALUES(" & TransferDataSet.Tables("Growth").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Growth").Rows(RecNum)(2) * 2 & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(3) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(4) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(5) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(6) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(7) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(8) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(9) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(10) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(11) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(12) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(13) & "," & _
             TransferDataSet.Tables("Growth").Rows(RecNum)(14) & "," & _
            TransferDataSet.Tables("Growth").Rows(RecNum)(15) & "," & _
               TransferDataSet.Tables("Growth").Rows(RecNum)(16) & ")"

            Growth.ExecuteNonQuery()
        Next
        GrowthTrans.Commit()
        TransDb.Close()



        'IncidentalRate
        CmdStr = "SELECT * FROM IncidentalRate WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim IncidentalRatecm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim IncidentalRateIDDA As New System.Data.OleDb.OleDbDataAdapter
        IncidentalRateIDDA.SelectCommand = IncidentalRatecm
        Dim IncidentalRatecb As New OleDb.OleDbCommandBuilder
        IncidentalRatecb = New OleDb.OleDbCommandBuilder(IncidentalRateIDDA)
        If TransferDataSet.Tables.Contains("IncidentalRate") Then
            TransferDataSet.Tables("IncidentalRate").Clear()
        End If
        IncidentalRateIDDA.Fill(TransferDataSet, "IncidentalRate")
        Dim NumIncidentalRate As Integer
        NumIncidentalRate = TransferDataSet.Tables("IncidentalRate").Rows.Count

        Dim IncidentalRateTrans As OleDb.OleDbTransaction
        Dim IncidentalRate As New OleDbCommand
        TransDb.Open()
        IncidentalRateTrans = TransDb.BeginTransaction
        IncidentalRate.Connection = TransDb
        IncidentalRate.Transaction = IncidentalRateTrans
        NumRecs = TransferDataSet.Tables("IncidentalRate").Rows.Count
        For RecNum = 0 To NumRecs - 1
            IncidentalRate.CommandText = "INSERT INTO IncidentalRate (BasePeriodID,FisheryID,TimeStep,IncidentalRate) " & _
               "VALUES(" & TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(3) & ")"

            IncidentalRate.ExecuteNonQuery()
        Next
        IncidentalRateTrans.Commit()
        TransDb.Close()

        'MaturationRate
        CmdStr = "SELECT * FROM MaturationRate WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim MaturationRatecm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim MaturationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
        MaturationRateIDDA.SelectCommand = MaturationRatecm
        Dim MaturationRatecb As New OleDb.OleDbCommandBuilder
        MaturationRatecb = New OleDb.OleDbCommandBuilder(MaturationRateIDDA)
        If TransferDataSet.Tables.Contains("MaturationRate") Then
            TransferDataSet.Tables("MaturationRate").Clear()
        End If
        MaturationRateIDDA.Fill(TransferDataSet, "MaturationRate")
        Dim NumMaturationRate As Integer
        NumMaturationRate = TransferDataSet.Tables("MaturationRate").Rows.Count

        Dim MaturationRateTrans As OleDb.OleDbTransaction
        Dim MaturationRate As New OleDbCommand
        TransDb.Open()
        MaturationRateTrans = TransDb.BeginTransaction
        MaturationRate.Connection = TransDb
        MaturationRate.Transaction = MaturationRateTrans
        NumRecs = TransferDataSet.Tables("MaturationRate").Rows.Count
        For RecNum = 0 To NumRecs - 1
            MaturationRate.CommandText = "INSERT INTO MaturationRate (BasePeriodID,StockID,Age,TimeStep,MaturationRate) " & _
               "VALUES(" & TransferDataSet.Tables("MaturationRate").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(1) * 2 & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(4) & ")"

            MaturationRate.ExecuteNonQuery()
        Next
        For RecNum = 0 To NumRecs - 1
            MaturationRate.CommandText = "INSERT INTO MaturationRate (BasePeriodID,StockID,Age,TimeStep,MaturationRate) " & _
               "VALUES(" & TransferDataSet.Tables("MaturationRate").Rows(RecNum)(0) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(1) * 2 - 1 & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("MaturationRate").Rows(RecNum)(4) & ")"

            MaturationRate.ExecuteNonQuery()
        Next
        MaturationRateTrans.Commit()
        TransDb.Close()

        'NaturalMortality
        CmdStr = "SELECT * FROM NaturalMortality WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim NaturalMortalitycm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim NaturalMortalityIDDA As New System.Data.OleDb.OleDbDataAdapter
        NaturalMortalityIDDA.SelectCommand = NaturalMortalitycm
        Dim NaturalMortalitycb As New OleDb.OleDbCommandBuilder
        NaturalMortalitycb = New OleDb.OleDbCommandBuilder(NaturalMortalityIDDA)
        If TransferDataSet.Tables.Contains("NaturalMortality") Then
            TransferDataSet.Tables("NaturalMortality").Clear()
        End If
        NaturalMortalityIDDA.Fill(TransferDataSet, "NaturalMortality")
        Dim NumNaturalMortality As Integer
        NumNaturalMortality = TransferDataSet.Tables("NaturalMortality").Rows.Count

        Dim NaturalMortalityTrans As OleDb.OleDbTransaction
        Dim NaturalMortality As New OleDbCommand
        TransDb.Open()
        NaturalMortalityTrans = TransDb.BeginTransaction
        NaturalMortality.Connection = TransDb
        NaturalMortality.Transaction = NaturalMortalityTrans
        NumRecs = TransferDataSet.Tables("NaturalMortality").Rows.Count
        For RecNum = 0 To NumRecs - 1
            NaturalMortality.CommandText = "INSERT INTO NaturalMortality (BasePeriodID,Age,TimeStep,NaturalMortalityRate) " & _
               "VALUES(" & TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(4) & ")"

            NaturalMortality.ExecuteNonQuery()
        Next
        NaturalMortalityTrans.Commit()
        TransDb.Close()

        'ShakerMortRate
        CmdStr = "SELECT * FROM ShakerMortRate WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim ShakerMortRatecm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim ShakerMortRateIDDA As New System.Data.OleDb.OleDbDataAdapter
        ShakerMortRateIDDA.SelectCommand = ShakerMortRatecm
        Dim ShakerMortRatecb As New OleDb.OleDbCommandBuilder
        ShakerMortRatecb = New OleDb.OleDbCommandBuilder(ShakerMortRateIDDA)
        If TransferDataSet.Tables.Contains("ShakerMortRate") Then
            TransferDataSet.Tables("ShakerMortRate").Clear()
        End If
        ShakerMortRateIDDA.Fill(TransferDataSet, "ShakerMortRate")
        Dim NumShakerMortRate As Integer
        NumShakerMortRate = TransferDataSet.Tables("ShakerMortRate").Rows.Count

        Dim ShakerMortRateTrans As OleDb.OleDbTransaction
        Dim ShakerMortRate As New OleDbCommand
        TransDb.Open()
        ShakerMortRateTrans = TransDb.BeginTransaction
        ShakerMortRate.Connection = TransDb
        ShakerMortRate.Transaction = ShakerMortRateTrans
        NumRecs = TransferDataSet.Tables("ShakerMortRate").Rows.Count
        For RecNum = 0 To NumRecs - 1
            ShakerMortRate.CommandText = "INSERT INTO ShakerMortRate (BasePeriodID,FisheryID,TimeStep,ShakerMortRate) " & _
               "VALUES(" & TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(4) & ")"

            ShakerMortRate.ExecuteNonQuery()
        Next
        ShakerMortRateTrans.Commit()
        TransDb.Close()

        'Stock
        CmdStr = "SELECT * FROM Stock"
        Dim Stockcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim StockIDDA As New System.Data.OleDb.OleDbDataAdapter
        StockIDDA.SelectCommand = Stockcm
        Dim Stockcb As New OleDb.OleDbCommandBuilder
        Stockcb = New OleDb.OleDbCommandBuilder(StockIDDA)
        If TransferDataSet.Tables.Contains("Stock") Then
            TransferDataSet.Tables("Stock").Clear()
        End If
        StockIDDA.Fill(TransferDataSet, "Stock")
        Dim NumStock As Integer
        NumStock = TransferDataSet.Tables("Stock").Rows.Count

        Dim StockTrans As OleDb.OleDbTransaction
        Dim Stock As New OleDbCommand
        TransDb.Open()
        StockTrans = TransDb.BeginTransaction
        Stock.Connection = TransDb
        Stock.Transaction = StockTrans
        NumRecs = TransferDataSet.Tables("Stock").Rows.Count
        For RecNum = 0 To NumRecs - 1
            Stock.CommandText = "INSERT INTO Stock (Species,StockVersion,StockID,ProductionRegionNumber,ManagementUnitNumber,StockName,StockLongName) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(2) * 2 & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(4) * 2 & "," & _
                Chr(34) & "M-" & TransferDataSet.Tables("Stock").Rows(RecNum)(5) & Chr(34) & "," & _
               Chr(34) & "Marked " & TransferDataSet.Tables("Stock").Rows(RecNum)(6) & Chr(34) & ")"

            Stock.ExecuteNonQuery()
        Next
        For RecNum = 0 To NumRecs - 1
            Stock.CommandText = "INSERT INTO Stock (Species,StockVersion,StockID,ProductionRegionNumber,ManagementUnitNumber,StockName,StockLongName) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(2) * 2 - 1 & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(4) * 2 - 1 & "," & _
                Chr(34) & "U-" & TransferDataSet.Tables("Stock").Rows(RecNum)(5) & Chr(34) & "," & _
               Chr(34) & "UnMarked " & TransferDataSet.Tables("Stock").Rows(RecNum)(6) & Chr(34) & ")"

            Stock.ExecuteNonQuery()
        Next
        StockTrans.Commit()
        TransDb.Close()

        'TerminalFisheryFlag
        CmdStr = "SELECT * FROM TerminalFisheryFlag WHERE BasePeriodID = " & SelectedBasePeriodID.ToString & ";"
        Dim TerminalFisheryFlagcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim TerminalFisheryFlagIDDA As New System.Data.OleDb.OleDbDataAdapter
        TerminalFisheryFlagIDDA.SelectCommand = TerminalFisheryFlagcm
        Dim TerminalFisheryFlagcb As New OleDb.OleDbCommandBuilder
        TerminalFisheryFlagcb = New OleDb.OleDbCommandBuilder(TerminalFisheryFlagIDDA)
        If TransferDataSet.Tables.Contains("TerminalFisheryFlag") Then
            TransferDataSet.Tables("TerminalFisheryFlag").Clear()
        End If
        TerminalFisheryFlagIDDA.Fill(TransferDataSet, "TerminalFisheryFlag")
        Dim NumTerminalFisheryFlag As Integer
        NumTerminalFisheryFlag = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count

        Dim TerminalFisheryFlagTrans As OleDb.OleDbTransaction
        Dim TerminalFisheryFlag As New OleDbCommand
        TransDb.Open()
        TerminalFisheryFlagTrans = TransDb.BeginTransaction
        TerminalFisheryFlag.Connection = TransDb
        TerminalFisheryFlag.Transaction = TerminalFisheryFlagTrans
        NumRecs = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count
        For RecNum = 0 To NumRecs - 1
            TerminalFisheryFlag.CommandText = "INSERT INTO TerminalFisheryFlag (BasePeriodID,FisheryID,TimeStep,TerminalFlag) " & _
               "VALUES(" & TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(4) & ")"

            TerminalFisheryFlag.ExecuteNonQuery()
        Next
        TerminalFisheryFlagTrans.Commit()
        TransDb.Close()

        'TimeStep
        CmdStr = "SELECT * FROM TimeStep"
        Dim TimeStepcm As New OleDb.OleDbCommand(CmdStr, CalibrationDB)
        Dim TimeStepIDDA As New System.Data.OleDb.OleDbDataAdapter
        TimeStepIDDA.SelectCommand = TimeStepcm
        Dim TimeStepcb As New OleDb.OleDbCommandBuilder
        TimeStepcb = New OleDb.OleDbCommandBuilder(TimeStepIDDA)
        If TransferDataSet.Tables.Contains("TimeStep") Then
            TransferDataSet.Tables("TimeStep").Clear()
        End If
        TimeStepIDDA.Fill(TransferDataSet, "TimeStep")
        Dim NumTimeStep As Integer
        NumTimeStep = TransferDataSet.Tables("TimeStep").Rows.Count

        Dim TimeStepTrans As OleDb.OleDbTransaction
        Dim TimeStep As New OleDbCommand
        TransDb.Open()
        TimeStepTrans = TransDb.BeginTransaction
        TimeStep.Connection = TransDb
        TimeStep.Transaction = TimeStepTrans
        NumRecs = TransferDataSet.Tables("TimeStep").Rows.Count
        For RecNum = 0 To NumRecs - 1
            TimeStep.CommandText = "INSERT INTO TimeStep (Species,VersionNumber,TimeStepID,TimeStepName,TimeStepTitle) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("TimeStep").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("TimeStep").Rows(RecNum)(2) & "," & _
                Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(3) & Chr(34) & "," & _
                Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(4) & Chr(34) & ")"

            TimeStep.ExecuteNonQuery()
        Next
        TimeStepTrans.Commit()
        TransDb.Close()


    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then 'cohort reconstruction w/o catch or escapement expansions or non-retention calcs
            OOBStatus = 4
            Firstpass = False
            NoExpansions = True
        End If
    End Sub
End Class