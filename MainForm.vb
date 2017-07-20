
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO

Public Class mainform

    'Dim SConStr
    Dim dbConnection As New OleDbConnection
    Dim dbConnection2 As New OleDbConnection
    Dim dbConnection3 As New OleDbConnection
    Dim i As Integer
    Dim iResult As String
    Dim dbAdapter As OleDbDataAdapter = New OleDb.OleDbDataAdapter()

    Dim dsSupportData As New DataSet


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Private Sub ListDBTables_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Set the Title and Filter properties, so that only .mdb 
        '    files are listed.

        OpenFileDialog1.Title = "Select a Database File"
        OpenFileDialog1.Filter = "DataBase Files(*.accdb)|*.accdb"
        'Set the CheckFileExists property so that a warning 
        '    appears if the user types a filename of a non-existent 
        '    file.  Then show the dialog.
        OpenFileDialog1.CheckFileExists = True
        iResult = OpenFileDialog1.ShowDialog()


        'Make sure the user did not click the Cancel button And 
        '    specified a file name for the file to be created.  
        If iResult <> Windows.Forms.DialogResult.Cancel And _
                    OpenFileDialog1.FileName.Length <> 0 Then
            FVSdatabasename = OpenFileDialog1.FileName
            SConStr = "Provider=Microsoft.ACE.OLEDB.12.0;"
            'Use the user-selected database file as the
            '    Data Source value. 
            SConStr &= "Data Source=" & _
                    OpenFileDialog1.FileName & ";"
            lblDBName.Text = OpenFileDialog1.FileName

            If Not dbConnection Is Nothing Then
                dbConnection.Close()
                dbConnection = Nothing
            End If
            'Dynamically create a New instance of an 
            '    OleDbConnection object and connect 
            '    it to the database specified in the
            '    ConnectionString (sConStr).
            dbConnection = New OleDbConnection(SConStr)
            'Open the Connection 
            dbConnection.Open()
            '    Use the GetOleDbSchemaTable method of the 
            '    Connection object to extract a DataTable 
            '    that contains the names of the Tables that 
            '    the database contains.
            Dim schemaTable As DataTable = _
                    dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, _
                    New Object() {Nothing, Nothing, Nothing, "TABLE"})
            'Once we're extracted the Table information, the 
            '    Connection can be closed.
            dbConnection.Close()
            'Clear the lstTables listbox before adding 
            '    items to it.
            lstTables.Items.Clear()
            'The Rows.Count property of the schemaTable, 
            '    is equal to the number of tables that the 
            '    database contains.
            For i = 0 To schemaTable.Rows.Count - 1
                'The Items collection is the columns of the 
                '    schemaTable.  Columns 0, and 1 are Null, 
                '    column 2 contains the table names.
                lstTables.Items.Add(schemaTable.Rows(i).Item(2))
            Next i
            If OOBStatus < 3 Then
                MsgBox("Please select the table with the out-of-base CWT recoveries.")
            Else
                MsgBox("Please select the table with the all-stocks or in-base CWT recoveries from Frambuilder.")
            End If
        End If
    End Sub



    Public Sub lstTables_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTables.SelectedIndexChanged

        Dim sFieldInfo, sDataType As String

        Dim i As Integer
        'Store the Table name selected in lstTables in sTableName

        oobTableName = lstTables.Items(lstTables.SelectedIndex)

        'Create an OleDbCommand object that specifies an SQL 
        '    Select command string that includes all the columns 
        '    from the selected Table as the first parameter, and the 
        '    name of the OleDbConnection object that contains the 
        '    table (dbConnection) as the second parameter.

        Dim selectCMD As OleDbCommand = _
                  New OleDbCommand("SELECT * FROM " & _
                           "[" & oobTableName & "]", dbConnection)
        'Bind the Data Adapter to the selected Table
        dbAdapter.SelectCommand = selectCMD

        'Create the DataSet object. By creating it locally 
        '    (as apposed to creating it in the Declarations  
        '    section) it is recreated each time, so new tables 
        '    are NOT appended to the DataSet.
        Dim dbDataSet As DataSet = New DataSet()
        'Clear and fill the new DataSet with the records 
        '    from the selected Table.
        dbDataSet.Clear()

        dbAdapter.Fill(dbDataSet, oobTableName)

        'Clear the lstColumns listbox before filling it
        lstColumns.Items.Clear()
        'Fill the lstColumns listbox with the names and data 
        '    type of each column in a row (record).
        For i = 0 To dbDataSet.Tables(0).Columns.Count - 1
            'Use a With clause here to make the code more readable
            With dbDataSet.Tables(0).Columns(i)
                '.DataType.ToString returns the data type preceded 
                '    by the word System, like this:   System.Integer
                'The following line of code extracts just the data type 
                '    (i.e. Integer) that follows System.
                sDataType = .DataType.ToString.Substring( _
                        InStrRev(.DataType.ToString, "."))
                'Concatenate the DataType info to the 
                '    ColumnName for display in the listbox.
                sFieldInfo = .ColumnName & " -- " & _
                                          sDataType.ToLower
                lstColumns.Items.Add(sFieldInfo)
            End With
        Next i
        'Create a BindingSource control and DataTable control
        Dim MyBindingSource As New BindingSource
        Dim MyDataTable As New DataTable

        'Fill MyDataTable with the records in dbAdapter
        dbAdapter.Fill(MyDataTable)

        'Bind the DataSource property to MyDataTable 
        MyBindingSource.DataSource = MyDataTable

        'Bind the DataSource property of DataGridView to 
        '    MyBindingSource to display the records.
        DataGridView1.DataSource = MyBindingSource

    End Sub

    Private Sub btnBigSmall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBigSmall.Click
        Me.Width += 100
        Me.Height += 50
        'Me.DataGridView1.Height += 50
        GroupBox3.Width += 100
        GroupBox3.Height += 50
        DataGridView1.Width += 100
        DataGridView1.Height += 50
    End Sub

    Public Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click


        Me.Close()

        dbConnection2 = New OleDbConnection(SConStr)
        'Open the Connection 
        dbConnection2.Open()

        'fill datasets using data adapters
        Dim daSupportInfo As New OleDbDataAdapter("Select * From SummaryInfo", dbConnection2)
        If OOBStatus < 3 Then
            Dim daOOBCWT As New OleDbDataAdapter("Select * From " & oobTableName & "", dbConnection2)
            daOOBCWT.Fill(dsSupportData, oobTableName)
        End If
        Dim daVB As New OleDbDataAdapter("Select * From Growth", dbConnection2)
        Dim daTermFlag As New OleDbDataAdapter("Select * From TerminalFisheryFlag", dbConnection2)
        Dim daSizeLimit As New OleDbDataAdapter("Select * From SizeLimits", dbConnection2)
        Dim daNatMort As New OleDbDataAdapter("Select * From NaturalMortality", dbConnection2)
        Dim daNonRet As New OleDbDataAdapter("Select * From NonRetention", dbConnection2)
        Dim daOtherMort As New OleDbDataAdapter("Select * From IncidentalRate", dbConnection2)
        Dim daShakerMortRate As New OleDbDataAdapter("Select * From ShakerMortRate", dbConnection2)
        Dim daEncRateAdjust As New OleDbDataAdapter("Select * From EncounterRateAdjustment", dbConnection2)
        Dim daBPCatch As New OleDbDataAdapter("Select * From BasePeriodCatch", dbConnection2)
        Dim daImpute As New OleDbDataAdapter("Select * From ImputeRecoveries", dbConnection2)
        Dim daTargetEncounterRates As New OleDbDataAdapter("Select * From TargetEncounterRates", dbConnection2)
        Dim daShakerInclusion As New OleDbDataAdapter("Select * From ShakerInclusion", dbConnection2)
        Dim daBPEscapements As New OleDbDataAdapter("Select * From BPEscapements", dbConnection2)
        Dim daFishScalar As New OleDbDataAdapter("Select * From FishScalar", dbConnection2)
        Dim daBPSizeLimit As New OleDbDataAdapter("Select * From BPSizeLimits", dbConnection2)
        Dim daWeighting As New OleDbDataAdapter("Select * From Weighting", dbConnection2)
        Dim daProportionTreatyNet As New OleDbDataAdapter("Select * From ProportionTreatyNet", dbConnection2)
        Dim daFisheryFlag As New OleDbDataAdapter("Select * From FisheryFlag", dbConnection2)
        Select Case OOBStatus
            Case 2, 3
                Dim daCWTAll As New OleDbDataAdapter("Select * From CWTAll", dbConnection2)
                daCWTAll.Fill(dsSupportData, "CWTAll")
            Case 4
                Dim daCWTAll As New OleDbDataAdapter("Select * From " & oobTableName & "", dbConnection2)
                daCWTAll.Fill(dsSupportData, oobTableName)
        End Select
        '########## Begin JC Update; 9/25/2015 ##########

        Dim daSurrogateFishBPER As New OleDbDataAdapter("Select * From SurrogateFishBPER", dbConnection2)
        daSurrogateFishBPER.Fill(dsSupportData, "SurrogateFishBPER")

        '########### End JC Update; 9/25/2015 ###########
        daSupportInfo.Fill(dsSupportData, "SummaryInfo")
        daVB.Fill(dsSupportData, "Growth")
        daTermFlag.Fill(dsSupportData, "TerminalFisheryFlag")
        daSizeLimit.Fill(dsSupportData, "SizeLimits")
        daNatMort.Fill(dsSupportData, "NaturalMortality")
        daNonRet.Fill(dsSupportData, "NonRetention")
        daOtherMort.Fill(dsSupportData, "IncidentalRate")
        daShakerMortRate.Fill(dsSupportData, "ShakerMortRate")
        daEncRateAdjust.Fill(dsSupportData, "EncounterRateAdjustment")
        daBPCatch.Fill(dsSupportData, "BasePeriodCatch")
        daImpute.Fill(dsSupportData, "ImputeRecoveries")
        daTargetEncounterRates.Fill(dsSupportData, "TargetEncounterRates")
        daShakerInclusion.Fill(dsSupportData, "ShakerInclusion")
        daBPEscapements.Fill(dsSupportData, "BPEscapements")
        daFishScalar.Fill(dsSupportData, "FishScalar")
        daSizeLimit.Fill(dsSupportData, "BPSizeLimits")
        daWeighting.Fill(dsSupportData, "Weighting")
        daProportionTreatyNet.Fill(dsSupportData, "ProportionTreatyNet")
        daFisheryFlag.Fill(dsSupportData, "FisheryFlag")

        dbConnection2.Close()

        'read in miscellaneous information from SupportInfo table
        NumStk = dsSupportData.Tables("SummaryInfo").Rows(0)("NumStk")
        NumFish = dsSupportData.Tables("SummaryInfo").Rows(0)("NumFish")
        NumSteps = dsSupportData.Tables("SummaryInfo").Rows(0)("NumSteps")
        MaxAge = dsSupportData.Tables("SummaryInfo").Rows(0)("MaxAge")
        MinBY = dsSupportData.Tables("SummaryInfo").Rows(0)("MinBroodYear")
        MaxBY = dsSupportData.Tables("SummaryInfo").Rows(0)("MaxBroodYear")
        CT = dsSupportData.Tables("SummaryInfo").Rows(0)("ConvergenceTolerance")
        MaxAgeEncAdj = dsSupportData.Tables("SummaryInfo").Rows(0)("MaxAgeEncounterRate")
        BasePeriodID = dsSupportData.Tables("SummaryInfo").Rows(0)("BasePeriodID")
        ReDim PntTime(NumSteps)
        PntTime(1) = dsSupportData.Tables("SummaryInfo").Rows(0)("MidPointT1")
        PntTime(2) = dsSupportData.Tables("SummaryInfo").Rows(0)("MidPointT2")
        PntTime(3) = dsSupportData.Tables("SummaryInfo").Rows(0)("MidPointT3")
        RunID = dsSupportData.Tables("SummaryInfo").Rows(0)("RunID")
        BaseYear = dsSupportData.Tables("SummaryInfo").Rows(0)("BaseYear")

        'read the FRAMBUILDER CWT table into a list 
        'use adjusted values instead of catch values when <> 0
        'Dim RecordCWT As CWTData
        'Dim CWTList As New List(Of CWTData)
        'code no longer needed, as FRAMbuilder presummarizes brood years AHB 5/19/2016
        If OOBStatus < 3 Then
            For irow = 0 To dsSupportData.Tables(oobTableName).Rows.Count - 1
                RecordCWT.cBaseID = dsSupportData.Tables(oobTableName).Rows(irow)("BasePeriodID")
                If RecordCWT.cBaseID = BasePeriodID Then
                    If dsSupportData.Tables(oobTableName).Rows(irow)("Adjusted").ToString = "" Then
                        'RecordCWT.cCatch = dsSupportData.Tables("CWT").Rows(irow)("Adjusted")
                        RecordCWT.cCatch = dsSupportData.Tables(oobTableName).Rows(irow)("Catch")
                    Else
                        RecordCWT.cCatch = dsSupportData.Tables(oobTableName).Rows(irow)("Adjusted")
                    End If
                    RecordCWT.cBY = dsSupportData.Tables(oobTableName).Rows(irow)("BroodYear")
                    'RecordCWT.cAdj = dsSupportData.Tables("CWT").Rows(irow)("Adjusted")
                    RecordCWT.cStk = dsSupportData.Tables(oobTableName).Rows(irow)("StockID")
                    RecordCWT.cAge = dsSupportData.Tables(oobTableName).Rows(irow)("Age")
                    RecordCWT.cFish = dsSupportData.Tables(oobTableName).Rows(irow)("FisheryID")
                    RecordCWT.cStage = dsSupportData.Tables(oobTableName).Rows(irow)("Stage")
                    RecordCWT.cTStep = dsSupportData.Tables(oobTableName).Rows(irow)("TimeStep")
                    RecordCWT.cLookUp = dsSupportData.Tables(oobTableName).Rows(irow)("StockID") & "-" & dsSupportData.Tables(oobTableName).Rows(irow)("BroodYear") _
                                        & "-" & dsSupportData.Tables(oobTableName).Rows(irow)("Stage")

                    CWTList.Add(RecordCWT)
                End If
            Next
        End If

        If OOBStatus > 0 Then
            If OOBStatus = 2 Then
                oobTableName = "CWTALL"
            End If


            ReDim CWTAll(NumStk, MaxAge, NumFish, NumSteps)
            Dim AllCatch As Double
            For irow = 0 To dsSupportData.Tables(oobTableName).Rows.Count - 1
                If dsSupportData.Tables(oobTableName).Rows(irow)("BasePeriodID") = BasePeriodID Then
                    If dsSupportData.Tables(oobTableName).Rows(irow)("Adjusted").ToString = "" Then
                        AllCatch = dsSupportData.Tables(oobTableName).Rows(irow)("Catch")
                    Else
                        AllCatch = dsSupportData.Tables(oobTableName).Rows(irow)("Adjusted")
                    End If
                    STk = dsSupportData.Tables(oobTableName).Rows(irow)("StockID")
                    Age = dsSupportData.Tables(oobTableName).Rows(irow)("Age")
                    Fish = dsSupportData.Tables(oobTableName).Rows(irow)("FisheryID")
                    TStep = dsSupportData.Tables(oobTableName).Rows(irow)("TimeStep")
                    CWTAll(STk, Age, Fish, TStep) = AllCatch
                End If
            Next
        End If


        'read VB parameters
        ReDim L(NumStk, 1)
        ReDim K(NumStk, 1)
        ReDim T0(NumStk, 1)
        ReDim CV(NumStk, MaxAge, 1)
        For irow = 0 To dsSupportData.Tables("Growth").Rows.Count - 1
            If dsSupportData.Tables("Growth").Rows(irow)("BasePeriodID") = BasePeriodID Then
                STk = dsSupportData.Tables("Growth").Rows(irow)("StockID")

                L(STk, 0) = dsSupportData.Tables("Growth").Rows(irow)("LImmature")
                K(STk, 0) = dsSupportData.Tables("Growth").Rows(irow)("KImmature")
                T0(STk, 0) = dsSupportData.Tables("Growth").Rows(irow)("TImmature")
                CV(STk, 2, 0) = dsSupportData.Tables("Growth").Rows(irow)("CV2Immature")
                CV(STk, 3, 0) = dsSupportData.Tables("Growth").Rows(irow)("CV3Immature")
                CV(STk, 4, 0) = dsSupportData.Tables("Growth").Rows(irow)("CV4Immature")
                CV(STk, 5, 0) = dsSupportData.Tables("Growth").Rows(irow)("CV5Immature")
                L(STk, 1) = dsSupportData.Tables("Growth").Rows(irow)("LMature")
                K(STk, 1) = dsSupportData.Tables("Growth").Rows(irow)("KMature")
                T0(STk, 1) = dsSupportData.Tables("Growth").Rows(irow)("TMature")
                CV(STk, 2, 1) = dsSupportData.Tables("Growth").Rows(irow)("CV2Mature")
                CV(STk, 3, 1) = dsSupportData.Tables("Growth").Rows(irow)("CV3Mature")
                CV(STk, 4, 1) = dsSupportData.Tables("Growth").Rows(irow)("CV4Mature")
                CV(STk, 5, 1) = dsSupportData.Tables("Growth").Rows(irow)("CV5Mature")
            End If
        Next
        'Read base period escapements
        ReDim ObsEscpmnt(NumStk)
        n = 0
        For irow = 0 To dsSupportData.Tables("BPEscapements").Rows.Count - 1
            EscBaseID = dsSupportData.Tables("BPEscapements").Rows(irow)("BasePeriodID")
            If EscBaseID = BasePeriodID Then
                n = n + 1
                STk = dsSupportData.Tables("BPEscapements").Rows(irow)("StockID")
                ObsEscpmnt(STk) = dsSupportData.Tables("BPEscapements").Rows(irow)("Escapement")
            End If
        Next
        If n = 0 Then
            MsgBox("The Base Period Escapement Table does not contain the correct Base Period Number from Summary Info Table.")
        End If

        'READ TERMINAL FISHERY FLAGS
        ReDim TermFlag(NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("TerminalFisheryFlag").Rows.Count - 1
            If dsSupportData.Tables("TerminalFisheryFlag").Rows(irow)("BasePeriodID") = BasePeriodID Then
                Fish = dsSupportData.Tables("TerminalFisheryFlag").Rows(irow)("FisheryID")
                TStep = dsSupportData.Tables("TerminalFisheryFlag").Rows(irow)("TimeStep")

                TermFlag(Fish, TStep) = dsSupportData.Tables("TerminalFisheryFlag").Rows(irow)("TerminalFlag")
            End If
        Next
        ' READ TARGET ENCOUNTER RATES
        ReDim TargEncRate(NumFish, NumSteps)
        For irow = 0 To dsSupportData.Tables("TargetEncounterRates").Rows.Count - 1

            Fish = dsSupportData.Tables("TargetEncounterRates").Rows(irow)("FisheryID")
            TStep = dsSupportData.Tables("TargetEncounterRates").Rows(irow)("TimeStep")

            TargEncRate(Fish, TStep) = dsSupportData.Tables("TargetEncounterRates").Rows(irow)("TargetEncounterRate")
        Next

        ' READ Shaker Inclusion Flags
        ReDim StkCheck(NumStk, NumFish)
        For irow = 0 To dsSupportData.Tables("ShakerInclusion").Rows.Count - 1

            Fish = dsSupportData.Tables("ShakerInclusion").Rows(irow)("FisheryID")
            STk = dsSupportData.Tables("ShakerInclusion").Rows(irow)("StockID")

            StkCheck(STk, Fish) = dsSupportData.Tables("ShakerInclusion").Rows(irow)("ShakerInclusionFlag")
        Next


        'READ MINIMUM SIZE LIMITS


        ReDim MinSize(2100, NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("SizeLimits").Rows.Count - 1
            'If dsSupportData.Tables("SizeLimits").Rows(irow)("RunID") = RunID Then
            Fish = dsSupportData.Tables("SizeLimits").Rows(irow)("FisheryID")
            TStep = dsSupportData.Tables("SizeLimits").Rows(irow)("TimeStep")
            FishYear = dsSupportData.Tables("SizeLimits").Rows(irow)("FishingYear")
            MinSize(FishYear, Fish, TStep) = dsSupportData.Tables("SizeLimits").Rows(irow)("MinimumSize")
            'End If
        Next

        'READ BASE PERIOD MINIMUM SIZE LIMITS

        'not used by program, instead program uses base period year from "SummaryInfo" table to fine the size limit in the size limit table
        ReDim BPMinSize(NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("BPSizeLimits").Rows.Count - 1
            Fish = dsSupportData.Tables("BPSizeLimits").Rows(irow)("FisheryID")
            TStep = dsSupportData.Tables("BPSizeLimits").Rows(irow)("TimeStep")
            BPMinSize(Fish, TStep) = dsSupportData.Tables("BPSizeLimits").Rows(irow)("MinimumSize")
        Next


        'READ NATURAL MORTALITY RATES

        ReDim SurvRate(MaxAge, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("NaturalMortality").Rows.Count - 1
            If dsSupportData.Tables("NaturalMortality").Rows(irow)("BasePeriodID") = BasePeriodID Then
                Age = dsSupportData.Tables("NaturalMortality").Rows(irow)("Age")
                TStep = dsSupportData.Tables("NaturalMortality").Rows(irow)("TimeStep")

                SurvRate(Age, TStep) = 1 - dsSupportData.Tables("NaturalMortality").Rows(irow)("NaturalMortalityRate")
            End If
        Next

        'READ SHAKER MORTALITY RATES

        ReDim ShakMortRate(NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("ShakerMortRate").Rows.Count - 1
            If dsSupportData.Tables("ShakerMortRate").Rows(irow)("BasePeriodID") = BasePeriodID Then
                Fish = dsSupportData.Tables("ShakerMortRate").Rows(irow)("FisheryID")
                TStep = dsSupportData.Tables("ShakerMortRate").Rows(irow)("TimeStep")

                ShakMortRate(Fish, TStep) = dsSupportData.Tables("ShakerMortRate").Rows(irow)("ShakerMortRate")
            End If
        Next

        'READ OTHER MORTALITY (DROPOFF, PREDATION, ETC.,
        'EXPRESSED AS PERCENT OF CATCH).
        ReDim OtherMort(NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("IncidentalRate").Rows.Count - 1
            If dsSupportData.Tables("IncidentalRate").Rows(irow)("BasePeriodID") = BasePeriodID Then
                Fish = dsSupportData.Tables("IncidentalRate").Rows(irow)("FisheryID")
                TStep = dsSupportData.Tables("IncidentalRate").Rows(irow)("TimeStep")

                OtherMort(Fish, TStep) = dsSupportData.Tables("IncidentalRate").Rows(irow)("IncidentalRate")
            End If
        Next

        'READ ENCOUNTER RATE Adjustments
        ReDim EncRateAdj(MaxAge, NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("EncounterRateAdjustment").Rows.Count - 1
            If dsSupportData.Tables("EncounterRateAdjustment").Rows(irow)("BasePeriodID") = BasePeriodID Then
                Age = dsSupportData.Tables("EncounterRateAdjustment").Rows(irow)("Age")
                TStep = dsSupportData.Tables("EncounterRateAdjustment").Rows(irow)("TimeStep")
                Fish = dsSupportData.Tables("EncounterRateAdjustment").Rows(irow)("FisheryID")
                EncRateAdj(Age, Fish, TStep) = dsSupportData.Tables("EncounterRateAdjustment").Rows(irow)("EncounterRateAdjustment")
            End If
        Next

        'READ CNR PARAMETERS

        ReDim CNRFlag(1, 2100, NumFish, NumSteps + 1)
        ReDim CNRInput(1, 2100, 4, NumFish, NumSteps + 1)

        For irow = 0 To dsSupportData.Tables("NonRetention").Rows.Count - 1

            Fish = dsSupportData.Tables("NonRetention").Rows(irow)("FisheryID")
            TStep = dsSupportData.Tables("NonRetention").Rows(irow)("TimeStep")
            Yr = dsSupportData.Tables("NonRetention").Rows(irow)("Year")
            BaseType = dsSupportData.Tables("NonRetention").Rows(irow)("Type")

            CNRFlag(BaseType, Yr, Fish, TStep) = dsSupportData.Tables("NonRetention").Rows(irow)("NonRetentionFlag")
            CNRInput(BaseType, Yr, 1, Fish, TStep) = dsSupportData.Tables("NonRetention").Rows(irow)("CNRInput1")
            CNRInput(BaseType, Yr, 2, Fish, TStep) = dsSupportData.Tables("NonRetention").Rows(irow)("CNRInput2")
            CNRInput(BaseType, Yr, 3, Fish, TStep) = dsSupportData.Tables("NonRetention").Rows(irow)("CNRInput3")
            CNRInput(BaseType, Yr, 4, Fish, TStep) = dsSupportData.Tables("NonRetention").Rows(irow)("CNRInput4")

        Next

        'READ CATCH DATA


        ReDim TrueCatch(NumFish)
        ReDim CatchFlag(NumFish)
        ReDim ExternalModelStockProportion(NumFish)
        Dim CatchBaseID As Integer
        n = 0
        For irow = 0 To dsSupportData.Tables("BasePeriodCatch").Rows.Count - 1
            CatchBaseID = dsSupportData.Tables("BasePeriodCatch").Rows(irow)("BasePeriodID")
            If CatchBaseID = BasePeriodID Then
                n = n + 1
                Fish = dsSupportData.Tables("BasePeriodCatch").Rows(irow)("FisheryID")
                TrueCatch(Fish) = dsSupportData.Tables("BasePeriodCatch").Rows(irow)("BPCatch")
                CatchFlag(Fish) = dsSupportData.Tables("BasePeriodCatch").Rows(irow)("Flag")
                ExternalModelStockProportion(Fish) = dsSupportData.Tables("BasePeriodCatch").Rows(irow)("ModelStockProportion")
            End If
        Next
        If n = 0 Then
            MsgBox("The Base Period Catch Table does not contain the correct Base Period Number from Summary Info Table.")
        End If

        'READ FLAGS FOR FISHERIES TO IMPUTE DATA
        'the imputed fishery is listed in the index of the array, the source fishery number is the value


        For irow = 0 To dsSupportData.Tables("ImputeRecoveries").Rows.Count - 1
            ImputeItem.cBaseType = dsSupportData.Tables("ImputeRecoveries").Rows(irow)("Type")
            ImputeItem.cSurrogateFish = dsSupportData.Tables("ImputeRecoveries").Rows(irow)("SurrogateFishery")
            ImputeItem.cSurrogateTStep = dsSupportData.Tables("ImputeRecoveries").Rows(irow)("SurrogateTimeStep")
            ImputeItem.cRecipientFish = dsSupportData.Tables("ImputeRecoveries").Rows(irow)("RecipientFishery")
            ImputeItem.cRecipientTStep = dsSupportData.Tables("ImputeRecoveries").Rows(irow)("RecipientTimeStep")

            ImputeList.Add(ImputeItem)
        Next

        '########## Begin JC Update; 9/25/2015 ##########



        'Read old base period exploitation rates for fisheries/TSteps using old BP as surrogate
        ReDim SurrogateFishBP_ER(NumStk, MaxAge, NumFish, NumSteps)
        Dim BPER As Double
        For irow = 0 To dsSupportData.Tables("SurrogateFishBPER").Rows.Count - 1
            STk = dsSupportData.Tables("SurrogateFishBPER").Rows(irow)("StockID")
            Age = dsSupportData.Tables("SurrogateFishBPER").Rows(irow)("Age")
            Fish = dsSupportData.Tables("SurrogateFishBPER").Rows(irow)("FisheryID")
            TStep = dsSupportData.Tables("SurrogateFishBPER").Rows(irow)("TimeStep")
            BPER = dsSupportData.Tables("SurrogateFishBPER").Rows(irow)("ExploitationRate")
            SurrogateFishBP_ER(STk, Age, Fish, TStep) = BPER
        Next

        '########### End JC Update; 9/25/2015 ###########

        'read Fishery Scalars

        ReDim FishScalar(MaxBY + 5, NumFish, NumSteps)
        For irow = 0 To dsSupportData.Tables("FishScalar").Rows.Count - 1


            TStep = dsSupportData.Tables("FishScalar").Rows(irow)("TimeStep")
            FishYear = dsSupportData.Tables("FishScalar").Rows(irow)("FishingYear")
            Fish = dsSupportData.Tables("FishScalar").Rows(irow)("FisheryID")
            FishScalar(FishYear, Fish, TStep) = dsSupportData.Tables("FishScalar").Rows(irow)("FishScalar")
        Next

        ' READ Weights for combining CWT recoveries from multiple broody years
        'no longer needed as weighting occurs in Frambuilder
        For irow = 0 To dsSupportData.Tables("Weighting").Rows.Count - 1
            WeightItem.wStk = dsSupportData.Tables("Weighting").Rows(irow)("StockID")
            WeightItem.wBY = dsSupportData.Tables("Weighting").Rows(irow)("BroodYear")
            WeightItem.wStage = dsSupportData.Tables("Weighting").Rows(irow)("Stage")
            WeightItem.wWeight = dsSupportData.Tables("Weighting").Rows(irow)("Weight")
            WeightItem.wFlag = dsSupportData.Tables("Weighting").Rows(irow)("Flag")

            WeightList.Add(WeightItem)
        Next

        'read Proportion Treaty Net
        ReDim PropTNet(NumFish + 1, NumSteps + 1)
        For irow = 0 To dsSupportData.Tables("ProportionTreatyNet").Rows.Count - 1
            TStep = dsSupportData.Tables("ProportionTreatyNet").Rows(irow)("TimeStep")
            Fish = dsSupportData.Tables("ProportionTreatyNet").Rows(irow)("FisheryID")
            PropTNet(Fish, TStep) = dsSupportData.Tables("ProportionTreatyNet").Rows(irow)("PpnTreaty")
        Next

        'read Fishery Flag
        ReDim FishFlag(NumFish + 1)
        For irow = 0 To dsSupportData.Tables("FisheryFlag").Rows.Count - 1
            Fish = dsSupportData.Tables("FisheryFlag").Rows(irow)("FisheryID")
            FishFlag(Fish) = dsSupportData.Tables("FisheryFlag").Rows(irow)("FishFlag")
        Next
        '***************************************************************************
        'CHDAT
        '*************************************************************************************
        'Impute recoveries using surrogate fisheries
        'If OOBStatus < 3 Then
secondpass:



        FileLength = OpenFileDialog1.FileName.Length
        FullName = OpenFileDialog1.FileName
        filepath = "a"
        For i = FileLength To 0 Step -1
            findme = Mid(FullName, i, 1)
            If findme = "\" Then
                filepath = Mid(FullName, 1, i)
                i = 0
            End If
        Next

        If OOBStatus < 4 Then
            Impute()
        End If
        'End If
        '*************************************************************************************
        'Reject recoveries without legal population size

        Dim errfile As String
        errfile = filepath & "errfile.txt"
        FileOpen(12, errfile, OpenMode.Output)



        'Check for legal population when catch exists
        CheckLegal()
        '************************************************************************************

        'CheckCNR()

        Dim a As CWTData

        'Print(11, "Stage, BY, Stk, Age, Fish, TStep, Catch" & vbCrLf)
        'For Each a In CWTList
        'Print(11, a.cStage & ", ")
        'Print(11, a.cBY & ", ")
        'Print(11, a.cStk & ", ")
        'Print(11, a.cAge & ",    ")
        'Print(11, a.cFish & ", ")
        'Print(11, a.cTStep & ", ")
        'Print(11, a.cCatch & vbCrLf)
        'Next

        'FileClose()

        '************************************************************************************
        'SET FLAGS FOR INCLUSION OF STOCK IN SHAKER COMPUTATIONS
        'AHB: no longer used
        If Firstpass = False Then
            'ShakDistr2()
        End If
        ''*********************************************************************************
        ''SAVE RESULTS FROM SHAKER INCLUSION COMPUTATIONS
        'no need to save - results are stored in array StkCheck(stk,fish)
        'SaveStkCheck()
        ''****************************************************************************************
        '' read in and print the external exp rates
        'AHB couldn't find an example that had non-zero fisheries
        'If BaseRun = "N" Or BaseRun = "n" Then
        '    Impute2()
        'End If
        '***************************************************************************************
        'Start ChCal
        '***************************************************************************************

        ReDim Shaker(MaxAge, NumFish, NumSteps)
        ReDim CNR(NumStk, MaxAge, NumFish, NumSteps)
        'Dim CWTRow As CWTData
        Dim TheDictionary As New Dictionary(Of String, Integer)
        Dim LookupItem As String
        Dim x As Integer
        ''ReDim TotMortTerm(MaxAge, NumSteps)
        ''ReDim TotMortPterm(MaxAge, NumSteps)
        Dim pair As KeyValuePair(Of String, Integer)
        ReDim Age2Cohort(NumStk)
        ReDim ExRate(100, NumStk, 2, MaxAge, NumFish, NumSteps)
        ReDim MatRate(NumStk, MaxAge, NumSteps) 'might have to redeminsion if it is needed by brood year and stage
        ReDim TotalStk(NumStk)

        If Firstpass = False Then
            BaseType = 1 'all stocks run
        Else
            BaseType = 0 'OOB run
        End If

        If Firstpass = True Then


            TheDictionary.Clear()
            n = 1
            x = 1
            For Each CWTRow In CWTList
                x = x + 1
                LookupItem = CWTRow.cLookUp
                If Not TheDictionary.ContainsKey(LookupItem) Then
                    TheDictionary.Add(LookupItem, n)
                    n = n + 1
                End If
            Next

            'add catches by stock, BY, stage, TS, age over fisheries 
            'Dim testcatch As String
            'testcatch = "C:\data\calibration\07Qbasic\testcatch"
            'FileOpen(15, testcatch, OpenMode.Output)
            'Print(15, "BY, Stock,  Age, Fish, TStep, catch" & vbCrLf)


            ImputeListMain.Clear()

            ERGreater1file = filepath & "ERGreater1.txt"
            FileOpen(13, ERGreater1file, OpenMode.Output)

            For Each pair In TheDictionary ' deals with each unique stock, broodyear, stage (fing/year) combination
                DictionaryKey = pair.Key

                'retrieve the stock number, brood year, and stage
                If Mid(DictionaryKey, 2, 1) = "-" Then
                    Stknum = CInt(Mid(DictionaryKey, 1, 1))
                    BYNum = CInt(Mid(DictionaryKey, 3, 4))
                    StageNum = CInt(Mid(DictionaryKey, 8, 1))
                Else
                    Stknum = CInt(Mid(DictionaryKey, 1, 2))
                    BYNum = CInt(Mid(DictionaryKey, 4, 4))
                    StageNum = CInt(Mid(DictionaryKey, 9, 1))
                End If



                MinStk = Stknum
                MaxStk = Stknum

                'overwrites target encounter rates from Access

                'For Fish = 1 To NumFish - 1
                '    For TStep = 1 To NumSteps
                '        TargEncRate(Fish, TStep) = 1
                '    Next
                'Next


                'Compute CWT Escapement
                CompEscape()

                'COMPUTE CWT EXPANSION FACTOR

                ExpFact = 1


                'COMPUTE TOTAL CATCH IN EACH FISHERY PRIOR TO ADJUSTMENT
                ' OF RECOVERIES
                Addcatch()

                'COMPUTE ADJUSTMENT FACTORS FOR RECOVERIES - not needed for OOB run 
                'adjcatch()

                'COMPUTE TOTAL CATCH IN EACH FISHERY AFTER ADJUSTMENT - not needed for OOB run
                ' OF RECOVERIES
                'Addcatch()


                'RECONSTRUCT COHORT FROM CATCH DATA
                CompCohort()


                'ITERATIVE LOOP TO RECONSTRUCT COHORT WITH NONCATCH MORTALITY.
                ' LOOP CONTINUES UNTIL AGE 2 COHORT ABUNDANCE STABILIZES FOR ALL STOCKS.
                TotLanded = 0
                ConvFlag = "FALSE"
                While ConvFlag = "FALSE"
                    IncMort() 'shaker mort & CNR
                    SaveCohort()
                    CompCohort()
                    CohortCheck()
                End While
                If OOBStatus > 2 Or Firstpass = False Then
                    'SaveDat()
                Else

                    'Compute Maturation Rates
                    MaturationRate()
                    'Compute Exploitation Rates in Recovery Year
                    ExploitationRate()
                    Forward()
                    'create a newlist with imputed OOB recoveries
                End If
            Next 'new stock, BY, stage combo
            FileClose(13)
        Else
            ReDim ShakerAll(NumStk, MaxAge, NumFish, NumSteps)
            MinStk = 1
            MaxStk = NumStk
            AdjustedCatch = False

            If NoExpansions = True Then 'no escapement expansions
                ReDim EscExpFact(NumStk)
                For STk = 1 To NumStk
                    EscExpFact(STk) = 1
                Next STk
            Else
                CompExpFact()
            End If

            Addcatch()
            adjcatch()
            Addcatch()
            CompCohort()

            ConvFlag = "FALSE"
            While ConvFlag = "FALSE"

                'ImputeOldBPERs()
                If NoExpansions = False Then
                    ImputeOldBPERs()
                    IncMort() 'shaker mort - sublegal release mortality
                End If
                SaveCohort()
                CompCohort()
                CohortCheck()
            End While

        End If




        If Firstpass = True Then

            Merge() 'merges information from different brood years and life stages


            'save merged out of base recoveries to CWTOOBMerged
            'Try
            dbConnection2 = New OleDbConnection(SConStr)
            'Open the Connection 
            dbConnection2.Open()
            Dim daCWTOOBMerged As New OleDbDataAdapter("Select * From CWTOOBMerged", dbConnection2)
            'delete all records
            Dim cmd As New OleDbCommand
            cmd = New OleDbCommand("Delete from CWTOOBMerged", dbConnection2)
            cmd.ExecuteNonQuery()
            'add merged CWT data 
            Dim sqlstring As String
            Dim cmd2 As OleDbCommand

            For STk = 1 To NumStk
                For Age = 2 To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps
                            If MergedCWT(STk, Age, Fish, TStep) <> 0 Then
                                sqlstring = "INSERT INTO CWTOOBMerged(BasePeriodID,Species,StockID,AGE,FisheryID,TimeStep,Catch,RunTime) VALUES('" & BasePeriodID & "','" & 1 & " ','" & STk & "','" & Age & "','" & Fish & "','" & TStep & "','" & FormatNumber(MergedCWT(STk, Age, Fish, TStep), 4) & " ','" & Now() & "');"
                                'sqlstring = "INSERT INTO CWTOOBMerged(BasePeriodID,Species,StockID,AGE,FisheryID,TimeStep,Catch) VALUES('" & BasePeriodID & "','" & 1 & " ','" & STk & "','" & Age & "','" & Fish & "','" & TStep & "','" & FormatNumber(MergedCWT(STk, Age, Fish, TStep), 4) & " ');"
                                cmd2 = New OleDbCommand(sqlstring, dbConnection2)
                                cmd2.ExecuteNonQuery()
                            End If
                        Next
                    Next
                Next
            Next
            dbConnection2.Close()
        Else
            SaveDat()


            Dim AEQTable As New DataTable
            Dim BaseCohortTable As New DataTable
            Dim FisheryModelStockProportionTable As New DataTable
            Dim BaseExploitationRateTable As New DataTable
            Dim MaturationRateTable As New DataTable

            AEQTable.Columns.Add("BasePeriodID", GetType(Integer))
            AEQTable.Columns.Add("StockID", GetType(Integer))
            AEQTable.Columns.Add("Age", GetType(Integer))
            AEQTable.Columns.Add("TimeStep", GetType(Integer))
            AEQTable.Columns.Add("AEQ", GetType(Double))
            AEQTable.Columns.Add("RunTime", GetType(String))

            BaseCohortTable.Columns.Add("BasePeriodID", GetType(Integer))
            BaseCohortTable.Columns.Add("StockID", GetType(Integer))
            BaseCohortTable.Columns.Add("Age", GetType(Integer))
            BaseCohortTable.Columns.Add("BaseCohortSize", GetType(Double))

            FisheryModelStockProportionTable.Columns.Add("BasePeriodID", GetType(Integer))
            FisheryModelStockProportionTable.Columns.Add("FisheryID", GetType(Integer))
            FisheryModelStockProportionTable.Columns.Add("ModelStockProportion", GetType(Double))

            BaseExploitationRateTable.Columns.Add("BasePeriodID", GetType(Integer))
            BaseExploitationRateTable.Columns.Add("StockID", GetType(Integer))
            BaseExploitationRateTable.Columns.Add("Age", GetType(Integer))
            BaseExploitationRateTable.Columns.Add("FisheryID", GetType(Integer))
            BaseExploitationRateTable.Columns.Add("TimeStep", GetType(Integer))
            BaseExploitationRateTable.Columns.Add("ExploitationRate", GetType(Double))
            BaseExploitationRateTable.Columns.Add("SubLegalEncounterRate", GetType(Double))
            BaseExploitationRateTable.Columns.Add("RunTime", GetType(String))

            MaturationRateTable.Columns.Add("BasePeriodID", GetType(Integer))
            MaturationRateTable.Columns.Add("StockID", GetType(Integer))
            MaturationRateTable.Columns.Add("Age", GetType(Integer))
            MaturationRateTable.Columns.Add("TimeStep", GetType(Integer))
            MaturationRateTable.Columns.Add("MaturationRate", GetType(Double))
            MaturationRateTable.Columns.Add("RunTime", GetType(String))
            ''add data from SaveDat

            For STk = 1 To NumStk
                For Age = 2 To MaxAge
                    'set cohort size to at least 1, otherwise BKFRAM has to iterate for a long time to reach target
                    If StartCohort(STk, Age) < 1 Then
                        StartCohort(STk, Age) = 1
                    End If
                    BaseCohortTable.Rows.Add(BasePeriodID, STk, Age, FormatNumber(StartCohort(STk, Age), 3))
                    For TStep = 1 To NumSteps + 1
                        AEQTable.Rows.Add(BasePeriodID, STk, Age, TStep, FormatNumber(AEQ(STk, Age, TStep), 6), Now().ToString)
                        MaturationRateTable.Rows.Add(BasePeriodID, STk, Age, TStep, FormatNumber(MatRate(STk, Age, TStep), 6), Now().ToString)
                        For Fish = 1 To NumFish - 1
                            If ExRateAll(STk, Age, Fish, TStep) <> 0 Or EncRateAllShaker(STk, Age, Fish, TStep) <> 0 Then
                                BaseExploitationRateTable.Rows.Add(BasePeriodID, STk, Age, Fish, TStep, FormatNumber(ExRateAll(STk, Age, Fish, TStep), 8), FormatNumber(EncRateAllShaker(STk, Age, Fish, TStep), 8), Now().ToString)
                            End If
                        Next
                    Next
                Next
            Next

            For Fish = 1 To NumFish - 1
                FisheryModelStockProportionTable.Rows.Add(BasePeriodID, Fish, FormatNumber(ModelStkPPN(Fish), 4))
            Next


            dbConnection3 = New OleDbConnection(SConStr)
            'Open the Connection 
            dbConnection3.Open()
            Dim daAEQ As New OleDbDataAdapter("Select * From AEQ", dbConnection3)
            Dim daBaseCohort As New OleDbDataAdapter("Select * From BaseCohort", dbConnection3)
            Dim daModelStkPPN As New OleDbDataAdapter("Select * From FisheryModelStockProportion", dbConnection3)
            Dim daBaseExploitationRate As New OleDbDataAdapter("Select * From BaseExploitationRate", dbConnection3)
            Dim daMaturation As New OleDbDataAdapter("Select * From MaturationRate", dbConnection3)
            'delete all records
            Dim cmd10, cmd11, cmd12, cmd13, cmd14 As New OleDbCommand
            cmd10 = New OleDbCommand("Delete from AEQ", dbConnection3)
            cmd10.ExecuteNonQuery()

            cmd11 = New OleDbCommand("Delete from BaseCohort", dbConnection3)
            cmd11.ExecuteNonQuery()

            cmd12 = New OleDbCommand("Delete from FisheryModelStockProportion", dbConnection3)
            cmd12.ExecuteNonQuery()

            cmd13 = New OleDbCommand("Delete from BaseExploitationRate", dbConnection3)
            cmd13.ExecuteNonQuery()

            cmd14 = New OleDbCommand("Delete from MaturationRate", dbConnection3)
            cmd14.ExecuteNonQuery()




            Dim oleCommander As New OleDb.OleDbCommandBuilder(daAEQ)
            daAEQ.Update(AEQTable)

            Dim oleCommander2 As New OleDb.OleDbCommandBuilder(daBaseCohort)
            daBaseCohort.Update(BaseCohortTable)

            Dim oleCommander3 As New OleDb.OleDbCommandBuilder(daModelStkPPN)
            daModelStkPPN.Update(FisheryModelStockProportionTable)

            Dim oleCommander5 As New OleDb.OleDbCommandBuilder(daMaturation)
            daMaturation.Update(MaturationRateTable)

            Dim oleCommander4 As New OleDb.OleDbCommandBuilder(daBaseExploitationRate)
            daBaseExploitationRate.Update(BaseExploitationRateTable)



            dbConnection3.Close()
        End If

        If OOBStatus = 2 Then
            'create joined allstocks & OOB CWT array
            For STk = 1 To NumStk
                For Age = 2 To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps
                            If TotalStk(STk) <> 0 Then
                                CWTAll(STk, Age, Fish, TStep) = MergedCWT(STk, Age, Fish, TStep)

                            End If

                        Next
                    Next
                Next
            Next
            OOBStatus = 3
            GoTo secondpass 'call me lazy
        Else
            FileClose()
            Me.Close()
        End If
        StartForm.Show()
        FileClose(11)
    End Sub

    Private Sub mainform_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub chkOldBPSurrogateFish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class