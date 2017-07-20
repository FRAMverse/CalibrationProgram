Imports System.Collections.Generic
Imports System.Data.OleDb

Module Var
    Public a As Integer
    Public awatch As Double
    Public AbsZ As Double
    Public AdjCWTCatch(,,,,,) As Double
    Public AEQ(,,) As Double
    Public AdjustedCatch As Boolean
    Public Age As Integer
    Public AgeCatch(,,) As Double
    Public Age2Cohort() As Double
    Public aline As String
    Public AnnSurvRate() As Double
    Public AnnualCatch() As Double
    Public Area As Double
    Public BasePeriodID As Integer
    Public BaseID() As Integer
    Public BPMinSize(,) As Integer
    Public BaseRun As String
    Public BaseType As Integer
    Public BaseYear As Integer
    Public BPERFisheries() As Integer
    Public BroodYear As String
    Public BY As Integer
    Public BYFlag As Integer
    Public BYnew As Integer
    Public BYNum As Integer
    Public BYSTKCatch As Double
    Public BYWeight As Double
    Public Calfile As String
    Public CalibrationDB As New OleDbConnection
    Public Catcha(,,,) As Double
    Public Catchb As Double
    Public CatchFlag() As Integer
    Public Check As Integer
    Public CheckFish As Integer
    Public CNR(,,,) As Double
    Public CNRLeg(,,,) As Double
    Public CNRSub(,,,) As Double
    Public CNRFish As Integer
    Public CNRFlag(,,,) As Integer
    Public CNRInput(,,,,) As Double
    Public Cohort(,) As Double
    Public Cohort2 As Double
    Public CohortAll(,,,) As Double
    Public CohortTerm(,) As Double
    Public coma As String
    Public CompCatch(,,) As Double
    Public CompEscapement(,) As Double
    Public Conc() As Double
    Public ConvFlag As Boolean
    Public ConvTol As String
    Public CT As Double
    Public CV(,,)
    Public CWTAll(,,,) As Double
    Public CWTCatch(,,,,,) As Double
    'Public CWTRow As CWTData
    'Public TheDictionary As New Dictionary(Of String, Integer)
    'Public LookupItem As String
    'Public x As Integer
    'Public pair As KeyValuePair(Of String, Integer)
    'Public RecordCWT As CWTData
    'Public CWTList As New List(Of CWTData)
    Public CWTEscpmnt As Double
    Public DictionaryKey As String
    Public OldDif As Double
    Public Dif As Double
    Public EDTFile As String
    Public Encounters As Double
    Public EncRateAdj(,,) As Double
    Public EncRateAllShaker(,,,) As Double
    Public ErrFile As String
    Public ERGreater1file As String
    Public Escape(,) As Double
    Public EscBaseID As Integer
    Public EscapeAll(,,) As Double
    Public ExAdjFact(,,) As Double
    Public Expander As Double
    Public ExpFact As Double
    Public EscExpFact() As Double
    Public ExpFactor(,,) As Double
    Public ExRate(,,,,,) As Double
    Public ExRateAll(,,,) As Double
    Public ExplScale(,) As Double
    Public ExternalModelStockProportion() As Single
    Public FVSdatabasepath As String
    Public FVSdatabasename As String
    Public FirstFish As Integer
    Public Firstpass As Boolean
    Public FirstStk As Integer
    Public Fish As Integer
    Public Fisha As Integer
    Public fishb() As String
    Public FishCohort As Double
    Public FishExpCWTAll(,,,) As Double
    Public FishFlag() As Integer
    Public FishName() As String
    Public FisheryNum As String
    Public FishNum As Integer
    Public FishForm As String
    Public FishScalar(,,) As Double
    Public FishYear As Integer
    Public FullName As String
    Public findme As String
    Public FileLength As Integer
    Public filepath As String
    Public Group As String
    Public ImputeFish(,,,) As Integer
    Public ImputeListMain As New List(Of CWTData)
    Public imputerecoveries As CWTData
    Public imputerecoveriesmain As CWTData
    Public InFileName As String
    Public InFileName2 As String
    Public InFilePath As String
    Public InpFile As String
    Public Iread As Integer
    Public Ireadstk As Integer
    Public irow As Integer
    Public junk As String
    Public K(,)
    Public L(,)
    Public LabelMe As String
    Public LegalProp As Double
    Public Lengtha As Integer
    Public LowerLim As Double
    Public MatCatch As Double
    Public MatRate(,,) As Double
    Public MaxAge As Integer
    Public MaxAgeEncAdj As Integer
    Public MaxAgeERAdj As Integer
    Public MaxBY As Integer
    Public MaxNAge As Integer = 5
    Public MaxNFish As Integer = 3
    Public MaxNSteps As Integer = 4
    Public MaxNStk As Integer = 2
    Public MaxStk As Integer
    Public Mean As Double
    Public MergedCWT(,,,) As Double
    Public MinBY As Integer
    Public MinSize(,,)
    Public MinStk As Integer
    Public MixedCatch As Double
    Public ModelStkPPN() As Double
    Public MonthNames() As String
    Public MortRate As String
    Public mynumber As Double
    Public NameLength As Integer
    Public NewFile As String
    Public NewSize As Integer
    Public NoExpansions As Boolean
    Public NumCNR As String
    Public NumFish As Integer
    Public NumImpute As String
    Public NumImputeER As String
    Public NumNewSize As String
    Public NumNonZero As Double
    Public NumOfStks As String
    Public NumSteps As Integer
    Public NumStk As Integer
    Public NumTerm As String
    Public numyears As Integer
    Public ObsEscpmnt() As Double
    Public OldBPSurrogateFishCheck As Boolean '##### JC Update; 9/25/2015 ####
    Public OtherMort(,) As Double
    Public OOBStatus As Integer
    Public oobTableName As String
    Public Outfile As String
    Public PntTime() As Single
    Public PointFish() As Double
    Public PointStk() As Integer
    Public PropLegCatch(,,) As Double
    Public PropSubCatch(,,) As Double
    Public PropLeg As Double
    Public PropSubPop(,) As Double
    Public PTerm As Integer = 0
    Public PropTNet(,) As Double
    Public Rmax() As Double
    Public Record As String
    Public Recorda As String
    Public RecordsetSelectionType As Integer
    Public Recoveries As Double
    Public RecAdjFactor() As Double
    Public RunID As Integer
    Public ScaleFile As String
    Public SConStr As String
    Public Scratch As String
    Public SD As Double
    Public SelAge As Integer
    Public SeqAge As Integer
    Public SelectedBasePeriodID As Integer
    Public ShakDistrMeth As Integer
    Public Shaker(,,) As Double ' age, fishery, timestep
    Public ShakerAll(,,,) As Double
    Public ShakEncRate As Double
    Public ShakMortRate(,) As Double
    Public SizeLimit(,) As Integer
    Public Stage As Integer
    Public StageNum As Integer
    Public Start As String
    Public StartCohort(,) As Double
    Public STk As Integer
    Public Stknum As Integer
    Public StockCatch(,) As Double
    Public StockCatchProp(,) As Double
    Public StkCheck(,) As Double
    Public StkFishCatch() As Double
    Public Stkform As String
    Public stock() As String
    Public stringa As String
    Public stringescape(,,) As String
    Public SubLegProp As Double
    Public SumCWT As Double
    Public SupportFile As String
    Public SurrogateFishBP_ER(,,,) As Double '##### JC Update; 9/25/2015 ####
    Public SurvRate(,) As Double
    Public T0(,)
    Public TargEncRate(,) As Double
    Public TransDBName As String
    Public TransferDataSet As New System.Data.DataSet()
    Public TransferBPLongName As String
    Public TempCatch As Double
    Public TempConc As Double
    Public TempExRate As Double
    Public TempPoint As Double
    Public Tempu As Double
    Public Term As Integer = 1
    Public TermFish As String
    Public TermFlag(,) As Integer
    Public TermStat As Integer
    Public Time1 As Double
    Public TimeCatch(,) As Double
    Public Title As String
    Public TotalCNR(,) As Double
    Public TotCatch(,) As Double
    Public TotLanded As Double
    Public TotLandedItem As StkBYCWTSumData
    Public TotLandedItem2 As StkBYCWTSumData
    Public TotEnc As Double
    Public TotExRate As Double
    Public TotExpCWTAll(,,,) As Double
    Public TotalShakerTerm(,) As Double
    Public TotalShakerPTerm(,) As Double
    Public TotalStk() As Double
    Public TrueCatch() As Double
    Public TotalMort() As Double
    Public TotMort() As Double
    Public TotMortTerm(,) As Double
    Public TotMortPterm(,) As Double
    Public TStep As Integer
    Public TransferBPName As String
    Public SubLegalProp As Double
    Public TotSubLegalPop As Double
    Public u() As Double
    Public WeightItem As WeightData
    Public WeightList As New List(Of WeightData)
    Public Yr As Integer
    Public Z As Double
    Public ImputeItem As ImputeData
    Public CWTRow As CWTData
    Public ImputeList As New List(Of ImputeData)
    Public n As Integer
    Public sublist As New List(Of CWTData)
    Public TotalLandedList As New List(Of StkBYCWTSumData)
    Public CWTList As New List(Of CWTData)
    Public RecordCWT As CWTData
    Public errpath As String
    Public Indextracker() As Integer
End Module
