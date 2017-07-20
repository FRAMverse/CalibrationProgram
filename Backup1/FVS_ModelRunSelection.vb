Imports System
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.IO.File
Imports System.Data.OleDb
Imports System.Data

Public Class FVS_ModelRunSelection

   '- Run List Selection Variables
    Public Shared RunID(150) As Integer
    Public Shared BaseID(50) As Integer
   Public Shared RunIDName(150) As String
   Public Shared RunBasePeriodID(150) As Integer
   Public PrnLine, PrnPart As String
    Public rssw As StreamWriter
    Public iresult


    
   Public Sub FillRunList()
        If Exists(FVSdatabasename) Then
        Else
            'Open calibration database
            OpenFileDialog3.Title = "Select the Calibration Database File"
            OpenFileDialog3.Filter = "DataBase Files(*.mdb;*.accdb)|*.MDB;*.ACCDB"
            'Set the CheckFileExists property so that a warning 
            '    appears if the user types a filename of a non-existent 
            '    file.  Then show the dialog.
            OpenFileDialog3.CheckFileExists = True
            MsgBox("Please select the calibration support database.")
            iresult = OpenFileDialog3.ShowDialog()


            'Make sure the user did not click the Cancel button And 
            '    specified a file name for the file to be created.  
            If iresult <> Windows.Forms.DialogResult.Cancel And _
                        OpenFileDialog3.FileName.Length <> 0 Then
                FVSdatabasename = OpenFileDialog3.FileName
                             
            End If
        End If


        CalibrationDB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FVSdatabasename & ";"
        CalibrationDB.Open()
        Dim drd1 As OleDb.OleDbDataReader
        Dim cmd1 As New OleDb.OleDbCommand()
        ReDim BaseID(100)
        cmd1.Connection = CalibrationDB

        cmd1.CommandText = "SELECT * FROM SummaryInfo ORDER BY BasePeriodID"
        
        drd1 = cmd1.ExecuteReader
        
        Dim int1 As Integer
        int1 = 0
        ListBox1.Items.Clear()
        'If drd1.Read() <> False Then
        Do While drd1.Read

            '- Fill CheckedListBox Items


            BaseID(int1) = drd1.GetValue(9)
            ListBox1.Items.Add(BaseID(int1) & "                BasePeriodID")
            '- Set RunID Array Values

            'RunBasePeriodID(int1) = drd1.GetInt32(5)
            'RunIDName(int1) = drd1.GetString(3)
            int1 = int1 + 1

        Loop
        CalibrationDB.Close()
        'End If
    End Sub

   Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
      Dim ListIndex, Result As Integer
      ListIndex = CInt(ListBox1.SelectedIndex.ToString())
      If ListIndex = -1 Then
         MsgBox("ERROR - You selected a blank line" & vbCrLf & "Please try again!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      

   End Sub

   Private Sub CmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdCancel.Click
     
        RecordsetSelectionType = 9
        Me.Close()
        StartForm.Visible = True

   End Sub

   


    Private Sub FVS_ModelRunSelection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        RSTitle.Text = "Model Run TRANSFER Selections"
        ListBox1.SelectionMode = SelectionMode.MultiExtended
        TransferButton.Visible = True

        Me.BringToFront()
        FillRunList()
    End Sub


    Private Sub TransferButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransferButton.Click
        Dim SelectedBasePeriodIDString As String


        SelectedBasePeriodIDString = ListBox1.SelectedItem
        SelectedBasePeriodIDString = Mid(SelectedBasePeriodIDString, 1, 5)
        SelectedBasePeriodID = CInt(SelectedBasePeriodIDString)

        Me.Close()
        StartForm.Visible = True
    End Sub
End Class