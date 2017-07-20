Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO

Public Class SetWeight

    'Dim objConnection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & SupportFile)
    Dim objConnection As New OleDbConnection(SConStr)
    Dim objdataadapter As New OleDbDataAdapter()
    Dim objdataset As New DataSet()




    Private Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        objdataadapter.SelectCommand = New OleDbCommand()
        objdataadapter.SelectCommand.Connection = objConnection
        objdataadapter.SelectCommand.CommandText = _
            "SELECT DataPkey, StockID, BroodYear, Stage, Weight, Flag " & _
            "FROM Weighting "
        objdataadapter.SelectCommand.CommandType = CommandType.Text
        objConnection.Open()

        objdataadapter.Fill(objdataset, "Weighting")
        'objConnection.Close()
        grdWeights.AutoGenerateColumns = True
        grdWeights.DataSource = objdataset
        grdWeights.DataMember = "Weighting"

        'objdataadapter = Nothing
        'objConnection = Nothing


    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(objdataadapter)

        ' modify data in DataSet and commit changes

        builder.GetUpdateCommand()

        ' Without the OleDbCommandBuilder this line would fail.
        objdataadapter.Update(objdataset, "Weighting")
        objConnection.Close()
        objdataadapter = Nothing
        objConnection = Nothing
        Me.Close()

       

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
         
        objdataadapter.FillSchema(objdataset, SchemaType.Mapped)
        objdataadapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        objdataadapter.Fill(objdataset, "Weighting")
        grdWeights.Refresh()



    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click


        objConnection.Close()
        objdataadapter = Nothing
        objConnection = Nothing
        Me.Close()

    End Sub

    Private Sub ToolTip1_Popup(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PopupEventArgs) Handles ToolTip1.Popup

    End Sub
End Class