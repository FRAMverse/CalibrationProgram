Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Public Class SetFlags
    Dim objConnection2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & SupportFile)
    Dim objdataadapter2 As New OleDbDataAdapter()
    Dim objdataset2 As New DataSet()

    Private Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        objdataadapter2.SelectCommand = New OleDbCommand()
        objdataadapter2.SelectCommand.Connection = objConnection2
        objdataadapter2.SelectCommand.CommandText = _
            "SELECT FisheryFlagKey, FisheryID, FishFlag " & _
            "FROM FisheryFlag "
        objdataadapter2.SelectCommand.CommandType = CommandType.Text
        objConnection2.Open()

        objdataadapter2.Fill(objdataset2, "FisheryFlag")
        'objConnection.Close()
        grdFlags.AutoGenerateColumns = True
        grdFlags.DataSource = objdataset2
        grdFlags.DataMember = "FisheryFlag"

        'objdataadapter = Nothing
        'objConnection = Nothing


    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(objdataadapter2)

        ' Code to modify data in DataSet here 

        builder.GetUpdateCommand()

        ' Without the OleDbCommandBuilder this line would fail.
        objdataadapter2.Update(objdataset2, "FisheryFlag")
        objConnection2.Close()
        objdataadapter2 = Nothing
        objConnection2 = Nothing
        Me.Close()



    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        objdataadapter2.FillSchema(objdataset2, SchemaType.Mapped)
        objdataadapter2.MissingSchemaAction = MissingSchemaAction.AddWithKey

        objdataadapter2.Fill(objdataset2, "FisheryFlag")
        grdFlags.Refresh()



    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click


        objConnection2.Close()
        objdataadapter2 = Nothing
        objConnection2 = Nothing
        Me.Close()

    End Sub



End Class