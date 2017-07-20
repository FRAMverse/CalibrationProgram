Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Public Class SetPpnTreaty
    Dim objConnection3 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & SupportFile)
    Dim objdataadapter3 As New OleDbDataAdapter()
    Dim objdataset3 As New DataSet()
    Private Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        objdataadapter3.SelectCommand = New OleDbCommand()
        objdataadapter3.SelectCommand.Connection = objConnection3
        objdataadapter3.SelectCommand.CommandText = _
            "SELECT ppnTreatyKey, FisheryID, TimeStepID, PpnTreaty " & _
            "FROM ProportionTreatyNet "
        objdataadapter3.SelectCommand.CommandType = CommandType.Text
        objConnection3.Open()

        objdataadapter3.Fill(objdataset3, "ProportionTreatyNet")
        'objConnection.Close()
        grdPpnTreaty.AutoGenerateColumns = True
        grdPpnTreaty.DataSource = objdataset3
        grdPpnTreaty.DataMember = "ProportionTreatyNet"

        'objdataadapter = Nothing
        'objConnection = Nothing


    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(objdataadapter3)

        ' Code to modify data in DataSet here 

        builder.GetUpdateCommand()

        ' Without the OleDbCommandBuilder this line would fail.
        objdataadapter3.Update(objdataset3, "ProportionTreatyNet")
        objConnection3.Close()
        objdataadapter3 = Nothing
        objConnection3 = Nothing
        Me.Close()



    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        objdataadapter3.FillSchema(objdataset3, SchemaType.Mapped)
        objdataadapter3.MissingSchemaAction = MissingSchemaAction.AddWithKey

        objdataadapter3.Fill(objdataset3, "ProportionTreatyNet")
        grdPpnTreaty.Refresh()



    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click


        objConnection3.Close()
        objdataadapter3 = Nothing
        objConnection3 = Nothing
        Me.Close()

    End Sub
End Class