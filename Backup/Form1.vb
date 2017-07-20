Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO

Public Class CHDAT
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CHDAT))
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(200, 273)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(73, 27)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "START"
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.BackColor = System.Drawing.Color.Chocolate
        Me.Label1.Font = New System.Drawing.Font("Script MT Bold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(50, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(437, 204)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CHDAT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 357)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "CHDAT"
        Me.Text = "CHDAT"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '**********************************************************************
        'Program Name: CHDAT14.BAS
        '**********************************************************************
        'February 16, 1996
        'Version 14.0
        'Jim Scott, NWIFC
        '03/05/95 - JBS removed many arrays to reduce memory requirements and
        '           modified the way shaker mortality rates are read.  Allowed
        '           for multiple ages of encounter rate adjustments.
        '12/07/95 - Modified read and write of size limits for consistency with
        '           cohort analysis.
        '02/16/96 - (1) Modified read and write of recovery adjustment flags for
        '           consistency with cohort analysis.
        '           (2) Modified subroutine Impute to allow exploitation rates
        '           in any fishery to be duplicated in a second fishery.  Flags
        '           are read from file .CHK.
        '02/06/98 - (1) Modified read and write of recovery adjustment flags for
        '           consistency with cohort analysis.
        '02/26/98 - Modified to read and convert new time sequencing.
        '01/07/00 - read data to impute catches from external Exp rates and pass through
        '           to CHCAL  Simmons
        '07/28/03 - deactivated print statement for Simmons imputed catches sub
        '8/21/07 - AHB; converted to VB
        '           included ability to select multiple chk files
        '**********************************************************************

        '**********************************************************************************
        'Select CHK files
        
        Me.Hide()
        MsgBox("Please select the Access Calibration Support Database by selecting OPEN SUPPORT FILE")
        mainform.Show()

        
    End Sub
End Class
