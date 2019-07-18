Imports System
Imports InfiniExpress.BLL

Public Class LedgerSummary

    Inherits System.Windows.Forms.Form
    Dim GlbAccount As New ClassAccount()

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
        MDI.numOfForms = 0
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(LedgerSummary))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.DataGrid1})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(488, 260)
        Me.Panel1.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(10, 231)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(65, 22)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Close"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.DataGrid1.CaptionText = "Accounts Summary"
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.DataGrid1.Location = New System.Drawing.Point(8, 10)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.White
        Me.DataGrid1.Size = New System.Drawing.Size(472, 214)
        Me.DataGrid1.TabIndex = 3
        '
        'LedgerSummary
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(496, 271)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LedgerSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Ledger Accounts Summary"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub LedgerSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LedgerSummaryPopulate()
    End Sub

    Sub LedgerSummaryPopulate()

        Dim objDT As New DataTable()
        Dim objDS As New DataSet()

        objDS = GlbAccount.LedgerSummary

        Dim ts As New DataGridTableStyle()

        ' Add the custom column style.
        Dim cs As New DataGridTextBoxColumn()
        cs.Width = 230
        cs.MappingName = "tname"        ' Map to decimal column.
        cs.HeaderText = "Account Description"
        cs.ReadOnly = True
        ts.GridColumnStyles.Add(cs)

        ' Add the standard column style.
        cs = New DataGridTextBoxColumn()
        cs.Width = 110
        cs.MappingName = "TDR"          ' Map to integer column.
        cs.HeaderText = "Debit " & Chr(9)
        cs.Format = "0.00"
        cs.Alignment = HorizontalAlignment.Right
        cs.TextBox.TextAlign = HorizontalAlignment.Right
        'cs.TextBox.Text = Format(cs.TextBox.Text, "0.00")
        cs.ReadOnly = True
        ts.GridColumnStyles.Add(cs)

        ' Add the standard column style.
        cs = New DataGridTextBoxColumn()
        cs.Width = 110
        cs.MappingName = "TCR"           ' Map to integer column.
        cs.HeaderText = "Credit " & Chr(9)
        cs.Format = "0.00"
        cs.Alignment = HorizontalAlignment.Right
        cs.TextBox.TextAlign = HorizontalAlignment.Right
        'cs.TextBox.Text = Format(cs.TextBox.Text, "0.00")
        cs.ReadOnly = True
        ts.GridColumnStyles.Add(cs)
        ts.MappingName = objDS.Tables(0).TableName     ' Map table style to TestTable.

        ts.RowHeadersVisible = False
        ts.PreferredRowHeight = 3

        DataGrid1.TableStyles.Add(ts)
        DataGrid1.DataSource = objDS
        DataGrid1.DataMember = objDS.Tables(0).TableName
        'DataGrid1.Enabled = False

        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose(True)
    End Sub
End Class

