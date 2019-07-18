Imports System
Imports InfiniExpress.BLL

Public Class TaxCode

    Inherits System.Windows.Forms.Form
    Dim GlbAccount As New ClassAccount()
    Dim ClsGlobal As New ClassGlobal()

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
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dgrdTaxRates As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TaxCode))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgrdTaxRates = New System.Windows.Forms.DataGrid()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.dgrdTaxRates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgrdTaxRates, Me.Label2, Me.Button2, Me.btnSave})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(364, 316)
        Me.Panel1.TabIndex = 0
        '
        'dgrdTaxRates
        '
        Me.dgrdTaxRates.AllowNavigation = False
        Me.dgrdTaxRates.AllowSorting = False
        Me.dgrdTaxRates.BackgroundColor = System.Drawing.Color.White
        Me.dgrdTaxRates.CaptionText = "Tax Codes Summary"
        Me.dgrdTaxRates.CaptionVisible = False
        Me.dgrdTaxRates.DataMember = ""
        Me.dgrdTaxRates.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgrdTaxRates.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.dgrdTaxRates.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgrdTaxRates.Location = New System.Drawing.Point(14, 68)
        Me.dgrdTaxRates.Name = "dgrdTaxRates"
        Me.dgrdTaxRates.ParentRowsBackColor = System.Drawing.Color.White
        Me.dgrdTaxRates.RowHeadersVisible = False
        Me.dgrdTaxRates.Size = New System.Drawing.Size(334, 208)
        Me.dgrdTaxRates.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(336, 64)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "The list of tax rates and codes appear below according to default settings. You m" & _
        "ay change these simply by modifying the appropriate fields and clicking the Save" & _
        " button. "
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(82, 284)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(60, 18)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(12, 284)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(60, 18)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        '
        'TaxCode
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(372, 325)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TaxCode"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Tax Codes Summary"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgrdTaxRates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub TaxCode_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim objDS As DataSet
        objDS = New DataSet()
        objDS = ClsGlobal.TaxCodeSummary

        Dim cs As DataGridTextBoxColumn
        Dim ts = New DataGridTableStyle()
        Dim lastwidth As Integer

        cs = New DataGridTextBoxColumn()
        cs.HeaderText = "Tax ID"
        cs.MappingName = "TaxId"
        cs.Width = dgrdTaxRates.Width / 3 + 18
        lastwidth = cs.Width
        cs.ReadOnly = True
        ts.GridColumnStyles.Add(cs)

        cs = New DataGridTextBoxColumn()
        cs.HeaderText = "Tax Rate"
        cs.MappingName = "TaxRate"
        cs.Format = "0.00"
        cs.Width = 198
        cs.ReadOnly = False

        ts.MappingName = objDS.Tables(0).TableName
        ts.GridColumnStyles.Add(cs)
        ts.RowHeadersVisible = False

        dgrdTaxRates.TableStyles.Add(ts)
        dgrdTaxRates.DataSource = objDS
        dgrdTaxRates.DataMember = objDS.Tables(0).TableName
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(dgrdTaxRates.DataSource, dgrdTaxRates.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

    End Sub

    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim r As Integer
        Dim TCode As String, TRate As Double
        Dim TaxId As TextBox
        Dim objDS As New DataSet()
        Dim dgc As New DataGridCell()

        objDS = ClsGlobal.TaxCodeSummary

        For r = 0 To objDS.Tables(0).Rows.Count - 1 '9          'row    'column

            TCode = Convert.ToString(dgrdTaxRates.Item(r, 0))
            If Not IsNumeric(dgrdTaxRates.Item(r, 1)) Then
                MessageBox.Show("Tax rate for T" & r & " is  not valid.     ", "Invalid Value", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If Convert.ToDouble(dgrdTaxRates.Item(r, 1)) < 0.0 Or Convert.ToDouble(dgrdTaxRates.Item(r, 1)) > 99.99 Then
                MessageBox.Show("Tax rate for T" & r & " is  not valid. Values can range from 0.00 to 99.99 only.      ", "Invalid Value", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If dgrdTaxRates.Item(r, 1) <> objDS.Tables(0).Rows(r)(2) Then ' if changed !!
                TRate = Convert.ToDouble(dgrdTaxRates.Item(r, 1))
                GlbAccount.UpdateTaxCodeSummary(TCode, TRate)
            End If

        Next r

        If ClsGlobal.ExecuteCommand() = True Then
            MessageBox.Show("Tax code rates have been updated.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        Me.Dispose(True)
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose(True)
    End Sub

End Class
