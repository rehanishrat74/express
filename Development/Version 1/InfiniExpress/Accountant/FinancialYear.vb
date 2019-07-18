Imports System
Imports InfiniExpress.BLL

Public Class FinancialYear

    Inherits System.Windows.Forms.Form
    Dim GlbAccount As New ClassAccount()
    Dim GlbGlobal As New ClassGlobal()

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
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FinancialYear))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cbMonth = New System.Windows.Forms.ComboBox()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbMonth, Me.txtYear, Me.Label5, Me.Label4, Me.btnBack, Me.btnSave, Me.Label3})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(344, 192)
        Me.Panel1.TabIndex = 0
        '
        'cbMonth
        '
        Me.cbMonth.BackColor = System.Drawing.Color.White
        Me.cbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbMonth.Location = New System.Drawing.Point(124, 88)
        Me.cbMonth.Name = "cbMonth"
        Me.cbMonth.Size = New System.Drawing.Size(116, 21)
        Me.cbMonth.TabIndex = 0
        '
        'txtYear
        '
        Me.txtYear.BackColor = System.Drawing.Color.White
        Me.txtYear.Location = New System.Drawing.Point(124, 120)
        Me.txtYear.MaxLength = 4
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(68, 20)
        Me.txtYear.TabIndex = 1
        Me.txtYear.Text = ""
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label5.Location = New System.Drawing.Point(20, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 15)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Financial Year"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(20, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Financial Month"
        '
        'btnBack
        '
        Me.btnBack.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnBack.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnBack.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.ForeColor = System.Drawing.Color.White
        Me.btnBack.Location = New System.Drawing.Point(92, 160)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(60, 18)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "&Close"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(22, 160)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(60, 18)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "&Save"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(20, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(308, 56)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "You can select an appropriate month and year to determine the start of the compan" & _
        "y's financial year. Use the pull down menu to the select the month and enter the" & _
        " year in the field provided. "
        '
        'FinancialYear
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(352, 201)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FinancialYear"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Financial Year"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FinancialYear_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ChkTrans As Boolean, FYear As String = ""
        cbMonth.Items.Clear()
        Dim Reader As IDataReader = GlbAccount.FinancialYear

        While Reader.Read
            cbMonth.Items.Add(Trim(CStr(Reader("FMonth"))))
            If FYear = "" Then
                FYear = Trim(CStr(Reader("FYear")))
            End If
        End While
        Reader.Close()
        cbMonth.SelectedIndex = 0
        txtYear.Text = FYear
        ChkTrans = GlbAccount.CheckFinancialYearTransaction
        If ChkTrans = True Then
            txtYear.ReadOnly = True
            cbMonth.Enabled = False
            btnSave.Enabled = False
            txtYear.Enabled = False
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim FinalYear As String, Fyear As String
        Dim CurrentYear As String, Nextyear As String
        Dim FinalMonth As String = Trim(CType(cbMonth.SelectedItem, String))

        If txtYear.Enabled = False Then
            Exit Sub
        End If
        If IsNumeric(txtYear.Text) = False Then
            MessageBox.Show("Please enter a correct financial year.     ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If Val(txtYear.Text) >= 2000 And Val(txtYear.Text) <= 2050 Then
            FinalYear = txtYear.Text
            Fyear = txtYear.Text
            CurrentYear = Fyear
            Nextyear = Trim(Str(Val(CurrentYear) + 1))
        Else
            MessageBox.Show("Your financial year must be in between 2000 & 2050.      ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        GlbAccount.AddFinancialYear(FinalMonth, Fyear, Nextyear)

        If GlbGlobal.ExecuteCommand() = True Then
            MessageBox.Show("Financial year has been updated.       ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.Dispose(True)
        End If
    End Sub

End Class
