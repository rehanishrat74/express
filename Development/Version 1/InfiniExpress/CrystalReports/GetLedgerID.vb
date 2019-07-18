Imports InfiniExpress.BLL
Public Class GetLedgerID
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
        MDI.numOfForms = 0
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbLedgerID As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(GetLedgerID))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbLedgerID = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Button1, Me.Label1, Me.cmbLedgerID})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(324, 150)
        Me.Panel1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(124, 114)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(74, 18)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Get Report"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Ledger Name"
        '
        'cmbLedgerID
        '
        Me.cmbLedgerID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLedgerID.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbLedgerID.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.cmbLedgerID.Location = New System.Drawing.Point(100, 62)
        Me.cmbLedgerID.Name = "cmbLedgerID"
        Me.cmbLedgerID.Size = New System.Drawing.Size(176, 21)
        Me.cmbLedgerID.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(296, 32)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "In this report only those ledger name shows which relate with transaction"
        '
        'GetLedgerID
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(332, 159)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GetLedgerID"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "GetLedgerID"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim myds As New DataSet()
    Dim GlbReports As New ClassReports()
    Private Sub GetLedgerID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myds = GlbReports.GetLedgerID()
        cmbLedgerID.DataSource = myds.Tables(0)
        cmbLedgerID.DisplayMember = "Name"
        cmbLedgerID.ValueMember = "LedgerID"
        'cmbLedgerID.DisplayMember = "LedgerID"
        'cmbLedgerID.ValueMember = "LedgerID"
    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ShowReport.LedgerId = CStr(cmbLedgerID.SelectedValue)
        ShowReport.LedgerName = cmbLedgerID.Text
        Dim objSR As New ShowReport()
        Me.Dispose(True)
        objSR.Show()
        'MDI.numOfForms = 1
    End Sub

End Class
