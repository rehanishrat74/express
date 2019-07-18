Public Class CustomerIDForActivity
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
    Friend WithEvents cmbCustID As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbCustID = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmbCustID
        '
        Me.cmbCustID.Location = New System.Drawing.Point(144, 56)
        Me.cmbCustID.Name = "cmbCustID"
        Me.cmbCustID.Size = New System.Drawing.Size(121, 21)
        Me.cmbCustID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Customer ID"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(160, 128)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Get Report"
        '
        'CustomerIDForActivity
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(408, 213)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.Label1, Me.cmbCustID})
        Me.Location = New System.Drawing.Point(135, 35)
        Me.Name = "CustomerIDForActivity"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "CustomerIDForActivity"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim myds As New DataSet()
    Private Sub CustomerIDForActivity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        myds = ReportQueries.GetCustID()
        cmbCustID.DataSource = myds.Tables(0)
        cmbCustID.DisplayMember = "CustomerID"
        cmbCustID.ValueMember = "CustomerID"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ShowReport.CustId = CStr(cmbCustID.SelectedValue)
        Dim objSR As New ShowReport()
        objSR.Show()
        Me.Hide()
    End Sub
End Class
