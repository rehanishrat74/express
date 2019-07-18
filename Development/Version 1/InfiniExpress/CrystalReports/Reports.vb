Public Class Reports
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RadioButton11 As System.Windows.Forms.RadioButton
    Friend WithEvents RBCustomerList As System.Windows.Forms.RadioButton
    Friend WithEvents RBCustomerActivity As System.Windows.Forms.RadioButton
    Friend WithEvents RBVendorList As System.Windows.Forms.RadioButton
    Friend WithEvents RBVendorActivity As System.Windows.Forms.RadioButton
    Friend WithEvents RBPurchaseInvoices As System.Windows.Forms.RadioButton
    Friend WithEvents RBLedgerBalance As System.Windows.Forms.RadioButton
    Friend WithEvents RBLedgerAcitivity As System.Windows.Forms.RadioButton
    Friend WithEvents RBPurchasePayments As System.Windows.Forms.RadioButton
    Friend WithEvents RBSalesReceipts As System.Windows.Forms.RadioButton
    Friend WithEvents RBSalesInvoice As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RBPurchasePayments = New System.Windows.Forms.RadioButton()
        Me.RBSalesReceipts = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.RBLedgerAcitivity = New System.Windows.Forms.RadioButton()
        Me.RBLedgerBalance = New System.Windows.Forms.RadioButton()
        Me.RBPurchaseInvoices = New System.Windows.Forms.RadioButton()
        Me.RBVendorActivity = New System.Windows.Forms.RadioButton()
        Me.RBVendorList = New System.Windows.Forms.RadioButton()
        Me.RBSalesInvoice = New System.Windows.Forms.RadioButton()
        Me.RBCustomerActivity = New System.Windows.Forms.RadioButton()
        Me.RBCustomerList = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RadioButton11 = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.RBPurchasePayments, Me.RBSalesReceipts, Me.Button1, Me.RBLedgerAcitivity, Me.RBLedgerBalance, Me.RBPurchaseInvoices, Me.RBVendorActivity, Me.RBVendorList, Me.RBSalesInvoice, Me.RBCustomerActivity, Me.RBCustomerList, Me.Label3, Me.Label2, Me.Label1, Me.RadioButton11})
        Me.GroupBox1.Location = New System.Drawing.Point(32, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(488, 264)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reports"
        '
        'RBPurchasePayments
        '
        Me.RBPurchasePayments.Location = New System.Drawing.Point(192, 168)
        Me.RBPurchasePayments.Name = "RBPurchasePayments"
        Me.RBPurchasePayments.Size = New System.Drawing.Size(128, 24)
        Me.RBPurchasePayments.TabIndex = 13
        Me.RBPurchasePayments.Text = "Purchase Payments"
        '
        'RBSalesReceipts
        '
        Me.RBSalesReceipts.Location = New System.Drawing.Point(16, 168)
        Me.RBSalesReceipts.Name = "RBSalesReceipts"
        Me.RBSalesReceipts.TabIndex = 12
        Me.RBSalesReceipts.Text = "Sales Receipts"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(192, 216)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "View Report"
        '
        'RBLedgerAcitivity
        '
        Me.RBLedgerAcitivity.Location = New System.Drawing.Point(352, 104)
        Me.RBLedgerAcitivity.Name = "RBLedgerAcitivity"
        Me.RBLedgerAcitivity.TabIndex = 10
        Me.RBLedgerAcitivity.Text = "Ledger Acitivity"
        '
        'RBLedgerBalance
        '
        Me.RBLedgerBalance.Location = New System.Drawing.Point(352, 72)
        Me.RBLedgerBalance.Name = "RBLedgerBalance"
        Me.RBLedgerBalance.TabIndex = 9
        Me.RBLedgerBalance.Text = "Ledger Balance"
        '
        'RBPurchaseInvoices
        '
        Me.RBPurchaseInvoices.Location = New System.Drawing.Point(192, 136)
        Me.RBPurchaseInvoices.Name = "RBPurchaseInvoices"
        Me.RBPurchaseInvoices.Size = New System.Drawing.Size(120, 24)
        Me.RBPurchaseInvoices.TabIndex = 8
        Me.RBPurchaseInvoices.Text = "Purchase Invoices"
        '
        'RBVendorActivity
        '
        Me.RBVendorActivity.Cursor = System.Windows.Forms.Cursors.Default
        Me.RBVendorActivity.Location = New System.Drawing.Point(192, 104)
        Me.RBVendorActivity.Name = "RBVendorActivity"
        Me.RBVendorActivity.TabIndex = 7
        Me.RBVendorActivity.Text = "Vendor Activity"
        '
        'RBVendorList
        '
        Me.RBVendorList.Location = New System.Drawing.Point(192, 72)
        Me.RBVendorList.Name = "RBVendorList"
        Me.RBVendorList.TabIndex = 6
        Me.RBVendorList.Text = "Vendor List"
        '
        'RBSalesInvoice
        '
        Me.RBSalesInvoice.Location = New System.Drawing.Point(16, 136)
        Me.RBSalesInvoice.Name = "RBSalesInvoice"
        Me.RBSalesInvoice.TabIndex = 5
        Me.RBSalesInvoice.Text = "Sales Invoice"
        '
        'RBCustomerActivity
        '
        Me.RBCustomerActivity.Location = New System.Drawing.Point(16, 104)
        Me.RBCustomerActivity.Name = "RBCustomerActivity"
        Me.RBCustomerActivity.Size = New System.Drawing.Size(120, 24)
        Me.RBCustomerActivity.TabIndex = 4
        Me.RBCustomerActivity.Text = "Customer Activity"
        '
        'RBCustomerList
        '
        Me.RBCustomerList.Location = New System.Drawing.Point(16, 72)
        Me.RBCustomerList.Name = "RBCustomerList"
        Me.RBCustomerList.TabIndex = 3
        Me.RBCustomerList.Text = "Customer List"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(360, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Accountant"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(184, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Purchase Reports"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "   Sales Reports"
        '
        'RadioButton11
        '
        Me.RadioButton11.Location = New System.Drawing.Point(16, 72)
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.TabIndex = 3
        Me.RadioButton11.Text = "Customer List"
        '
        'Reports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 309)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox1})
        Me.Location = New System.Drawing.Point(135, 35)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Reports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Reports"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If RBCustomerList.Checked = True Then
            ShowReport.ReportName = "CustomerList"
            Dim SR As New ShowReport()
            SR.Show()
            Me.Hide()
            SR.Activate()
            Me.Dispose(True)
        End If

        If RBCustomerActivity.Checked = True Then
            ShowReport.ReportName = "CustomerActivity"
            Dim CidforActivity As New GetCustomerID()
            CidforActivity.Show()
            Me.Hide()
            CidforActivity.Activate()
            Me.Dispose(True)
        End If


        If RBSalesInvoice.Checked = True Then
            ShowReport.ReportName = "SalesInvoice"
            Dim CidforActivity As New GetCustomerID()
            CidforActivity.Show()
            Me.Hide()
            CidforActivity.Activate()
            Me.Dispose(True)
        End If

        If RBSalesReceipts.Checked = True Then
            ShowReport.ReportName = "SalesReceipts"
            Dim CidforActivity As New GetCustomerID()
            CidforActivity.Show()
            Me.Hide()
            CidforActivity.Activate()
            Me.Dispose(True)
        End If

        If RBVendorList.Checked = True Then
            ShowReport.ReportName = "VendorList"
            Dim SR As New ShowReport()
            SR.Show()
            Me.Hide()
            sr.Activate()
            Me.Dispose(True)
        End If

        If RBVendorActivity.Checked = True Then
            ShowReport.ReportName = "VendorActivity"
            Dim VendorId As New GetVendorID()
            VendorId.Show()
            Me.Hide()
            VendorId.Activate()
            Me.Dispose(True)
        End If

        If RBPurchaseInvoices.Checked = True Then
            ShowReport.ReportName = "PurchaseInvoices"
            Dim VendorId As New GetVendorID()
            VendorId.Show()
            Me.Hide()
            VendorId.Activate()
            Me.Dispose(True)
        End If

        If RBPurchasePayments.Checked = True Then
            ShowReport.ReportName = "PurchasePayments"
            Dim VendorId As New GetVendorID()
            VendorId.Show()
            Me.Hide()
            VendorId.Activate()
            Me.Dispose(True)
        End If

        If RBLedgerBalance.Checked = True Then
            ShowReport.ReportName = "LedgerBalance"
            Dim SR As New ShowReport()
            SR.Show()
            Me.Hide()
            sr.Activate()
            Me.Dispose(True)
        End If

        If RBLedgerAcitivity.Checked = True Then
            ShowReport.ReportName = "LedgerActivity"
            Dim LedgerId As New GetLedgerID()
            LedgerId.Show()
            Me.Hide()
            LedgerId.Activate()
            Me.Dispose(True)
        End If

    End Sub
End Class
