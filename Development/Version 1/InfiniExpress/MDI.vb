Imports vbAccelerator.Components.Controls
Imports System.IO
Imports System.Threading
Imports InfiniExpress.BLL

Public Class MDI

    Inherits System.Windows.Forms.Form

#Region "Local Variables"
    Public Shared numOfForms As Integer
    Dim ExpLogo As New LogoForm()
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Changing BackColor Of MDI Forms
        Dim c As Control
        For Each c In Me.Controls
            If TypeOf c Is MdiClient Then
                'c.BackColor = Color.CornflowerBlue
                c.BackColor = Color.FromArgb(102, 153, 204)
                Exit For
            End If
        Next

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        Dim mResponse As Integer
        mResponse = MessageBox.Show("Please Synchronize your Database before closing the application.    " & vbNewLine & "Whould you like to synchronize.", "Infini Accounts Express", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If mResponse = 6 Then
            Dim child As New FrmBar
            child.MdiParent = Me
            child.Show()
            numOfForms = 1
        Else
            MyBase.Dispose(disposing)
            Application.Exit()
        End If

    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCreateCustomer As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSalesInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSalesCreditNote As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemReceipts As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCreateVendor As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorSummary As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorAmend As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemPurchaseInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemPurchaseCreditNote As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorPayment As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorActivity As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemPurchaseRefund As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCompanyPreferences As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemBankDetail As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemAccountsSummary As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVAT As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemFinancialYear As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemTaxCodes As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemReports As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemFile As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemHelp As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemExit As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Private WithEvents AcclExplorerBar1 As vbAccelerator.Components.Controls.acclExplorerBar
    Friend WithEvents MenuItemCustomerSummare As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCustomerAmend As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCustomerActivity As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSalesRefund As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCustomerListReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemCustomerActivityReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSalesInvoicesReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSalesReceiptsReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorListReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemVendorActivityReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemPurchaseInvoicesReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemPurchasePaymentsReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemLedgerBalanceReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemLedgerActivityReport As System.Windows.Forms.MenuItem
    Friend WithEvents _menuIcons As Chris.Beckett.MenuImageLib.MenuImage
    Friend WithEvents MenuItemContents As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemAboutInfini As System.Windows.Forms.MenuItem
    Friend WithEvents MenuTools As System.Windows.Forms.MenuItem
    Friend WithEvents MenuImport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuBackup As System.Windows.Forms.MenuItem
    Friend WithEvents MenuRestore As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemSynchronize As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MDI))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItemFile = New System.Windows.Forms.MenuItem
        Me.MenuItemExit = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItemCreateCustomer = New System.Windows.Forms.MenuItem
        Me.MenuItemCustomerSummare = New System.Windows.Forms.MenuItem
        Me.MenuItemCustomerAmend = New System.Windows.Forms.MenuItem
        Me.MenuItemSalesInvoices = New System.Windows.Forms.MenuItem
        Me.MenuItemSalesCreditNote = New System.Windows.Forms.MenuItem
        Me.MenuItemReceipts = New System.Windows.Forms.MenuItem
        Me.MenuItemCustomerActivity = New System.Windows.Forms.MenuItem
        Me.MenuItemSalesRefund = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItemCreateVendor = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorSummary = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorAmend = New System.Windows.Forms.MenuItem
        Me.MenuItemPurchaseInvoice = New System.Windows.Forms.MenuItem
        Me.MenuItemPurchaseCreditNote = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorPayment = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorActivity = New System.Windows.Forms.MenuItem
        Me.MenuItemPurchaseRefund = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItemCompanyPreferences = New System.Windows.Forms.MenuItem
        Me.MenuItemBankDetail = New System.Windows.Forms.MenuItem
        Me.MenuItemAccountsSummary = New System.Windows.Forms.MenuItem
        Me.MenuItemVAT = New System.Windows.Forms.MenuItem
        Me.MenuItemFinancialYear = New System.Windows.Forms.MenuItem
        Me.MenuItemTaxCodes = New System.Windows.Forms.MenuItem
        Me.MenuTools = New System.Windows.Forms.MenuItem
        Me.MenuImport = New System.Windows.Forms.MenuItem
        Me.MenuBackup = New System.Windows.Forms.MenuItem
        Me.MenuRestore = New System.Windows.Forms.MenuItem
        Me.MenuItemSynchronize = New System.Windows.Forms.MenuItem
        Me.MenuItemReports = New System.Windows.Forms.MenuItem
        Me.MenuItemCustomerListReport = New System.Windows.Forms.MenuItem
        Me.MenuItemCustomerActivityReport = New System.Windows.Forms.MenuItem
        Me.MenuItemSalesInvoicesReport = New System.Windows.Forms.MenuItem
        Me.MenuItemSalesReceiptsReport = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorListReport = New System.Windows.Forms.MenuItem
        Me.MenuItemVendorActivityReport = New System.Windows.Forms.MenuItem
        Me.MenuItemPurchaseInvoicesReport = New System.Windows.Forms.MenuItem
        Me.MenuItemPurchasePaymentsReport = New System.Windows.Forms.MenuItem
        Me.MenuItemLedgerBalanceReport = New System.Windows.Forms.MenuItem
        Me.MenuItemLedgerActivityReport = New System.Windows.Forms.MenuItem
        Me.MenuItemHelp = New System.Windows.Forms.MenuItem
        Me.MenuItemContents = New System.Windows.Forms.MenuItem
        Me.MenuItemAboutInfini = New System.Windows.Forms.MenuItem
        Me.AcclExplorerBar1 = New vbAccelerator.Components.Controls.acclExplorerBar
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me._menuIcons = New Chris.Beckett.MenuImageLib.MenuImage(Me.components)
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemFile, Me.MenuItem4, Me.MenuItem2, Me.MenuItem20, Me.MenuTools, Me.MenuItemReports, Me.MenuItemHelp})
        '
        'MenuItemFile
        '
        Me.MenuItemFile.Index = 0
        Me.MenuItemFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemExit})
        Me.MenuItemFile.Text = "&File"
        '
        'MenuItemExit
        '
        Me.MenuItemExit.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemExit, "0")
        Me.MenuItemExit.OwnerDraw = True
        Me.MenuItemExit.Text = "&Exit"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemCreateCustomer, Me.MenuItemCustomerSummare, Me.MenuItemCustomerAmend, Me.MenuItemSalesInvoices, Me.MenuItemSalesCreditNote, Me.MenuItemReceipts, Me.MenuItemCustomerActivity, Me.MenuItemSalesRefund})
        Me.MenuItem4.Text = "&Sales"
        '
        'MenuItemCreateCustomer
        '
        Me.MenuItemCreateCustomer.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemCreateCustomer, "1")
        Me.MenuItemCreateCustomer.OwnerDraw = True
        Me.MenuItemCreateCustomer.Text = "&Create Customer"
        '
        'MenuItemCustomerSummare
        '
        Me.MenuItemCustomerSummare.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuItemCustomerSummare, "2")
        Me.MenuItemCustomerSummare.OwnerDraw = True
        Me.MenuItemCustomerSummare.Text = "Customer &Summary"
        '
        'MenuItemCustomerAmend
        '
        Me.MenuItemCustomerAmend.Index = 2
        Me._menuIcons.SetMenuImage(Me.MenuItemCustomerAmend, "3")
        Me.MenuItemCustomerAmend.OwnerDraw = True
        Me.MenuItemCustomerAmend.Text = "Customer &Amend"
        '
        'MenuItemSalesInvoices
        '
        Me.MenuItemSalesInvoices.Index = 3
        Me._menuIcons.SetMenuImage(Me.MenuItemSalesInvoices, "4")
        Me.MenuItemSalesInvoices.OwnerDraw = True
        Me.MenuItemSalesInvoices.Text = "Sales &Invoices"
        '
        'MenuItemSalesCreditNote
        '
        Me.MenuItemSalesCreditNote.Index = 4
        Me._menuIcons.SetMenuImage(Me.MenuItemSalesCreditNote, "5")
        Me.MenuItemSalesCreditNote.OwnerDraw = True
        Me.MenuItemSalesCreditNote.Text = "Sales Credit &Notes"
        '
        'MenuItemReceipts
        '
        Me.MenuItemReceipts.Index = 5
        Me._menuIcons.SetMenuImage(Me.MenuItemReceipts, "6")
        Me.MenuItemReceipts.OwnerDraw = True
        Me.MenuItemReceipts.Text = "Customer &Receipts"
        '
        'MenuItemCustomerActivity
        '
        Me.MenuItemCustomerActivity.Index = 6
        Me._menuIcons.SetMenuImage(Me.MenuItemCustomerActivity, "7")
        Me.MenuItemCustomerActivity.OwnerDraw = True
        Me.MenuItemCustomerActivity.Text = "Customer &Activity"
        '
        'MenuItemSalesRefund
        '
        Me.MenuItemSalesRefund.Index = 7
        Me._menuIcons.SetMenuImage(Me.MenuItemSalesRefund, "8")
        Me.MenuItemSalesRefund.OwnerDraw = True
        Me.MenuItemSalesRefund.Text = "Sales R&efund"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemCreateVendor, Me.MenuItemVendorSummary, Me.MenuItemVendorAmend, Me.MenuItemPurchaseInvoice, Me.MenuItemPurchaseCreditNote, Me.MenuItemVendorPayment, Me.MenuItemVendorActivity, Me.MenuItemPurchaseRefund})
        Me.MenuItem2.Text = "&Purchase"
        '
        'MenuItemCreateVendor
        '
        Me.MenuItemCreateVendor.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemCreateVendor, "9")
        Me.MenuItemCreateVendor.OwnerDraw = True
        Me.MenuItemCreateVendor.Text = "&Create Vendor"
        '
        'MenuItemVendorSummary
        '
        Me.MenuItemVendorSummary.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorSummary, "10")
        Me.MenuItemVendorSummary.OwnerDraw = True
        Me.MenuItemVendorSummary.Text = "Vendor &Summary"
        '
        'MenuItemVendorAmend
        '
        Me.MenuItemVendorAmend.Index = 2
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorAmend, "11")
        Me.MenuItemVendorAmend.OwnerDraw = True
        Me.MenuItemVendorAmend.Text = "Vendor &Amend"
        '
        'MenuItemPurchaseInvoice
        '
        Me.MenuItemPurchaseInvoice.Index = 3
        Me._menuIcons.SetMenuImage(Me.MenuItemPurchaseInvoice, "12")
        Me.MenuItemPurchaseInvoice.OwnerDraw = True
        Me.MenuItemPurchaseInvoice.Text = "Purchase &Invoices"
        '
        'MenuItemPurchaseCreditNote
        '
        Me.MenuItemPurchaseCreditNote.Index = 4
        Me._menuIcons.SetMenuImage(Me.MenuItemPurchaseCreditNote, "13")
        Me.MenuItemPurchaseCreditNote.OwnerDraw = True
        Me.MenuItemPurchaseCreditNote.Text = "Purchase Credit &Notes"
        '
        'MenuItemVendorPayment
        '
        Me.MenuItemVendorPayment.Index = 5
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorPayment, "14")
        Me.MenuItemVendorPayment.OwnerDraw = True
        Me.MenuItemVendorPayment.Text = "Vendor &Payments"
        '
        'MenuItemVendorActivity
        '
        Me.MenuItemVendorActivity.Index = 6
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorActivity, "15")
        Me.MenuItemVendorActivity.OwnerDraw = True
        Me.MenuItemVendorActivity.Text = "Vendor A&ctivity"
        '
        'MenuItemPurchaseRefund
        '
        Me.MenuItemPurchaseRefund.Index = 7
        Me._menuIcons.SetMenuImage(Me.MenuItemPurchaseRefund, "16")
        Me.MenuItemPurchaseRefund.OwnerDraw = True
        Me.MenuItemPurchaseRefund.Text = "Purchases &Refund"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 3
        Me.MenuItem20.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemCompanyPreferences, Me.MenuItemBankDetail, Me.MenuItemAccountsSummary, Me.MenuItemVAT, Me.MenuItemFinancialYear, Me.MenuItemTaxCodes})
        Me.MenuItem20.Text = "&Accountant"
        '
        'MenuItemCompanyPreferences
        '
        Me.MenuItemCompanyPreferences.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemCompanyPreferences, "17")
        Me.MenuItemCompanyPreferences.OwnerDraw = True
        Me.MenuItemCompanyPreferences.Text = "&Company Preferrences"
        '
        'MenuItemBankDetail
        '
        Me.MenuItemBankDetail.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuItemBankDetail, "18")
        Me.MenuItemBankDetail.OwnerDraw = True
        Me.MenuItemBankDetail.Text = "&Bank Amend"
        '
        'MenuItemAccountsSummary
        '
        Me.MenuItemAccountsSummary.Index = 2
        Me._menuIcons.SetMenuImage(Me.MenuItemAccountsSummary, "19")
        Me.MenuItemAccountsSummary.OwnerDraw = True
        Me.MenuItemAccountsSummary.Text = "&Accounts Summary"
        '
        'MenuItemVAT
        '
        Me.MenuItemVAT.Index = 3
        Me._menuIcons.SetMenuImage(Me.MenuItemVAT, "20")
        Me.MenuItemVAT.OwnerDraw = True
        Me.MenuItemVAT.Text = "&Value Added Tax Return"
        '
        'MenuItemFinancialYear
        '
        Me.MenuItemFinancialYear.Index = 4
        Me._menuIcons.SetMenuImage(Me.MenuItemFinancialYear, "21")
        Me.MenuItemFinancialYear.OwnerDraw = True
        Me.MenuItemFinancialYear.Text = "&Financial Year"
        '
        'MenuItemTaxCodes
        '
        Me.MenuItemTaxCodes.Index = 5
        Me._menuIcons.SetMenuImage(Me.MenuItemTaxCodes, "22")
        Me.MenuItemTaxCodes.OwnerDraw = True
        Me.MenuItemTaxCodes.Text = "&Tax Codes"
        '
        'MenuTools
        '
        Me.MenuTools.Index = 4
        Me.MenuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuImport, Me.MenuBackup, Me.MenuRestore, Me.MenuItemSynchronize})
        Me.MenuTools.Text = "&Tools"
        '
        'MenuImport
        '
        Me.MenuImport.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuImport, "26")
        Me.MenuImport.OwnerDraw = True
        Me.MenuImport.Text = "&Export Data"
        '
        'MenuBackup
        '
        Me.MenuBackup.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuBackup, "27")
        Me.MenuBackup.OwnerDraw = True
        Me.MenuBackup.Text = "&Backup"
        '
        'MenuRestore
        '
        Me.MenuRestore.Index = 2
        Me._menuIcons.SetMenuImage(Me.MenuRestore, "28")
        Me.MenuRestore.OwnerDraw = True
        Me.MenuRestore.Text = "&Restore"
        '
        'MenuItemSynchronize
        '
        Me.MenuItemSynchronize.Index = 3
        Me._menuIcons.SetMenuImage(Me.MenuItemSynchronize, "29")
        Me.MenuItemSynchronize.OwnerDraw = True
        Me.MenuItemSynchronize.Text = "&Synchronize"
        '
        'MenuItemReports
        '
        Me.MenuItemReports.Index = 5
        Me.MenuItemReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemCustomerListReport, Me.MenuItemCustomerActivityReport, Me.MenuItemSalesInvoicesReport, Me.MenuItemSalesReceiptsReport, Me.MenuItemVendorListReport, Me.MenuItemVendorActivityReport, Me.MenuItemPurchaseInvoicesReport, Me.MenuItemPurchasePaymentsReport, Me.MenuItemLedgerBalanceReport, Me.MenuItemLedgerActivityReport})
        Me.MenuItemReports.Text = "&Reports"
        '
        'MenuItemCustomerListReport
        '
        Me.MenuItemCustomerListReport.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemCustomerListReport, "23")
        Me.MenuItemCustomerListReport.OwnerDraw = True
        Me.MenuItemCustomerListReport.Text = "&Customer List"
        '
        'MenuItemCustomerActivityReport
        '
        Me.MenuItemCustomerActivityReport.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuItemCustomerActivityReport, "23")
        Me.MenuItemCustomerActivityReport.OwnerDraw = True
        Me.MenuItemCustomerActivityReport.Text = "Customer &Activity"
        '
        'MenuItemSalesInvoicesReport
        '
        Me.MenuItemSalesInvoicesReport.Index = 2
        Me._menuIcons.SetMenuImage(Me.MenuItemSalesInvoicesReport, "23")
        Me.MenuItemSalesInvoicesReport.OwnerDraw = True
        Me.MenuItemSalesInvoicesReport.Text = "Sales &Invoices"
        '
        'MenuItemSalesReceiptsReport
        '
        Me.MenuItemSalesReceiptsReport.Index = 3
        Me._menuIcons.SetMenuImage(Me.MenuItemSalesReceiptsReport, "23")
        Me.MenuItemSalesReceiptsReport.OwnerDraw = True
        Me.MenuItemSalesReceiptsReport.Text = "Sales &Receipts"
        '
        'MenuItemVendorListReport
        '
        Me.MenuItemVendorListReport.Index = 4
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorListReport, "23")
        Me.MenuItemVendorListReport.OwnerDraw = True
        Me.MenuItemVendorListReport.Text = "Vendor &List"
        '
        'MenuItemVendorActivityReport
        '
        Me.MenuItemVendorActivityReport.Index = 5
        Me._menuIcons.SetMenuImage(Me.MenuItemVendorActivityReport, "23")
        Me.MenuItemVendorActivityReport.OwnerDraw = True
        Me.MenuItemVendorActivityReport.Text = "&Vendor Activity"
        '
        'MenuItemPurchaseInvoicesReport
        '
        Me.MenuItemPurchaseInvoicesReport.Index = 6
        Me._menuIcons.SetMenuImage(Me.MenuItemPurchaseInvoicesReport, "23")
        Me.MenuItemPurchaseInvoicesReport.OwnerDraw = True
        Me.MenuItemPurchaseInvoicesReport.Text = "Purchase I&nvoices"
        '
        'MenuItemPurchasePaymentsReport
        '
        Me.MenuItemPurchasePaymentsReport.Index = 7
        Me._menuIcons.SetMenuImage(Me.MenuItemPurchasePaymentsReport, "23")
        Me.MenuItemPurchasePaymentsReport.OwnerDraw = True
        Me.MenuItemPurchasePaymentsReport.Text = "Purchase &Payments"
        '
        'MenuItemLedgerBalanceReport
        '
        Me.MenuItemLedgerBalanceReport.Index = 8
        Me._menuIcons.SetMenuImage(Me.MenuItemLedgerBalanceReport, "23")
        Me.MenuItemLedgerBalanceReport.OwnerDraw = True
        Me.MenuItemLedgerBalanceReport.Text = "Ledger &Balance"
        '
        'MenuItemLedgerActivityReport
        '
        Me.MenuItemLedgerActivityReport.Index = 9
        Me._menuIcons.SetMenuImage(Me.MenuItemLedgerActivityReport, "23")
        Me.MenuItemLedgerActivityReport.OwnerDraw = True
        Me.MenuItemLedgerActivityReport.Text = "&Ledger Activity"
        '
        'MenuItemHelp
        '
        Me.MenuItemHelp.Index = 6
        Me.MenuItemHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemContents, Me.MenuItemAboutInfini})
        Me.MenuItemHelp.Text = "&Help"
        '
        'MenuItemContents
        '
        Me.MenuItemContents.Index = 0
        Me._menuIcons.SetMenuImage(Me.MenuItemContents, "25")
        Me.MenuItemContents.OwnerDraw = True
        Me.MenuItemContents.Text = "&Contents"
        '
        'MenuItemAboutInfini
        '
        Me.MenuItemAboutInfini.Index = 1
        Me._menuIcons.SetMenuImage(Me.MenuItemAboutInfini, "24")
        Me.MenuItemAboutInfini.OwnerDraw = True
        Me.MenuItemAboutInfini.Text = "&About Infini Accounts Express"
        '
        'AcclExplorerBar1
        '
        Me.AcclExplorerBar1.AnimateStateChanges = True
        Me.AcclExplorerBar1.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.AcclExplorerBar1.BackColorEnd = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.AcclExplorerBar1.BackColorStart = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.AcclExplorerBar1.Dock = System.Windows.Forms.DockStyle.Left
        Me.AcclExplorerBar1.DrawingStyle = vbAccelerator.Components.Controls.ExplorerBarDrawingStyle.Custom
        Me.AcclExplorerBar1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AcclExplorerBar1.ForeColor = System.Drawing.Color.Black
        Me.AcclExplorerBar1.ImageList = Nothing
        Me.AcclExplorerBar1.Mode = vbAccelerator.Components.Controls.ExplorerBarMode.Default
        Me.AcclExplorerBar1.Name = "AcclExplorerBar1"
        Me.AcclExplorerBar1.Redraw = True
        Me.AcclExplorerBar1.ShowFocusRect = True
        Me.AcclExplorerBar1.Size = New System.Drawing.Size(240, 573)
        Me.AcclExplorerBar1.TabIndex = 3
        Me.AcclExplorerBar1.TitleImageList = Nothing
        Me.AcclExplorerBar1.ToolTip = Nothing
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'ImageList2
        '
        Me.ImageList2.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit
        Me.ImageList2.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        '
        '_menuIcons
        '
        Me._menuIcons.ImageList = Me.ImageList1
        '
        'MDI
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 11)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(792, 573)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.AcclExplorerBar1})
        Me.Font = New System.Drawing.Font("Verdana", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MaximumSize = New System.Drawing.Size(1024, 768)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "MDI"
        Me.Text = " Infini Accounts Express"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Public Sub CheckForExistingInstance()
    '    If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
    '        'MessageBox.Show("Another Instance of this process is already running Multiple Instances Forbidden", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Application.Exit()
    '    End If
    'End Sub

    Private Sub MDI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Check Running Instance Of An Application
        'CheckForExistingInstance()

        Dim BFchild As New BackForm
        BFchild.MdiParent = Me
        BFchild.Sales.Visible = True
        BFchild.Purchase.Visible = False
        BFchild.Account.Visible = False
        BFchild.Reports.Visible = False
        BFchild.arrow1.Visible = True
        BFchild.arrow2.Visible = False
        BFchild.arrow3.Visible = False
        BFchild.arrow4.Visible = False
        BFchild.Panel2.Visible = False
        BFchild.Panel3.Visible = False
        BFchild.SalesReport.Visible = False
        BFchild.Show()

        Me.Visible = False
        Dim bar As New ExplorerBar
        Dim item As ExplorerBarLinkItem
        Dim baritems As New ExplorerBarItemCollection
        AcclExplorerBar1.ToolTip = ToolTip1
        AcclExplorerBar1.ImageList = ImageList1
        AcclExplorerBar1.TitleImageList = ImageList2
        AcclExplorerBar1.BackColor = System.Drawing.Color.FromArgb(102, 153, 204)

        bar.IconIndex = 0
        bar.Text = "SALES"
        bar.TitleForeColor = Color.White
        bar.TitleForeColorHot = Color.LightBlue
        bar.IsSpecial = True
        baritems = bar.Items
        bar.BackColor = System.Drawing.Color.FromArgb(65, 123, 181)
        bar.TitleBackColorStart = System.Drawing.Color.FromArgb(8, 36, 107)
        bar.TitleBackColorEnd = System.Drawing.Color.FromArgb(173, 203, 247)

        item = New ExplorerBarLinkItem
        item.IconIndex = 1
        item.Text = "Create Customer"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Create Customer"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 2
        item.Text = "Customer Summary"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer Summary"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 3
        item.Text = "Customer Amend"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer Amend"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 4
        item.Text = "Sales Invoices"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Sales Invoices"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 5
        item.Text = "Sales Credit Notes"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Sales Credit Notes"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 6
        item.Text = "Customer Receipts"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer Receipts"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 7
        item.Text = "Customer Activity"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer Activity"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 8
        item.Text = "Sales Refund"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Sales Refund"
        baritems.Add(item)
        AcclExplorerBar1.Bars.Add(bar)

        bar = New ExplorerBar
        item = New ExplorerBarLinkItem
        baritems = New ExplorerBarItemCollection
        bar.IconIndex = 1
        bar.Text = "PURCHASES"
        bar.TitleForeColor = Color.White
        bar.TitleForeColorHot = Color.LightBlue
        bar.IsSpecial = True
        baritems = bar.Items
        bar.BackColor = System.Drawing.Color.FromArgb(65, 123, 181)
        bar.TitleBackColorStart = System.Drawing.Color.FromArgb(8, 36, 107)
        bar.TitleBackColorEnd = System.Drawing.Color.FromArgb(173, 203, 247)

        item = New ExplorerBarLinkItem
        item.IconIndex = 9
        item.Text = "Create Vendor"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Create Vendor"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 10
        item.Text = "Vendor Summary"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor Summary"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 11
        item.Text = "Vendor Amend"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor Amend"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 12
        item.Text = "Purchase Invoices"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Purchase Invoices"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 13
        item.Text = "Purchase Credit Notes"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Purchase Credit Notes"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 14
        item.Text = "Vendor Payments"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor Payments"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 15
        item.Text = "Vendor Activity"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor Activity"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 16
        item.Text = "Purchases Refund"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Purchases Refund"
        baritems.Add(item)

        AcclExplorerBar1.Bars.Add(bar)
        bar.State = ExplorerBarState.Collapsed

        bar = New ExplorerBar
        item = New ExplorerBarLinkItem
        baritems = New ExplorerBarItemCollection
        bar.IconIndex = 2
        bar.Text = "ACCOUNTANT"
        bar.TitleForeColor = Color.White
        bar.TitleForeColorHot = Color.LightBlue
        bar.IsSpecial = True
        baritems = bar.Items
        bar.BackColor = System.Drawing.Color.FromArgb(65, 123, 181)
        bar.TitleBackColorStart = System.Drawing.Color.FromArgb(8, 36, 107)
        bar.TitleBackColorEnd = System.Drawing.Color.FromArgb(173, 203, 247)

        item = New ExplorerBarLinkItem
        item.IconIndex = 17
        item.Text = "Company Preferences"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Company Preferences"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 18
        item.Text = "Bank Amend"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Bank Amend"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 19
        item.Text = "Accounts Summary"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Accounts Summary"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 20
        item.Text = "Value Added Tax Return"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Value Added Tax Return"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 21
        item.Text = "Financial Year"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Financial Year"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 22
        item.Text = "Tax Codes"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Tax Codes"
        baritems.Add(item)

        AcclExplorerBar1.Bars.Add(bar)
        bar.State = ExplorerBarState.Collapsed

        '----------------------------
        bar = New ExplorerBar
        item = New ExplorerBarLinkItem
        baritems = New ExplorerBarItemCollection
        bar.IconIndex = 3
        bar.Text = "TOOLS"
        bar.TitleForeColor = Color.White
        bar.TitleForeColorHot = Color.LightBlue
        bar.IsSpecial = True
        baritems = bar.Items
        bar.BackColor = System.Drawing.Color.FromArgb(65, 123, 181)
        bar.TitleBackColorStart = System.Drawing.Color.FromArgb(8, 36, 107)
        bar.TitleBackColorEnd = System.Drawing.Color.FromArgb(173, 203, 247)

        item = New ExplorerBarLinkItem
        item.IconIndex = 26
        item.Text = "Export Data"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Export Data"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 27
        item.Text = "Backup"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Backup"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 28
        item.Text = "Restore"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Restore"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 29
        item.Text = "Synchronize"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Synchronize"
        baritems.Add(item)

        AcclExplorerBar1.Bars.Add(bar)
        bar.State = ExplorerBarState.Collapsed
        '----------------------------

        bar = New ExplorerBar
        item = New ExplorerBarLinkItem
        baritems = New ExplorerBarItemCollection
        bar.IconIndex = 4
        bar.Text = "REPORTS"
        bar.TitleForeColor = Color.White
        bar.TitleForeColorHot = Color.LightBlue
        bar.IsSpecial = True
        baritems = bar.Items
        bar.BackColor = System.Drawing.Color.FromArgb(65, 123, 181)
        bar.TitleBackColorStart = System.Drawing.Color.FromArgb(8, 36, 107)
        bar.TitleBackColorEnd = System.Drawing.Color.FromArgb(173, 203, 247)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Customer List"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer List"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Customer Activity"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Customer Activity"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Sales Invoices"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Sales Invoice"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Sales Receipts"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Sales Receipts"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Vendor List"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor List"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Vendor Activity"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Vendor Activity"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Purchase Invoices"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Purchase Invoices"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Purchase Payments"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Purchase Payments"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Ledger Balance"
        item.ForeColor = System.Drawing.Color.White
        item.ForeColorHot = Color.Yellow
        item.ToolTipText = "Ledger Balance"
        baritems.Add(item)

        item = New ExplorerBarLinkItem
        item.IconIndex = 23
        item.Text = "Ledger Activity"
        item.ForeColorHot = Color.Yellow
        item.ForeColor = System.Drawing.Color.White
        item.ToolTipText = "Ledger Activity"
        baritems.Add(item)

        AcclExplorerBar1.Bars.Add(bar)
        bar.State = ExplorerBarState.Collapsed

        Dim i As Int64
        For i = 0 To 25
            Thread.Sleep(100)
        Next
        Me.Visible = True
        Me.Activate()

        AcclExplorerBar1.Bars(0).MouseOver = True
        AcclExplorerBar1.Bars(1).MouseOver = True
        AcclExplorerBar1.Bars(2).MouseOver = True
        AcclExplorerBar1.Bars(3).MouseOver = True
        AcclExplorerBar1.Bars(4).MouseOver = True
        numOfForms = 0

        Dim mResponse As Integer
        mResponse = MessageBox.Show("Please Synchronize your Database before making any transactions.   " & vbNewLine & "Whould you like to synchronize.", "Infini Accounts Express", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If mResponse = 6 Then
            Dim child As New FrmBar
            child.MdiParent = Me
            child.Show()
            numOfForms = 1
        End If

    End Sub

    Private Sub bar_ItemClicked(ByVal sender As System.Object, ByVal e As vbAccelerator.Components.Controls.ExplorerBarItemClickEventArgs) Handles AcclExplorerBar1.ItemClick

        Dim NodeText As String
        Dim r As String = e.Item.Bar.Text
        NodeText = e.Item.ToolTipText
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If r = "SALES" Then
            Select Case NodeText

                Case "Create Customer"
                    Dim child As New CreateCustomer
                    child.Mdiparent = Me
                    child.Show()
                    numOfForms = 1

                Case "Customer Amend"
                    Dim child As New CustomerAmend
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Customer Summary"
                    Dim child As New CustomerSummary
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Sales Invoices"
                    Dim child As New SalesInvoice
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Sales Credit Notes"
                    Dim child As New SalesCreditInvoice
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Customer Receipts"
                    Dim child As New CustomerReceipts
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Sales Refund"
                    Dim child As New CustomerRefund
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Customer Activity"
                    Dim child As New CustomerActivity
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

            End Select

        ElseIf r = "PURCHASES" Then

            Select Case NodeText
                Case "Create Vendor"
                    Dim child As New CreateVendor
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Vendor Summary"
                    Dim child As New VendorSummary
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Vendor Amend"
                    Dim child As New VendorAmend
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Purchase Invoices"
                    Dim child As New PurchaseInvoice
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Purchase Credit Notes"
                    Dim child As New PurchaseCreditInvoice
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Vendor Activity"
                    Dim child As New VendorActivity
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Purchases Refund"
                    Dim child As New VendorRefund
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Vendor Payments"
                    Dim child As New VendorPayment
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

            End Select

        ElseIf r = "ACCOUNTANT" Then
            Select Case NodeText

                Case "Company Preferences"
                    Dim child As New CompanyPreferences
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Bank Amend"
                    Dim child As New BankAmend
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Accounts Summary"
                    Dim child As New LedgerSummary
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Value Added Tax Return"
                    Dim child As New VAT
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Financial Year"
                    Dim child As New FinancialYear
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Tax Codes"
                    Dim child As New TaxCode
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

            End Select

        ElseIf r = "TOOLS" Then
            Select Case NodeText

                Case "Export Data"
                    Dim child As New ImpExpUtility
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Backup"
                    Dim child As New Backup
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Restore"
                    Dim child As New Restore
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1

                Case "Synchronize"
                    Dim child As New FrmBar
                    child.MdiParent = Me
                    child.Show()
                    numOfForms = 1
            End Select

        ElseIf r = "REPORTS" Then
            Select Case NodeText
                Case "Customer List"
                    ShowReport.ReportName = "CustomerList"
                    Dim SR As New ShowReport
                    SR.Show()
                    SR.Activate()

                Case "Customer Activity"
                    ShowReport.ReportName = "CustomerActivity"
                    Dim CidforActivity As New GetCustomerID
                    CidforActivity.MdiParent = Me
                    CidforActivity.Show()
                    CidforActivity.Activate()

                Case "Sales Invoice"
                    ShowReport.ReportName = "SalesInvoice"
                    Dim CidforActivity As New GetCustomerID
                    CidforActivity.MdiParent = Me
                    CidforActivity.Show()
                    CidforActivity.Activate()

                Case "Sales Receipts"
                    ShowReport.ReportName = "SalesReceipts"
                    Dim CidforActivity As New GetCustomerID
                    CidforActivity.MdiParent = Me
                    CidforActivity.Show()
                    CidforActivity.Activate()

                Case "Vendor List"
                    ShowReport.ReportName = "VendorList"
                    Dim SR As New ShowReport
                    SR.Show()
                    sr.Activate()

                Case "Vendor Activity"
                    ShowReport.ReportName = "VendorActivity"
                    Dim VendorId As New GetVendorID
                    VendorId.MdiParent = Me
                    VendorId.Show()
                    VendorId.Activate()

                Case "Purchase Invoices"
                    ShowReport.ReportName = "PurchaseInvoices"
                    Dim VendorId As New GetVendorID
                    VendorId.MdiParent = Me
                    VendorId.Show()
                    VendorId.Activate()

                Case "Purchase Payments"
                    ShowReport.ReportName = "PurchasePayments"
                    Dim VendorId As New GetVendorID
                    VendorId.MdiParent = Me
                    VendorId.Show()
                    VendorId.Activate()

                Case "Ledger Balance"
                    ShowReport.ReportName = "LedgerBalance"
                    Dim SR As New ShowReport
                    SR.Show()
                    sr.Activate()

                Case "Ledger Activity"
                    ShowReport.ReportName = "LedgerActivity"
                    Dim LedgerId As New GetLedgerID
                    LedgerId.MdiParent = Me
                    LedgerId.Show()
                    LedgerId.Activate()

            End Select
        End If

    End Sub

#Region "MenuBarItems' Events"

    ' *********   SALES MENU  ************
    Private Sub MenuItemCreateCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCreateCustomer.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CreateCustomer
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemCustomerAmend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCustomerSummare.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CustomerAmend
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemCustomerSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCustomerAmend.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CustomerSummary
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemSalesInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSalesInvoices.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New SalesInvoice
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemSalesCreditNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSalesCreditNote.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New SalesCreditInvoice
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemReceipts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemReceipts.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CustomerReceipts
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub
    Private Sub MenuItemSalesRefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSalesRefund.Click

        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CustomerRefund
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    ' *********   PURCHASE MENU  ************
    Private Sub MenuItemCreateVendor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCreateVendor.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CreateVendor
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemVendorSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorSummary.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VendorSummary
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemVendorAmend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorAmend.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VendorAmend
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemPurchaseInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemPurchaseInvoice.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New PurchaseInvoice
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemPurchaseCreditNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemPurchaseCreditNote.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New PurchaseCreditInvoice
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemVendorPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorPayment.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VendorPayment
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemVendorActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorActivity.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VendorActivity
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemPurchaseRefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemPurchaseRefund.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VendorRefund
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemCompanyPreferences_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCompanyPreferences.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New CompanyPreferences
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemBankDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemBankDetail.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New BankAmend
        child.MdiParent = Me
        child.Show()
        numOfForms = 1

    End Sub

    Private Sub MenuItemAccountsSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemAccountsSummary.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New LedgerSummary
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVAT.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New VAT
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemFinancialYear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemFinancialYear.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New FinancialYear
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemTaxCodes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemTaxCodes.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New TaxCode
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemExit.Click
        Application.Exit()
    End Sub

    Private Sub MenuItemCustomerActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCustomerActivity.Click
        Dim child As New CustomerActivity
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub MenuItemCustomerListReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCustomerListReport.Click
        ShowReport.ReportName = "CustomerList"
        Dim SR As New ShowReport
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub MenuItemCustomerActivityReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemCustomerActivityReport.Click
        ShowReport.ReportName = "CustomerActivity"
        Dim CidforActivity As New GetCustomerID
        CidforActivity.MdiParent = Me
        CidforActivity.Show()
        CidforActivity.Activate()
    End Sub

    Private Sub MenuItemSalesInvoicesReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSalesInvoicesReport.Click
        ShowReport.ReportName = "SalesInvoice"
        Dim CidforActivity As New GetCustomerID
        CidforActivity.MdiParent = Me
        CidforActivity.Show()
        CidforActivity.Activate()
    End Sub

    Private Sub MenuItemSalesReceiptsReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSalesReceiptsReport.Click
        ShowReport.ReportName = "SalesReceipts"
        Dim CidforActivity As New GetCustomerID
        CidforActivity.MdiParent = Me
        CidforActivity.Show()
        CidforActivity.Activate()
    End Sub

    Private Sub MenuItemVendorListReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorListReport.Click
        ShowReport.ReportName = "VendorList"
        Dim SR As New ShowReport
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub MenuItemVendorActivityReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVendorActivityReport.Click

        ShowReport.ReportName = "VendorActivity"
        Dim VendorId As New GetVendorID
        VendorId.MdiParent = Me
        VendorId.Show()
        VendorId.Activate()

    End Sub

    Private Sub MenuItemPurchaseInvoicesReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemPurchaseInvoicesReport.Click

        ShowReport.ReportName = "PurchaseInvoices"
        Dim VendorId As New GetVendorID
        VendorId.MdiParent = Me
        VendorId.Show()
        VendorId.Activate()
    End Sub

    Private Sub MenuItemPurchasePaymentsReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemPurchasePaymentsReport.Click
        ShowReport.ReportName = "PurchasePayments"
        Dim VendorId As New GetVendorID
        VendorId.MdiParent = Me
        VendorId.Show()
        VendorId.Activate()
    End Sub

    Private Sub MenuItemLedgerBalanceReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemLedgerBalanceReport.Click
        ShowReport.ReportName = "LedgerBalance"
        Dim SR As New ShowReport
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub MenuItemLedgerActivityReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemLedgerActivityReport.Click
        ShowReport.ReportName = "LedgerActivity"
        Dim LedgerId As New GetLedgerID
        LedgerId.MdiParent = Me
        LedgerId.Show()
        LedgerId.Activate()
    End Sub

    Private Sub MenuItemAboutInfini_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemAboutInfini.Click
        Dim child As New FrmAbout
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

#End Region

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuImport.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New ImpExpUtility
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuBackup.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New Backup
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuRestore.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New Restore
        child.MdiParent = Me
        child.Show()
        numOfForms = 1
    End Sub

    Private Sub MenuItemSynchronize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemSynchronize.Click
        If numOfForms = 1 Then
            MessageBox.Show("You must close the open window first .", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New FrmBar
        child.MdiParent = Me
        child.Show()
        numOfForms = 1

    End Sub

End Class
