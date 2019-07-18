Imports System.IO
Imports InfiniExpress.BLL
Imports System.Data.OleDb
Public Class ShowReport
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
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ShowReport))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(880, 581)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'ShowReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(880, 581)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(135, 35)
        Me.Name = "ShowReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Local Variables"

    Public Shared ReportName As String
    Public Shared CustId As String
    Public Shared CustName As String
    Public Shared VendName As String
    Public Shared VendId As String
    Public Shared LedgerId As String
    Public Shared LedgerName As String
    Dim GlbReports As New ClassReports()
#End Region

    Private Sub ShowReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = ReportName & " Report"

        CrystalReportViewer1.DisplayGroupTree = False

        Select Case ReportName

            Case "Customer Creditcard Activity"
                Dim objCreditcard As New Customer_Creditcard_info
                Dim Ds As New DataSet

                Ds = GlbReports.ShowCreditCardInfo(ShowReport.CustId)
                If Ds.Tables(0).Rows.Count = 0 Then
                    Ds = New DataSet
                    Ds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(Ds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Customer Creditcard Activity"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    objCreditcard.SetDataSource(Ds)
                    CrystalReportViewer1.ReportSource = objCreditcard
                End If


            Case "CustomerList"
                Dim objCustList As New CustomerListReport
                Dim myds As New DataSet
                myds = GlbReports.GetCustList()
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Customer List"
                    CrystalReportViewer1.ReportSource = objDefault
                Else

                    objCustList.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objCustList
                End If


            Case "CustomerActivity"
                Dim objCustActivity As New CustomerActivityReport
                Dim myds As New DataSet
                myds = GlbReports.GetCustActivity(CustId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Customer Activity"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    objCustActivity.SetDataSource(myds)
                    CType(objCustActivity.ReportDefinition.ReportObjects.Item("Text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustId
                    CType(objCustActivity.ReportDefinition.ReportObjects.Item("Text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustName
                    CType(objCustActivity.ReportDefinition.ReportObjects.Item("Text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Customer Activity"
                    CrystalReportViewer1.ReportSource = objCustActivity
                End If


            Case "SalesInvoice"
                Dim objSalesInvoiceR As New SalesInvoiceReport
                Dim myds As New DataSet
                myds = GlbReports.GetSalesInvoice(CustId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Sales Invoice"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    CType(objSalesInvoiceR.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustId
                    CType(objSalesInvoiceR.ReportDefinition.ReportObjects.Item("Text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustName
                    objSalesInvoiceR.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objSalesInvoiceR

                End If


            Case "SalesReceipts"
                Dim objSalesReceiptsR As New SalesReceiptsReport
                Dim myds As New DataSet
                myds = GlbReports.GetSalesReceipts(CustId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Sales Receipts"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    CType(objSalesReceiptsR.ReportDefinition.ReportObjects.Item("Text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Sales Receipts"
                    CType(objSalesReceiptsR.ReportDefinition.ReportObjects.Item("Text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustId
                    CType(objSalesReceiptsR.ReportDefinition.ReportObjects.Item("Text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = CustName
                    objSalesReceiptsR.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objSalesReceiptsR
                End If

            Case "VendorList"
                Dim objVendList As New VendorListReport
                Dim myds As New DataSet
                myds = GlbReports.GetVendorList()
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor List"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    objVendList.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objVendList
                End If

            Case "VendorActivity"
                Dim objVendAct As New CustomerActivityReport
                Dim myds As New DataSet
                myds = GlbReports.GetVendorActivity(VendId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Activity"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    CType(objVendAct.ReportDefinition.ReportObjects.Item("Text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor ID: "
                    CType(objVendAct.ReportDefinition.ReportObjects.Item("Text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendId
                    CType(objVendAct.ReportDefinition.ReportObjects.Item("Text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Name: "
                    CType(objVendAct.ReportDefinition.ReportObjects.Item("Text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendName
                    CType(objVendAct.ReportDefinition.ReportObjects.Item("Text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Activity"
                    objVendAct.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objVendAct
                End If

            Case "PurchaseInvoices"
                Dim objPurchaseInv As New PurchaseInvoiceReport
                Dim myds As New DataSet
                myds = GlbReports.GetPurchaseInvoice(VendId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Purchase Invoices"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    CType(objPurchaseInv.ReportDefinition.ReportObjects.Item("Text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendId
                    CType(objPurchaseInv.ReportDefinition.ReportObjects.Item("Text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendName
                    objPurchaseInv.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objPurchaseInv
                End If

            Case "PurchasePayments"
                Dim objPurchasePay As New SalesReceiptsReport
                Dim myds As New DataSet
                myds = GlbReports.GetPurchasePayments(VendId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Purchase Payments"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    'CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendId
                    'CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendName
                    CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Purchase Payments"
                    CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor ID:"
                    CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Name:"
                    CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendId
                    CType(objPurchasePay.ReportDefinition.ReportObjects.Item("Text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = VendName
                    objPurchasePay.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objPurchasePay
                End If

            Case "LedgerBalance"
                Dim objLedgerBal As New LedgerBalanceReport
                Dim myds As New DataSet
                myds = GlbReports.GetLedgerBalance()
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Ledger Balance"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    objLedgerBal.SetDataSource(myds)
                    CrystalReportViewer1.ReportSource = objLedgerBal
                End If

            Case "LedgerActivity"
                Dim objLedgerAct As New LedgerActivityReport
                Dim myds As New DataSet
                myds = GlbReports.GetLedgerActivity(LedgerId)
                If myds.Tables(0).Rows.Count = 0 Then
                    myds = New DataSet
                    myds = GlbReports.GetCompanyInfo()
                    Dim objDefault As New DefaultReportforEmptyDS
                    objDefault.SetDataSource(myds)
                    CType(objDefault.ReportDefinition.ReportObjects.Item("Text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Ledger Activity"
                    CrystalReportViewer1.ReportSource = objDefault
                Else
                    objLedgerAct.SetDataSource(myds)
                    CType(objLedgerAct.ReportDefinition.ReportObjects.Item("Text14"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = LedgerName
                    CrystalReportViewer1.ReportSource = objLedgerAct
                    LedgerName = ""
                End If

        End Select


    End Sub
    
    'Private Sub tempCreateXMLSchema()
    '    Dim myds As New DataSet()
    '    Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
    '    Dim oleAdapt As New OleDbDataAdapter("SELECT CustomerId,  CustomerName  ,  (Street +' '  +  Town  + ' '  + Country + ' '  + PostalCode )   As Address  ,  ContactName,Email,Telephone , Fax FROM TblCustomer ", olecon)
    '    oleAdapt.Fill(myds)

    '    Dim Filename As String = Application.StartupPath & "\CustList.xml"
    '    Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
    '    Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
    '    myds.WriteXmlSchema(MyXmlTextWriter)
    '    myFileStream.Close()

    'End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class
