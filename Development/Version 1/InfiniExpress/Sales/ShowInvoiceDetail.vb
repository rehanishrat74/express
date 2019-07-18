Imports InfiniExpress.BLL
Public Class ShowInvoiceDetail
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
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ShowInvoiceDetail))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
        'ToolTip1
        '
        Me.ToolTip1.Active = False
        '
        'ShowInvoiceDetail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(880, 581)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.Name = "ShowInvoiceDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "ShowInvoiceDetailReport"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim clsSales As New ClassSales()
    Dim clsPurchases As New ClassPurchases
    Dim ClsReport As New ClassReports

    Public Sub MakeSalesInvoiceDetail(ByVal CustID As String, ByVal ParentIDs As String, ByVal ReportTitle As String)
        Dim ds As New DataSet
        Dim objSIDetailReport As New SalesInvoiceDetailReport

        If ReportTitle = "Purchase Invoice" Or ReportTitle = "Purchase Credit Invoice" Then
            ds = clsPurchases.GetPurchaseInvoiveDetail(CustID, ParentIDs)
            CType(objSIDetailReport.ReportDefinition.ReportObjects.Item("Text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Details"
            CType(objSIDetailReport.ReportDefinition.ReportObjects.Item("Text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor ID"
            CType(objSIDetailReport.ReportDefinition.ReportObjects.Item("Text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Vendor Name"


        Else
            'ParentIDs = ClsReport.GetParentIDno(ParentIDs)
            ds = clsSales.GetSalesInvoiveDetail(CustID, ParentIDs)
            'ds.WriteXml("C:\SalInvReport.XML")
        End If

        objSIDetailReport.SetDataSource(ds)
        CType(objSIDetailReport.ReportDefinition.ReportObjects.Item("Text23"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = Convert.ToString(ds.Tables(0).Rows.Count)
        CType(objSIDetailReport.ReportDefinition.ReportObjects.Item("Text27"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = ReportTitle
        CrystalReportViewer1.ReportSource = objSIDetailReport
        CrystalReportViewer1.DisplayGroupTree = False
    End Sub

    

End Class

