Imports Microsoft.VisualBasic
Imports InfiniExpress.BLL

Public Class VendorRefund

    Inherits System.Windows.Forms.Form
    Dim ClsPurchase As New ClassPurchases()
    Dim ClsGlobal As New ClassGlobal()
    Dim CRow As Integer, AutoNo As Integer
    Public Dr As IDataReader
    Public Ds As New DataSet()
    Dim VendID As String

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
    Protected WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents Button4 As System.Windows.Forms.Button
    Protected WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Protected WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents cmbvendorid As System.Windows.Forms.ComboBox
    Protected WithEvents btnOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VendorRefund))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbvendorid = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button4, Me.btnOk, Me.DataGrid1, Me.TextBox1, Me.Label2, Me.cmbvendorid, Me.Label1})
        Me.Panel1.Location = New System.Drawing.Point(5, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(546, 252)
        Me.Panel1.TabIndex = 1
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(88, 222)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(60, 18)
        Me.Button4.TabIndex = 34
        Me.Button4.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(16, 222)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 18)
        Me.btnOk.TabIndex = 33
        Me.btnOk.Text = "OK"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(16, 44)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(520, 168)
        Me.DataGrid1.TabIndex = 4
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(316, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(216, 20)
        Me.TextBox1.TabIndex = 3
        Me.TextBox1.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(226, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Vendor Name"
        '
        'cmbvendorid
        '
        Me.cmbvendorid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbvendorid.Location = New System.Drawing.Point(90, 16)
        Me.cmbvendorid.Name = "cmbvendorid"
        Me.cmbvendorid.Size = New System.Drawing.Size(134, 21)
        Me.cmbvendorid.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Vendor ID"
        '
        'VendorRefund
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(556, 263)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VendorRefund"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Purchase Refund"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub VendorRefund_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dr = ClsPurchase.LoadVendorID
        While Dr.Read
            CmbVendorID.Items.Add(Dr.Item("Vendorid"))
        End While
        Dr.Close()
        populatingGrid("")
        Ds = New DataSet()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose(True)
    End Sub

    Private Sub cmbvendorid_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbvendorid.SelectedIndexChanged
        VendID = (cmbvendorid.SelectedItem)
        Dr = ClsPurchase.GetVendorDetails(VendID)
        Dr.Read()
        If IsDBNull(Dr("Vendorname")) = False Then
            TextBox1.Text = Dr("Vendorname")
        End If
        Dr.Close()
        populatingGrid(VendID)
    End Sub

    Sub populatingGrid(ByVal tempCid As String)

        Ds = New DataSet()
        Dim Dtab As DataTable

        Dtab = PurchaseRefund(VendID)
        Ds.Merge((Dtab))
        DataGrid1.TableStyles.Clear()

        'Formating The DataGrid Bt Adding a DataGrid Table Style
        Dim DgridTstyle As New DataGridTableStyle()
        Dim TranNo As New DataGridBrowser.DataGridNoActiveCellColumn() 'parentid
        Dim Type As New DataGridBrowser.DataGridNoActiveCellColumn()    'InvTp
        Dim InvDate As New DataGridBrowser.DataGridNoActiveCellColumn() 'InvDate  
        Dim Details As New DataGridBrowser.DataGridNoActiveCellColumn() 'InvDetails
        Dim Amount As New DataGridBrowser.DataGridNoActiveCellColumn()  'ODebit
        Dim Net As New DataGridBrowser.DataGridNoActiveCellColumn() 'InvNet
        Dim VAT As New DataGridBrowser.DataGridNoActiveCellColumn()   'InvVat
        Dim TaxID As New DataGridBrowser.DataGridNoActiveCellColumn() 'TaxID
        Dim Paystatus As New DataGridBrowser.DataGridNoActiveCellColumn() 'paystatus
        Dim VendorID As New DataGridBrowser.DataGridNoActiveCellColumn() 'TransID

        'Setting Default Properties for datagrid like backcolor,forecolor,fonts
        With DgridTstyle
            .HeaderFont = New System.Drawing.Font _
            ("Verdana", 8.25F, System.Drawing.FontStyle.Regular, _
             System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
            .SelectionBackColor = Color.Navy
            .AlternatingBackColor = Color.Lavender
            .PreferredRowHeight = 3
        End With

        With TranNo
            .HeaderText = "Tran.No"
            .MappingName = "parentid"
            .Width = 55
            .ReadOnly = True
        End With

        With Type
            .HeaderText = "Type"
            .MappingName = "invtp"
            .Width = 40
            .ReadOnly = True
        End With

        With InvDate
            .HeaderText = "Date"
            .MappingName = "invdate"
            .Width = 70
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Details
            .HeaderText = "Details"
            .MappingName = "invdetails"
            .Width = 235
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Amount
            .HeaderText = "Amount" & " " & Chr(9)
            .Alignment = HorizontalAlignment.Right
            .MappingName = "OCredit"
            .Width = 100
            .ReadOnly = True
            .Format = "#0.00"
            .TextBox.ReadOnly = True
            .TextBox.TextAlign = HorizontalAlignment.Right
        End With

        With Net
            .HeaderText = ""
            .Width = 0
            .MappingName = "invnet"
            .TextBox.MaxLength = 10
            .Format = "#0.00"
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
        End With

        With VAT
            .HeaderText = ""
            .Width = 0
            .MappingName = "invvat"
            .TextBox.MaxLength = 10
            .Format = "#0.00"
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
        End With

        With TaxID
            .HeaderText = ""
            .Width = 0
            .MappingName = "taxid"
        End With

        With Paystatus
            .HeaderText = ""
            .Width = 0
            .MappingName = "Paystatus"
        End With

        With VendorID
            .HeaderText = ""
            .Width = 0
            .MappingName = "transid"
        End With

        'Add The GridColumnStyle to Datagrid Column Style
        With DgridTstyle.GridColumnStyles
            .Add(TranNo)
            .Add(Type)
            .Add(InvDate)
            .Add(Details)
            .Add(Amount)
            .Add(Net)
            .Add(VAT)
            .Add(TaxID)
            .Add(Paystatus)
            .Add(VendorID)
        End With

        'Bind The Datagrid to Dataset
        If Dtab.DefaultView.Count > 0 Then
            DataGrid1.DataSource = Ds.Tables(0)
            'Remove Empty Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False
            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = Ds.Tables(0).TableName '("TblTransaction")
            DataGrid1.TableStyles.Add(DgridTstyle)
            btnOk.Enabled = True
        Else
            'Populating Empty DataGrid 
            Ds.Clear()
            DataGrid1.DataSource = Ds.Tables(0)
            'Stoping Extra Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False
            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = Ds.Tables(0).TableName '("TblTransaction")
            DataGrid1.TableStyles.Add(DgridTstyle)
            btnOk.Enabled = False
        End If

    End Sub

    Public Function PurchaseRefund(ByVal Id As String) As DataTable

        Dim Tinvtp As String, Tinvdate As String, Tinvdetails As String
        Dim Tinvnet As Double, Tinvvat As Double, ToCredit As Double
        Dim Ttaxid As String, Tpaystatus As String, Ttransid As String
        Dim R3 As String, Paystatus As String, Tid As Integer
        Dim ScAmt As Double, SrAmt As Double, Tparentid As Integer
        Dim Rd As IDataReader

        'Creating DataTable
        Dim DT As New DataTable()
        Dim DCTNO As New DataColumn("Parentid", GetType(Integer))
        Dim DCType As New DataColumn("Invtp", GetType(String))
        Dim DCDate As New DataColumn("InvDate", GetType(String))
        Dim DCTDetail As New DataColumn("InvDetails", GetType(String))
        Dim DCNet As New DataColumn("InvNet", GetType(Double))
        Dim DCVat As New DataColumn("InvVat", GetType(Double))
        Dim DCAmt As New DataColumn("OCredit", GetType(Double))
        Dim DCTaxid As New DataColumn("Taxid", GetType(String))
        Dim DCPaystatus As New DataColumn("Paystatus", GetType(String))
        Dim DCVendorid As New DataColumn("transid", GetType(String))

        'Add The Column to the Datatable's Columns Collection
        DT.Columns.Add(DCTNO)
        DT.Columns.Add(DCType)
        DT.Columns.Add(DCDate)
        DT.Columns.Add(DCTDetail)
        DT.Columns.Add(DCNet)
        DT.Columns.Add(DCVat)
        DT.Columns.Add(DCAmt)
        DT.Columns.Add(DCTaxid)
        DT.Columns.Add(DCPaystatus)
        DT.Columns.Add(DCVendorid)

        'Creating DataRow
        Dim DR As DataRow
        Rd = ClsPurchase.PurchaseRefund(VendID)

        If Rd.Read Then
            Tid = Rd("parentid")
            Rd.Close()

            Rd = ClsPurchase.PurchaseRefund(VendID)
            While Rd.Read
                If Tid = Rd("Parentid") Then

                    Tparentid = Rd("parentid")
                    Tinvtp = Rd("Invtp")
                    Tinvdate = Rd("Invdate")
                    Tinvdetails = Rd("Invdetails")
                    Tinvnet = Rd("Invnet")
                    Tinvvat = Rd("Invvat")
                    ToCredit = SrAmt
                    Ttaxid = Rd("Taxid")
                    Tpaystatus = Rd("paystatus")
                    Ttransid = Rd("transid")

                    If Rd("fromtp") = "PC" Then
                        ScAmt = ScAmt + Rd("OCredit")
                    Else
                        SrAmt = SrAmt + Rd("OCredit")
                    End If

                    GoTo NextRecord 'MOVE TO NEXT RECORD

                Else

                    If SrAmt > 0 Then
                        If ScAmt + SrAmt = Rd("LCredit") Then
                            Paystatus = "f"
                        Else
                            Paystatus = "p"
                        End If

                        DR = DT.NewRow()
                    DR("parentid") = Tparentid
                    DR("Invtp") = Tinvtp
                    DR("Invdate") = Tinvdate
                    DR("Invdetails") = Tinvdetails
                    DR("Invnet") = Tinvnet
                    DR("Invvat") = Tinvvat
                    DR("OCredit") = Format(SrAmt, "0.00")
                    DR("Taxid") = Ttaxid
                    DR("Paystatus") = Tpaystatus
                    DR("transid") = Ttransid
                    DT.Rows.Add(DR) 'Adding DataRow To DataTable

                    End If

                    Tid = Rd("parentid") 'Assign New Parentid To Variable
                    Dim rr = Rd("fromtp")
                    Tparentid = Rd("parentid")
                    Tinvtp = Rd("Invtp")
                    Tinvdate = Rd("Invdate")
                    Tinvdetails = Rd("Invdetails")
                    Tinvnet = Rd("Invnet")
                    Tinvvat = Rd("Invvat")
                    ToCredit = SrAmt
                    Ttaxid = Rd("Taxid")
                    Tpaystatus = Rd("paystatus")
                    Ttransid = Rd("transid")

                    If Tid = Rd("parentid") Then
                        ScAmt = 0
                        SrAmt = 0
                        If Rd("fromtp") = "PC" Then
                            ScAmt = ScAmt + Rd("OCredit")
                        Else
                            SrAmt = SrAmt + Rd("OCredit")
                        End If
                    End If
                End If
NextRecord:
            End While
        Else
            Return DT
            Exit Function
        End If
        Rd.Close()

        If SrAmt > 0 Then
            DR = DT.NewRow()
            DR("parentid") = Tparentid
            DR("Invtp") = Tinvtp
            DR("Invdate") = Tinvdate
            DR("Invdetails") = Tinvdetails
            DR("Invnet") = Tinvnet
            DR("Invvat") = Tinvvat
            DR("OCredit") = Format(SrAmt, "0.00")
            DR("Taxid") = Ttaxid
            DR("Paystatus") = Tpaystatus
            DR("transid") = Ttransid
            DT.Rows.Add(DR) 'Adding DataRow To DataTable
        End If

        Return DT

    End Function

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim recordno As Double, vat As Double
        Dim ActivityNo As Double, Details As String, SFlag As String
        Dim transno As Double, totrans As String, fromtrans As String
        Dim TParentid As Integer = DataGrid1.Item(CRow, 0) '"Parentid"
        Dim TnetAmt As Double = DataGrid1.Item(CRow, 5) '"NetAmt"
        Dim TVatAmt As Double = DataGrid1.Item(CRow, 6) '"VatAmt"
        Dim TAmt As Double = DataGrid1.Item(CRow, 4) '"Amt"
        Dim TXid As String = DataGrid1.Item(CRow, 7) '"TAXID"
        Dim TPstatus As String = DataGrid1.Item(CRow, 8) '"paystatus"
        Dim VendID As String = DataGrid1.Item(CRow, 9) '"CustomerID"
        Dim ChildID As Double, ParentID As Double, Assparentid As Double
        Dim rightNow As DateTime = DateTime.Now
        Dim SysDate As String, PDAmt As Double, AssumeNo As Double
        Dim MyAuto1 As Integer, OSMaxNo As Integer
        Dim ReserveParentID As Double

        If TParentid = 0 Then Exit Sub
        SysDate = rightNow.ToString("dd/MM/yyyy")
        Dim MyDate As Date = Now
        MyDate = DateSerial(Year(MyDate), Month(MyDate), Day(MyDate))

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        Dr = ClsPurchase.VendorInvoiceRefund(TParentid)

        While Dr.Read
            If Dr("invtp") <> "PC" Then
                ClsGlobal.UpdateOutstanding(Dr("myauto"), DTime)
            Else
                Exit Sub
            End If
        End While
        Dr.Close()

        'Calculationg Dicount
        PDAmt = ClsPurchase.GetDCAgainstTransaction(TParentid, "PD")

        'Get Transaction No
        ParentID = ClsGlobal.GetTransactionNo
        'Get Ledger Autono
        AutoNo = ClsGlobal.GetLedgerAutoNo
        'Get New Outstanding Max. No
        OSMaxNo = ClsGlobal.GetOutStandingMaxOSNo()

        Assparentid = ParentID
        ChildID = ParentID

        If PDAmt > 0 Then
            AssumeNo = Assparentid + 2
        End If

        vat = TVatAmt
        Details = "Refund - " & TParentid

        'Set Refund = y for those invoice who is select by user to REFUND.
        ClsGlobal.ChkRefund("Y", "Y", TParentid, DTime)

        'Step 1
        ClsGlobal.InsertTransactionData(ParentID, "PC", TXid, SysDate, Details, TnetAmt, TVatAmt, MyDate, DTime, "N", "N", "N")

        totrans = Format(TAmt - PDAmt, "0.00") & " To PI " & ParentID + 1
        fromtrans = Format(TAmt - PDAmt, "0.00") & " From PC " & ParentID

        'Step 2
        'credit customer account
        ' Application("MATERIALSPURCHASEDACCOUNT") = "20000"
        ClsGlobal.InsertLedger(ParentID, ChildID, VendID, TnetAmt + TVatAmt, 0, "S", "20000", "", AutoNo, DTime)

        'Step3
        'credit outstanding for customer & refund
        ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, VendID, TAmt - PDAmt, 0, "S", totrans, "Refund", SysDate, 0, ParentID + 1, "T", DTime, OSMaxNo)

        ' Special Case For Sale Discount
        If PDAmt > 0 Then
            totrans = Format(PDAmt, "0.00") & " To PI " & AssumeNo
            OSMaxNo = OSMaxNo + 1
            ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, VendID, PDAmt, 0, "S", totrans, "Refund", SysDate, 0, CStr(AssumeNo), "T", DTime, OSMaxNo)
        End If

        'Set Paystatus Flag
        ReserveParentID = ParentID

        'Step4
        'credit for nc debtors control account
        'Application("CREDITORSCONTROLACCOUNT") = "71000"
        'Application("MATERIALSPURCHASEDACCOUNT") = "20000"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "71000", TnetAmt + TVatAmt, 0, "T", "20000", "", AutoNo, DTime)

        'Step5
        ''debit sales tax control account
        'Application("MATERIALSPURCHASEDACCOUNT") = "20000"
        'Application("VATINPUTACCOUNT") = "71101"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "71101", 0, TVatAmt, "T", "20000", "", AutoNo, DTime)

        'Step6
        'Application("MATERIALSPURCHASEDACCOUNT") = "20000"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "20000", 0, TnetAmt, "T", "REFUND", "", AutoNo, DTime)

        'allocation-1
        ParentID = ParentID + 1 'ClsGlobal
        ChildID = ParentID

        Details = "Allocate - " & TParentid

        'Step1
        ClsGlobal.InsertTransactionData(ParentID, "PI", "T9", SysDate, Details, TAmt - PDAmt, 0, MyDate, DTime, "-", "Y", "Y")

        'Step2
        'Debit Customer Account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, VendID, 0, TAmt - PDAmt, "S", "49999", "f", AutoNo, DTime)

        'Step3
        'Debit Outstanding For Customer & Refund
        OSMaxNo = OSMaxNo + 1
        ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, VendID, 0, TAmt - PDAmt, "S", fromtrans, "REFUND", SysDate, ParentID - 1, 0, "F", DTime, OSMaxNo)

        'Step4
        'Debit For NC Debtors Control Account
        'Application("CREDITORSCONTROLACCOUNT") = "71000"
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "71000", 0, TAmt - PDAmt, "T", "49999", "", AutoNo, DTime)

        'Step5
        'Credit Misposting account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "49999", TAmt - PDAmt, 0, "T", "REFUND", "", AutoNo, DTime)

        'Create SD Transaction
        If PDAmt > 0 Then

            ParentID = ParentID + 1
            ChildID = ParentID
            Details = "Allocate - " & TParentid

            'Step 1
            ClsGlobal.InsertTransactionData(ParentID, "PI", "T9", SysDate, Details, PDAmt, 0, MyDate, DTime, "-", "Y", "Y")

            'Step2
            'Debit Customer Account
            'Application("PURCHASEDISCOUNT") = "20010"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, VendID, 0, PDAmt, "S", "20010", "f", AutoNo, DTime)

            fromtrans = Format(PDAmt, "0.00") & " From PC " & ParentID - 2

            'Step3
            'Debit Outstanding For Customer & Refund
            OSMaxNo = OSMaxNo + 1
            ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, VendID, 0, PDAmt, "S", fromtrans, "REFUND", SysDate, ParentID - 2, 0, "F", DTime, OSMaxNo)

            'Step4
            'Debit For NC Debtors Control Account
            'Application("CREDITORSCONTROLACCOUNT") = "71000"
            'Application("PURCHASEDISCOUNT") = "20010"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, "71000", 0, PDAmt, "T", "20010", "", AutoNo, DTime)

            'Step5
            'Credit Discount
            'Application("PURCHASEDISCOUNT") = "20010"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, "20010", PDAmt, 0, "T", "REFUND", "", AutoNo, DTime)

        End If

        'JD FOR MISPOSTINGS Refunds for Vendor
        ParentID = ParentID + 1
        ChildID = ParentID
        Details = "Refunds - " & VendID

        ClsGlobal.InsertTransactionData(ParentID, "JC", "T9", SysDate, Details, TAmt - PDAmt, 0, MyDate, DTime, "-", "N", "N")

        'Debit NC Account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "49999", 0, TAmt - PDAmt, "T", "REFUND", "", AutoNo, DTime)

        'JC for bank
        ParentID = ParentID + 1
        ChildID = ParentID
        Details = "Refunds - " & VendID

        ClsGlobal.InsertTransactionData(ParentID, "JD", "T9", SysDate, Details, TAmt - PDAmt, 0, MyDate, DTime)

        'Credit NC Account
        'Application("DEFAULTBANK") = "70200"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "70200", TAmt - PDAmt, 0, "B", "70200", "", AutoNo, DTime)

        If (ClsGlobal.ExecuteCommand() = False) Then
            MsgBox("Transaction Faild......     ", MsgBoxStyle.Information)
            Exit Sub
        End If

        ClsGlobal.PaystatusFlag(ReserveParentID, "S", DTime)

        If (ClsGlobal.ExecuteCommand() = True) Then

            Dim MsgStr As String
            MsgStr = "Invoice has been refunded.   " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose(True)

        End If

    End Sub

    Private Sub DataGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.Click
        If DataGrid1.VisibleRowCount > 0 Then
            btnOk.Enabled = True
            CRow = DataGrid1.CurrentRowIndex
        End If
    End Sub

End Class
