Imports Microsoft.VisualBasic
Imports InfiniExpress.BLL

Public Class CustomerRefund

    Inherits System.Windows.Forms.Form
    Public ClsSales As New ClassSales()
    Public ClsGlobal As New ClassGlobal
    Public InfiniWebService As New InfiniExpress.InfiniWebService.InfiniService
    Dim CRow As Integer, AutoNo As Integer
    Dim CustID As String
    Dim Dr As IDataReader
    Dim Ds As New DataSet


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
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Protected WithEvents Button4 As System.Windows.Forms.Button
    Protected WithEvents Panel2 As System.Windows.Forms.Panel
    Protected WithEvents TextBox2 As System.Windows.Forms.TextBox
    Protected WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Protected WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn5 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn6 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn7 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn8 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn9 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn10 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn11 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn12 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents DataGridTextBoxColumn13 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents cmbcid As System.Windows.Forms.ComboBox
    Protected WithEvents btnOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CustomerRefund))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn5 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn6 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn7 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn8 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn9 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn10 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn11 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn12 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn13 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbcid = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button4, Me.btnOk, Me.DataGrid1, Me.TextBox1, Me.Label2, Me.cmbcid, Me.Label1})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(532, 242)
        Me.Panel1.TabIndex = 0
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(79, 216)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(60, 18)
        Me.Button4.TabIndex = 34
        Me.Button4.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.Enabled = False
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(10, 216)
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
        Me.DataGrid1.Location = New System.Drawing.Point(10, 38)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(512, 168)
        Me.DataGrid1.TabIndex = 4
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn3, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn5, Me.DataGridTextBoxColumn6, Me.DataGridTextBoxColumn7, Me.DataGridTextBoxColumn8, Me.DataGridTextBoxColumn9, Me.DataGridTextBoxColumn10, Me.DataGridTextBoxColumn11, Me.DataGridTextBoxColumn12, Me.DataGridTextBoxColumn13})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = ""
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Sheraz"
        Me.DataGridTextBoxColumn1.MappingName = "123"
        Me.DataGridTextBoxColumn1.Width = 75
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.MappingName = ""
        Me.DataGridTextBoxColumn2.Width = 75
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.MappingName = ""
        Me.DataGridTextBoxColumn3.Width = 75
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.MappingName = ""
        Me.DataGridTextBoxColumn4.Width = 75
        '
        'DataGridTextBoxColumn5
        '
        Me.DataGridTextBoxColumn5.Format = ""
        Me.DataGridTextBoxColumn5.FormatInfo = Nothing
        Me.DataGridTextBoxColumn5.MappingName = ""
        Me.DataGridTextBoxColumn5.Width = 75
        '
        'DataGridTextBoxColumn6
        '
        Me.DataGridTextBoxColumn6.Format = ""
        Me.DataGridTextBoxColumn6.FormatInfo = Nothing
        Me.DataGridTextBoxColumn6.MappingName = ""
        Me.DataGridTextBoxColumn6.Width = 75
        '
        'DataGridTextBoxColumn7
        '
        Me.DataGridTextBoxColumn7.Format = ""
        Me.DataGridTextBoxColumn7.FormatInfo = Nothing
        Me.DataGridTextBoxColumn7.MappingName = ""
        Me.DataGridTextBoxColumn7.Width = 75
        '
        'DataGridTextBoxColumn8
        '
        Me.DataGridTextBoxColumn8.Format = ""
        Me.DataGridTextBoxColumn8.FormatInfo = Nothing
        Me.DataGridTextBoxColumn8.MappingName = ""
        Me.DataGridTextBoxColumn8.Width = 75
        '
        'DataGridTextBoxColumn9
        '
        Me.DataGridTextBoxColumn9.Format = ""
        Me.DataGridTextBoxColumn9.FormatInfo = Nothing
        Me.DataGridTextBoxColumn9.MappingName = ""
        Me.DataGridTextBoxColumn9.Width = 75
        '
        'DataGridTextBoxColumn10
        '
        Me.DataGridTextBoxColumn10.Format = ""
        Me.DataGridTextBoxColumn10.FormatInfo = Nothing
        Me.DataGridTextBoxColumn10.MappingName = ""
        Me.DataGridTextBoxColumn10.Width = 75
        '
        'DataGridTextBoxColumn11
        '
        Me.DataGridTextBoxColumn11.Format = ""
        Me.DataGridTextBoxColumn11.FormatInfo = Nothing
        Me.DataGridTextBoxColumn11.MappingName = ""
        Me.DataGridTextBoxColumn11.Width = 75
        '
        'DataGridTextBoxColumn12
        '
        Me.DataGridTextBoxColumn12.Format = ""
        Me.DataGridTextBoxColumn12.FormatInfo = Nothing
        Me.DataGridTextBoxColumn12.MappingName = ""
        Me.DataGridTextBoxColumn12.Width = 75
        '
        'DataGridTextBoxColumn13
        '
        Me.DataGridTextBoxColumn13.Format = ""
        Me.DataGridTextBoxColumn13.FormatInfo = Nothing
        Me.DataGridTextBoxColumn13.MappingName = ""
        Me.DataGridTextBoxColumn13.Width = 75
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Enabled = False
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(304, 10)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(216, 21)
        Me.TextBox1.TabIndex = 3
        Me.TextBox1.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(202, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Customer Name"
        '
        'cmbcid
        '
        Me.cmbcid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbcid.Location = New System.Drawing.Point(95, 10)
        Me.cmbcid.Name = "cmbcid"
        Me.cmbcid.Size = New System.Drawing.Size(105, 21)
        Me.cmbcid.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(10, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer ID"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBox2, Me.ComboBox2, Me.Label3, Me.Label4, Me.Label5})
        Me.Panel2.Location = New System.Drawing.Point(142, 300)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(312, 138)
        Me.Panel2.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(124, 70)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(180, 21)
        Me.TextBox2.TabIndex = 4
        Me.TextBox2.Text = "TextBox1"
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(124, 40)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(106, 21)
        Me.ComboBox2.TabIndex = 3
        Me.ComboBox2.Text = "ComboBox1"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Navy
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(18, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 20)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Customer Name"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Navy
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(18, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 20)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Customer ID"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(228, 24)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Select Customer To Refund Invoice"
        '
        'CustomerRefund
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(540, 251)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1})
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustomerRefund"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Sales Refund Invoice"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose(True)
    End Sub

    Private Sub CustomerRefund_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dr = ClsSales.LoadCustomerID
        While Dr.Read
            cmbcid.Items.Add(Dr.Item("Customerid"))
        End While
        Dr.Close()
        populatingGrid("")
        Ds = New DataSet()

    End Sub

    Public Function SalesRefund(ByVal Id As String) As DataTable

        Dim Tinvtp As String, Tinvdate As String, Tinvdetails As String
        Dim Tinvnet As Double, Tinvvat As Double, Todebit As Double
        Dim Ttaxid As String, Tpaystatus As String, Ttransid As String
        Dim R3 As String, Paystatus As String, Tid As Integer
        Dim ScAmt As Double, SrAmt As Double, Tparentid As Integer
        Dim Rd As IDataReader

        Dim DT As New DataTable() 'Creating DataTable
        Dim DCTNO As New DataColumn("Parentid", GetType(Integer))
        Dim DCType As New DataColumn("Invtp", GetType(String))
        Dim DCDate As New DataColumn("InvDate", GetType(String))
        Dim DCTDetail As New DataColumn("InvDetails", GetType(String))
        Dim DCNet As New DataColumn("InvNet", GetType(Double))
        Dim DCVat As New DataColumn("InvVat", GetType(Double))
        Dim DCAmt As New DataColumn("ODebit", GetType(Double))
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
        Rd = ClsSales.SalesRefund(CustID)

        If Rd.Read Then
            Tid = Rd("parentid")
            Rd.Close()

            Rd = ClsSales.SalesRefund(CustID)
            While Rd.Read
                If Tid = Rd("Parentid") Then

                    Tparentid = Rd("parentid")
                    Tinvtp = Rd("Invtp")
                    Tinvdate = Rd("Invdate")
                    Tinvdetails = Rd("Invdetails")
                    Tinvnet = Rd("Invnet")
                    Tinvvat = Rd("Invvat")
                    Todebit = SrAmt
                    Ttaxid = Rd("Taxid")
                    Tpaystatus = Rd("paystatus")
                    Ttransid = Rd("transid")

                    If Rd("fromtp") = "SC" Then
                        ScAmt = ScAmt + Rd("Odebit")
                    Else
                        SrAmt = SrAmt + Rd("ODebit")
                    End If

                    GoTo NextRecord 'MOVE TO NEXT RECORD

                Else

                    If SrAmt > 0 Then
                        If ScAmt + SrAmt = Rd("Ldebit") Then
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
                        DR("Odebit") = SrAmt
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
                    Todebit = SrAmt
                    Ttaxid = Rd("Taxid")
                    Tpaystatus = Rd("paystatus")
                    Ttransid = Rd("transid")

                    If Tid = Rd("parentid") Then
                        ScAmt = 0
                        SrAmt = 0
                        If Rd("fromtp") = "SC" Then
                            ScAmt = ScAmt + Rd("Odebit")
                        Else
                            SrAmt = SrAmt + Rd("ODebit")
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
            DR("Odebit") = SrAmt 'Tinvnet + Tinvvat 
            DR("Taxid") = Ttaxid
            DR("Paystatus") = Tpaystatus
            DR("transid") = Ttransid
            DT.Rows.Add(DR) 'Adding DataRow To DataTable
        End If
        Return DT

    End Function

    Private Sub cmbcid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbcid.SelectedIndexChanged

        CustID = (cmbcid.SelectedItem)
        Dr = ClsSales.GetCustInfo(CustID)
        Dr.Read()
        If IsDBNull(Dr("Customername")) = False Then TextBox1.Text = Dr("Customername")
        Dr.Close()
        populatingGrid(CustID)

    End Sub

    Sub populatingGrid(ByVal tempCid As String)

        Ds = New DataSet()
        Dim Dtab As DataTable
        Dtab = SalesRefund(CustID)
        Ds.Merge(Dtab)

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
        End With

        With TranNo
            .HeaderText = "Tran.No"
            .MappingName = "parentid"
            .Width = 50
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
            .HeaderText = "Amount "
            .Alignment = HorizontalAlignment.Right
            .MappingName = "ODebit"
            .Width = 98
            .Format = "#0.00"
            .ReadOnly = True
            .TextBox.ReadOnly = True
            .TextBox.TextAlign = HorizontalAlignment.Right
        End With

        With Net
            .HeaderText = ""
            .Width = 0
            .MappingName = "invnet"
            .TextBox.MaxLength = 10
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
        End With

        With VAT
            .HeaderText = ""
            .Width = 0
            .MappingName = "invvat"
            .TextBox.MaxLength = 10
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

        'GRow = Ds.Tables(0).Rows.Count
        'Bind The Datagrid to Dataset
        If Dtab.DefaultView.Count > 0 Then
            DataGrid1.DataSource = Ds.Tables(0)
            'Stoping Extra Row In DataGrid At The End Of The Row
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
            DgridTstyle.MappingName = Ds.Tables(0).TableName
            DataGrid1.TableStyles.Add(DgridTstyle)
            btnOk.Enabled = False
        End If
    End Sub

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
        Dim CustID As String = DataGrid1.Item(CRow, 9) '"CustomerID"
        Dim ChildID As Double, ParentID As Double, Assparentid As Double
        Dim rightNow As DateTime = DateTime.Now
        Dim SysDate As String, SDAmt As Double, AssumeNo As Double
        Dim MyAuto1 As Integer, OSMaxNo As Integer
        Dim ReserveParentID As Double

        If TParentid = 0 Then Exit Sub
        SysDate = rightNow.ToString("dd/MM/yyyy")
        Dim MyDate As Date = Now
        MyDate = DateSerial(Year(MyDate), Month(MyDate), Day(MyDate))

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        Dr = ClsSales.CustomerInvoiceRefund(TParentid)

        While Dr.Read
            If Dr("invtp") <> "SC" Then
                ClsGlobal.UpdateOutstanding(Dr("myauto"), DTime)
            Else
                Exit Sub
            End If
        End While
        Dr.Close()

        'Calculationg Dicount
        SDAmt = ClsSales.GetDCAgainstTransaction(TParentid, "SD")

        'Get Transaction No
        ParentID = ClsGlobal.GetTransactionNo
        Assparentid = ParentID
        ChildID = ParentID

        'Get Ledger Autono
        AutoNo = ClsGlobal.GetLedgerAutoNo

        'Get New Outstanding Max. No
        OSMaxNo = ClsGlobal.GetOutStandingMaxOSNo()

        If SDAmt > 0 Then
            AssumeNo = Assparentid + 2
        End If

        vat = TVatAmt
        Details = "Refund - " & TParentid

        'Set Refund = y for those invoice who is select by user to REFUND.
        ClsGlobal.ChkRefund("Y", "Y", TParentid, DTime)

        'Step 1
        ClsGlobal.InsertTransactionData(ParentID, "SC", TXid, SysDate, Details, TnetAmt, TVatAmt, MyDate, DTime, "N", "N", "N")

        totrans = Format(TAmt - SDAmt, "0.00") & " To SI " & ParentID + 1
        fromtrans = Format(TAmt - SDAmt, "0.00") & " From SC " & ParentID

        'Step 2
        ClsGlobal.InsertLedger(ParentID, ChildID, CustID, 0, TnetAmt + TVatAmt, "C", "10000", "", AutoNo, DTime)

        'Step3
        'credit outstanding for customer & refund
        ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, CustID, 0, TAmt - SDAmt, "C", totrans, "Refund", SysDate, 0, ParentID + 1, "T", DTime, OSMaxNo)

        ' Special Case For Sale Discount
        If SDAmt > 0 Then
            totrans = Format(SDAmt, "0.00") & " To SI " & AssumeNo
            OSMaxNo = OSMaxNo + 1
            ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, CustID, 0, SDAmt, "C", totrans, "Refund", SysDate, 0, CStr(AssumeNo), "T", DTime, OSMaxNo)
        End If

        'Set Paystatus Flag
        ReserveParentID = ParentID

        'Step4
        'credit for nc debtors control account
        'Application("DEBTORSCONTROLACCOUNT") = "70100"
        'Application("SALESACCOUNT") = "10000"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "70100", 0, TnetAmt + TVatAmt, "T", "10000", "", AutoNo, DTime)

        'Step5
        ''debit sales tax control account
        'Application("SALESACCOUNT") = "10000"
        'Application("SALESTAXCONTROLACCOUNT") = "71100"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "71100", TVatAmt, 0, "T", "10000", "", AutoNo, DTime)

        'Step6
        'Application("SALESACCOUNT") = "10000"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "10000", TnetAmt, 0, "T", "REFUND", "", AutoNo, DTime)

        'allocation-1
        'Get Transaction No
        ParentID = ParentID + 1
        ChildID = ParentID

        Details = "Allocate - " & TParentid

        'Step1
        ClsGlobal.InsertTransactionData(ParentID, "SI", "T9", SysDate, Details, TAmt - SDAmt, 0, MyDate, DTime, "-", "Y", "Y")

        'Step2
        'Debit Customer Account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, CustID, TAmt - SDAmt, 0, "C", "49999", "f", AutoNo, DTime)

        'Step3
        'Debit Outstanding For Customer & Refund
        OSMaxNo = OSMaxNo + 1
        ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, CustID, TAmt - SDAmt, 0, "C", fromtrans, "REFUND", SysDate, ParentID - 1, 0, "F", DTime, OSMaxNo)

        'Step4
        'Debit For NC Debtors Control Account
        'Application("DEBTORSCONTROLACCOUNT") = "70100"
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "70100", TAmt - SDAmt, 0, "T", "49999", "", AutoNo, DTime)

        'Step5
        'Credit Misposting account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "49999", 0, TAmt - SDAmt, "T", "REFUND", "", AutoNo, DTime)

        'Create SD Transaction
        If SDAmt > 0 Then

            'Get Transaction No
            ParentID = ParentID + 1
            ChildID = ParentID

            Details = "Allocate - " & TParentid

            'Step 1
            ClsGlobal.InsertTransactionData(ParentID, "SI", "T9", SysDate, Details, SDAmt, 0, MyDate, DTime, "-", "Y", "Y")

            'Step2
            'Debit Customer Account
            'Application("SALESDISCOUNT") = "10100"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, CustID, SDAmt, 0, "C", "10100", "f", AutoNo, DTime)

            fromtrans = Format(SDAmt, "0.00") & " From SC " & ParentID - 2

            'Step3
            'Debit Outstanding For Customer & Refund
            OSMaxNo = OSMaxNo + 1
            ClsGlobal.InsertOutStanding(AutoNo, ParentID, ChildID, CustID, SDAmt, 0, "C", fromtrans, "REFUND", SysDate, ParentID - 2, 0, "F", DTime, OSMaxNo)

            'Debit For NC Debtors Control Account
            'Application("DEBTORSCONTROLACCOUNT") = "70100"
            'Application("SALESDISCOUNT") = "10100"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, "70100", SDAmt, 0, "T", "10100", "", AutoNo, DTime)

            'Step5
            'Credit Discount
            'Application("SALESDISCOUNT") = "10100"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(ParentID, ChildID, "10100", 0, SDAmt, "T", "REFUND", "", AutoNo, DTime)

        End If

        'JD FOR MISPOSTINGS Refunds for Vendor
        ParentID = ParentID + 1
        ChildID = ParentID
        Details = "Refunds - " & CustID

        ClsGlobal.InsertTransactionData(ParentID, "JC", "T9", SysDate, Details, TAmt - SDAmt, 0, MyDate, DTime, "-", "N", "N")

        'Debit NC Account
        'Application("MISPOSTING") = "49999"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "49999", TAmt - SDAmt, 0, "T", "REFUND", "", AutoNo, DTime)

        'JC for bank
        ParentID = ParentID + 1
        ChildID = ParentID
        Details = "Refunds - " & CustID

        ClsGlobal.InsertTransactionData(ParentID, "JD", "T9", SysDate, Details, TAmt - SDAmt, 0, MyDate, DTime)
        'Credit NC Account
        'Application("DEFAULTBANK") = "70200"
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentID, ChildID, "70200", 0, TAmt - SDAmt, "B", "70200", "", AutoNo, DTime)

        If (ClsGlobal.ExecuteCommand() = False) Then
            MsgBox("Transactions Faild......    ", MsgBoxStyle.Information)
            Exit Sub
        End If

        ClsGlobal.PaystatusFlag(ReserveParentID, "C", DTime)

        If (ClsGlobal.ExecuteCommand() = True) Then

            If IsService = True Then

                'Insert Refund Invoices For Collection service 
                Dim Invds As Boolean, InvNo As Integer
                Dim ChkStr As String
                InvNo = ClsSales.GetRefundedInvNo(TParentid)
                Dim InvDetail As String = "Refund From Infini Accounts Express - " & InvNo

                'Chk Creditcard N0
                ChkStr = ClsSales.checkCCNo(InvNo)

                'Insert Data In CreditcardInfo Table
                ClsSales.SendRefundInvoiceForServiceCollection(InvDetail, "Refund", InvNo, CustID, ChkStr, MyDate, TnetAmt + TVatAmt, DTime)

                ClsGlobal.ExecuteCommand()

            End If

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
