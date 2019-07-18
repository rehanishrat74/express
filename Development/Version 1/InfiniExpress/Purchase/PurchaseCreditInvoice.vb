Imports System
Imports System.Windows.Forms.CurrencyManager
Imports InfiniExpress.BLL
Imports InfiniExpress.DataGridTextBoxCombo

Public Class PurchaseCreditInvoice

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
        taxRatesArray = New Double() {1, 1, 1, 1, 1, 1, 1, 1, 1}
        MDI.numOfForms = 0
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCal As System.Windows.Forms.Button
    Friend WithEvents Cal As System.Windows.Forms.MonthCalendar
    Friend WithEvents txtcal As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtvat As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtnet As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents txtcustname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtGAmt As System.Windows.Forms.TextBox
    Friend WithEvents cmbVendorID As System.Windows.Forms.ComboBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Txtdate As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PurchaseCreditInvoice))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Txtdate = New System.Windows.Forms.TextBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnCal = New System.Windows.Forms.Button()
        Me.Cal = New System.Windows.Forms.MonthCalendar()
        Me.txtcal = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtGAmt = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtvat = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtnet = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.txtcustname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbVendorID = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Txtdate, Me.btnPrint, Me.btnCal, Me.Cal, Me.txtcal, Me.Label6, Me.Panel2, Me.btnOk, Me.Button1, Me.DataGrid1, Me.txtcustname, Me.Label2, Me.cmbVendorID, Me.Label1})
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(6, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(522, 300)
        Me.Panel1.TabIndex = 2
        '
        'Txtdate
        '
        Me.Txtdate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Txtdate.Location = New System.Drawing.Point(200, 264)
        Me.Txtdate.Name = "Txtdate"
        Me.Txtdate.Size = New System.Drawing.Size(96, 13)
        Me.Txtdate.TabIndex = 15
        Me.Txtdate.Text = ""
        Me.Txtdate.Visible = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnPrint.Enabled = False
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrint.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.ForeColor = System.Drawing.Color.White
        Me.btnPrint.Location = New System.Drawing.Point(144, 226)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(60, 20)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "Print"
        '
        'btnCal
        '
        Me.btnCal.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.btnCal.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCal.ForeColor = System.Drawing.Color.Navy
        Me.btnCal.Location = New System.Drawing.Point(488, 14)
        Me.btnCal.Name = "btnCal"
        Me.btnCal.Size = New System.Drawing.Size(18, 22)
        Me.btnCal.TabIndex = 3
        Me.btnCal.Text = "V"
        '
        'Cal
        '
        Me.Cal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal.Location = New System.Drawing.Point(300, 36)
        Me.Cal.Name = "Cal"
        Me.Cal.TabIndex = 14
        '
        'txtcal
        '
        Me.txtcal.BackColor = System.Drawing.Color.White
        Me.txtcal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcal.Enabled = False
        Me.txtcal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcal.ForeColor = System.Drawing.Color.Navy
        Me.txtcal.Location = New System.Drawing.Point(402, 14)
        Me.txtcal.Name = "txtcal"
        Me.txtcal.ReadOnly = True
        Me.txtcal.Size = New System.Drawing.Size(89, 21)
        Me.txtcal.TabIndex = 2
        Me.txtcal.Text = ""
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label6.Location = New System.Drawing.Point(320, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Invoice Date"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtGAmt, Me.Label5, Me.txtvat, Me.Label4, Me.txtnet, Me.Label3})
        Me.Panel2.Location = New System.Drawing.Point(324, 222)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(186, 70)
        Me.Panel2.TabIndex = 7
        '
        'txtGAmt
        '
        Me.txtGAmt.BackColor = System.Drawing.Color.White
        Me.txtGAmt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGAmt.Enabled = False
        Me.txtGAmt.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGAmt.ForeColor = System.Drawing.Color.Navy
        Me.txtGAmt.Location = New System.Drawing.Point(92, 44)
        Me.txtGAmt.Name = "txtGAmt"
        Me.txtGAmt.Size = New System.Drawing.Size(88, 21)
        Me.txtGAmt.TabIndex = 8
        Me.txtGAmt.Text = "0.00"
        Me.txtGAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label5.Location = New System.Drawing.Point(2, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 18)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Gross Amount"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtvat
        '
        Me.txtvat.BackColor = System.Drawing.Color.White
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Enabled = False
        Me.txtvat.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtvat.ForeColor = System.Drawing.Color.Navy
        Me.txtvat.Location = New System.Drawing.Point(92, 22)
        Me.txtvat.Name = "txtvat"
        Me.txtvat.Size = New System.Drawing.Size(88, 21)
        Me.txtvat.TabIndex = 6
        Me.txtvat.Text = "0.00"
        Me.txtvat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(92, 2)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 18)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Total VAT"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtnet
        '
        Me.txtnet.BackColor = System.Drawing.Color.White
        Me.txtnet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtnet.Enabled = False
        Me.txtnet.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnet.ForeColor = System.Drawing.Color.Navy
        Me.txtnet.Location = New System.Drawing.Point(2, 22)
        Me.txtnet.Name = "txtnet"
        Me.txtnet.Size = New System.Drawing.Size(88, 21)
        Me.txtnet.TabIndex = 4
        Me.txtnet.Text = "0.00"
        Me.txtnet.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(2, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 18)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Total Net"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.Location = New System.Drawing.Point(14, 226)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 20)
        Me.btnOk.TabIndex = 4
        Me.btnOk.Text = "Save"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(80, 226)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 20)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "&Close"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(14, 72)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(498, 147)
        Me.DataGrid1.TabIndex = 7
        '
        'txtcustname
        '
        Me.txtcustname.BackColor = System.Drawing.Color.White
        Me.txtcustname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustname.Enabled = False
        Me.txtcustname.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcustname.ForeColor = System.Drawing.Color.Navy
        Me.txtcustname.Location = New System.Drawing.Point(116, 38)
        Me.txtcustname.Name = "txtcustname"
        Me.txtcustname.ReadOnly = True
        Me.txtcustname.Size = New System.Drawing.Size(222, 21)
        Me.txtcustname.TabIndex = 1
        Me.txtcustname.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(14, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Vendor Name"
        '
        'cmbVendorID
        '
        Me.cmbVendorID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVendorID.Location = New System.Drawing.Point(116, 16)
        Me.cmbVendorID.Name = "cmbVendorID"
        Me.cmbVendorID.Size = New System.Drawing.Size(130, 21)
        Me.cmbVendorID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Vendor ID"
        '
        'PurchaseCreditInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(534, 309)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PurchaseCreditInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Purchase Credit Notes"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Define Local Variable"
    Private okToValidate As Boolean = True
    Dim dr As IDataReader
    Public clsPurchase As New ClassPurchases()
    Dim ClsSales As New ClassSales()
    Dim ClsGlobal As New ClassGlobal()
    Dim ComboTextCol As New DataGridComboBoxColumn()
    Dim myds As New DataSet()
    Dim AutoNo As Integer
    Dim ParentNo As Integer
    Dim ParentIDs As String
    Dim storeRowNo As Integer
    Dim numCallCurrentChaged As Integer
    Dim numLoad As Integer
    Dim maintainCC As Integer
    Dim VendId As String, InvDate As String, TranDate As Date
#End Region

    Private Sub PurchaseCreditInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ParentIDs = "" ' Initialization
        dr = clsPurchase.LoadVendorID
        While dr.Read
            cmbVendorID.Items.Add(dr.Item("Vendorid"))
        End While

        Cal.Visible = False
        Dim dt As New DataTable("mytable")
        Dim i As Integer
        dt.Columns.Add(New DataColumn("TNo", GetType(Integer)))
        dt.Columns.Add(New DataColumn("Service Description", GetType(String)))
        dt.Columns.Add(New DataColumn("TaxID", GetType(String)))
        dt.Columns.Add(New DataColumn("Net", GetType(Double)))
        dt.Columns.Add(New DataColumn("VAT", GetType(Double)))
        dt.Columns.Add(New DataColumn("Amount", GetType(Double)))

        Dim tableStyle As New DataGridTableStyle()

        tableStyle.MappingName = "mytable"
        myds = ClsGlobal.TaxCodeSummary()

        i = 0
        Dim TextCol As New DataGridTextBoxColumn()
        TextCol.Width = 30
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 1
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 185
        TextCol.TextBox.Focus()
        TextCol.TextBox.MaxLength = 50
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 2 ' THE COMBOBOX CONTROL
        ComboTextCol.MappingName = "TaxID" 'must be from the grid table...
        ComboTextCol.HeaderText = "TaxID"
        ComboTextCol.Width = 45
        ComboTextCol.NullText = "T1"
        ComboTextCol.ColumnComboBox.DataSource = myds.Tables(0).DefaultView 'dv;
        ComboTextCol.ColumnComboBox.DisplayMember = "TaxID"
        ComboTextCol.ColumnComboBox.ValueMember = "CountID"
        tableStyle.PreferredRowHeight = ComboTextCol.ColumnComboBox.Height + 2
        tableStyle.GridColumnStyles.Add(ComboTextCol)

        i = 3
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 80
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 12
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName & " " & Chr(9)
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 4
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 72
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 10
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName & " " & Chr(9)
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 5
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 80
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 10
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName & " " & Chr(9)
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        tableStyle.RowHeadersVisible = False

        i = 0
        While i < 5
            Dim dr As DataRow = dt.NewRow()
            dr(0) = Convert.ToString(i + 1)
            dr(1) = String.Empty
            dr(3) = 0
            dr(4) = 0
            dr(5) = 0
            dt.Rows.Add(dr)
            i += 1
        End While

        DataGrid1.TableStyles.Clear()
        DataGrid1.TableStyles.Add(tableStyle)
        dt.DefaultView.AllowNew = False
        DataGrid1.DataSource = dt
        numLoad = 1

    End Sub

    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
        If numLoad = 1 Then
            storeRowNo = DataGrid1.CurrentRowIndex
        End If

        numCallCurrentChaged += 1
        If numCallCurrentChaged > 1 Then
            numCallCurrentChaged = 0
            If (DataGrid1.CurrentCell.ColumnNumber = 2) Then
                Global.rowNumber = DataGrid1.CurrentCell.RowNumber
            End If
            Exit Sub
        End If

        If (DataGrid1.CurrentCell.ColumnNumber = 2) Then
            Global.rowNumber = DataGrid1.CurrentCell.RowNumber
            'Exit Sub
        End If

        maintainCC = DataGrid1.CurrentRowIndex
        DataGrid1.Item(storeRowNo, 4) = Format(DataGrid1.Item(storeRowNo, 3) * CType(Convert.ToDouble(myds.Tables(0).Rows(taxRatesArray(storeRowNo))(2)) / 100, Double), "0.00")
        DataGrid1.Item(storeRowNo, 5) = Format(DataGrid1.Item(storeRowNo, 4) + DataGrid1.Item(storeRowNo, 3), "0.00")

        numCallCurrentChaged += 1
        DataGrid1.Item(maintainCC, 4) = Format(DataGrid1.Item(maintainCC, 3) * CType(Convert.ToDouble(myds.Tables(0).Rows(taxRatesArray(maintainCC))(2)) / 100, Double), "0.00")
        DataGrid1.Item(maintainCC, 5) = Format(DataGrid1.Item(maintainCC, 4) + DataGrid1.Item(maintainCC, 3), "0.00")

        Dim i As Integer = 0
        Dim NetValue As Double = 0
        Dim VatValue As Double = 0
        Dim Amt As Double = 0

        For i = 0 To 4
            NetValue += DataGrid1.Item(i, 3)
            VatValue += DataGrid1.Item(i, 4)
            Amt += DataGrid1.Item(i, 5)
        Next

        txtGAmt.Text = Format(VatValue + NetValue, "0.00")
        txtvat.Text = Format(VatValue, "0.00")
        txtnet.Text = Format(NetValue, "0.00")
        storeRowNo = DataGrid1.CurrentRowIndex
        numCallCurrentChaged = 0
        numLoad = 0

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCal.Click
        If Cal.Visible = False Then
            Cal.Visible = True
        Else
            Cal.Visible = False
        End If
    End Sub

    Private Sub cmbVendorID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVendorID.SelectedIndexChanged
        Dim Vendid As String = (cmbVendorID.SelectedItem)
        dr = clsPurchase.GetVendorDetails(Vendid) ' Change it here
        dr.Read()
        If IsDBNull(dr("Vendorname")) = False Then txtcustname.Text = dr("Vendorname")
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim TempRow As Int32, pInvno As Integer, Inc As Integer = 1, j As Integer
        Dim TempCol As Int32
        Dim Desc As String
        Dim TaxID As String
        Dim Net As Double, netAmt As Double = 0
        Dim Vat As Double
        Dim Amt As Double
        Dim InvoiceExists As Boolean = False
        Dim PdfCollection As New Hashtable

        If CDbl(DataGrid1.Item(DataGrid1.CurrentCell.RowNumber, 5)) = 0 Then
            Dim Drow As Integer = DataGrid1.CurrentCell.RowNumber
            DataGrid1.Item(Drow, 4) = Format(DataGrid1.Item(Drow, 3) * CType(Convert.ToDouble(myds.Tables(0).Rows(taxRatesArray(Drow))(2)) / 100, Double), "0.00")
            DataGrid1.Item(Drow, 5) = Format(DataGrid1.Item(Drow, 4) + DataGrid1.Item(Drow, 3), "0.00")

            Dim i As Integer = 0
            Dim NetValue As Double = 0
            Dim VatValue As Double = 0

            For i = 0 To 4
                NetValue += DataGrid1.Item(i, 3)
                VatValue += DataGrid1.Item(i, 4)
                Amt += DataGrid1.Item(i, 5)
            Next

            txtGAmt.Text = Format(VatValue + NetValue, "0.00")
            txtvat.Text = Format(VatValue, "0.00")
            txtnet.Text = Format(NetValue, "0.00")
        End If

        If cmbVendorID.Text = "" Then
            MsgBox("Select vendor ID first.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        If txtcal.Text = "" Then
            MsgBox("Select invoice date first.    ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        For TempRow = 0 To 4
            If DataGrid1.Item(TempRow, 3) <> 0 Then
                InvoiceExists = True
                Exit For
            End If
        Next
        If InvoiceExists = False Then
            MessageBox.Show("Atleast one invoice should be entered.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For TempRow = 0 To 4
            If CStr(DataGrid1.Item(TempRow, 1)).IndexOf("'"c) <> -1 Then
                MessageBox.Show("Character "" ' "" is not allowed.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        VendId = cmbVendorID.Text
        InvDate = txtcal.Text
        TranDate = CDate(Txtdate.Text)

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        AutoNo = ClsGlobal.GetLedgerAutoNo
        ParentNo = ClsGlobal.GetTransactionNo()

        For TempRow = 0 To 4

            If DataGrid1.Item(TempRow, 3) <> 0 Then

                Desc = DataGrid1.Item(TempRow, 1)
                TaxID = Convert.ToString(myds.Tables(0).Rows(taxRatesArray(TempRow))(1))
                Net = DataGrid1.Item(TempRow, 3)
                Vat = DataGrid1.Item(TempRow, 4)
                Amt = DataGrid1.Item(TempRow, 5)

                'Adding item in collection
                PdfCollection.Add(Inc, ParentNo)

                'Inserting Data In Ledger,Tansaction & Outstading Table
                TransactionAndLedgerTable(VendId, TaxID, InvDate, Desc, Net, Vat, TranDate, DTime)
                ParentNo += 1
                AutoNo += 1
                Inc += 1
                netAmt = Format(netAmt + (Net + Vat), "0.00")
            End If
        Next

        'Get PdfInvno From Tblpdfinvoices
        pInvno = ClsSales.getPdfinvNo()

        'Insert data in tblpdfinvoices table
        ClsSales.Insertpdfinvoice(" ", TranDate, "PC", pInvno, VendId, netAmt, DTime)
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

        'Update Groupno field in Tblpdfinvoices table
        For j = 1 To PdfCollection.Count
            ClsSales.updateGroupno(pInvno, CInt(PdfCollection.Item(j)))
        Next

        If (ClsGlobal.ExecuteCommand() = True) Then

            Dim MsgStr As String
            MsgStr = "Purchase credit notes have been created. Now you can print it.    " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnPrint.Enabled = True
            btnOk.Enabled = False

        End If

    End Sub

    Sub TransactionAndLedgerTable(ByVal TransId As String, ByVal TaxId As String, ByVal InvDate As String, _
                              ByVal InvDetails As String, ByVal InvNet As Double, ByVal InvVat As Double, _
                              ByVal TranDate As Date, ByVal TDateTime As String)

        If ParentIDs.Length = 0 Then
            ParentIDs = Convert.ToString(ParentNo)
        Else
            ParentIDs = ParentIDs & "," & Convert.ToString(ParentNo)
        End If

        'Insert Data In TblTransaction
        ClsGlobal.InsertTransactionData(ParentNo, "PC", TaxId, InvDate, InvDetails, InvNet, InvVat, TranDate, TDateTime, "N")

        'Insert Data In TblLedger
        ClsGlobal.InsertLedger(ParentNo, ParentNo, "20000", 0, InvNet, "U", TransId, "", AutoNo, TDateTime)

        'Debtor Control Account For Ledger Transaction
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentNo, ParentNo, "71000", InvNet + InvVat, 0, "T", "20000", "", AutoNo, TDateTime)

        'Sales Tax Control Account Ledger Id Transaction
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentNo, ParentNo, "71101", 0, InvVat, "T", "20000", "", AutoNo, TDateTime)

        'Customer Account Transaction
        AutoNo = AutoNo + 1
        ClsGlobal.InsertLedger(ParentNo, ParentNo, TransId, InvNet + InvVat, 0, "S", "20000", "*", AutoNo, TDateTime)

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim objShowInvoiceDetail As New ShowInvoiceDetail
        objShowInvoiceDetail.MakeSalesInvoiceDetail(Convert.ToString(cmbVendorID.SelectedItem), ParentIDs, "Purchase Credit Invoice")
        objShowInvoiceDetail.Show()
        Me.Hide()
        objShowInvoiceDetail.Activate()
    End Sub

    Private Sub Cal_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Cal.DateSelected
        txtcal.Text = Format(Cal.SelectionStart, "dd/MM/yyyy")
        Txtdate.Text = Cal.SelectionStart
        Cal.Visible = False
    End Sub
End Class
