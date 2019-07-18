Imports System
Imports System.Data
Imports System.Windows.Forms.CurrencyManager
Imports InfiniExpress.BLL
Imports InfiniExpress.DataGridTextBoxCombo

Public Class SalesCreditInvoice

    Inherits System.Windows.Forms.Form

#Region "Define Local Varibal"

    Public Dr As IDataReader
    Private okToValidate As Boolean = True
    Public clsSales As New ClassSales()
    Dim GlbGlobal As New ClassGlobal()
    Dim ComboTextCol As New DataGridComboBoxColumn()
    Dim myds As New DataSet()
    Dim AutoNo As Integer
    Dim ParentNo As Integer
    Dim ParentIDs As String
    Dim storeRowNo As Integer
    Dim numCallCurrentChaged As Integer
    Dim numLoad As Integer
    Dim maintainCC As Integer
#End Region

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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents txtcustname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCal As System.Windows.Forms.Button
    Friend WithEvents Cal As System.Windows.Forms.MonthCalendar
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtcal As System.Windows.Forms.TextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents txtGamt As System.Windows.Forms.TextBox
    Friend WithEvents txtvat As System.Windows.Forms.TextBox
    Friend WithEvents txtnet As System.Windows.Forms.TextBox
    Friend WithEvents btncancel As System.Windows.Forms.Button
    Friend WithEvents cmbid As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Txtdate As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SalesCreditInvoice))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Txtdate = New System.Windows.Forms.TextBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnCal = New System.Windows.Forms.Button()
        Me.Cal = New System.Windows.Forms.MonthCalendar()
        Me.txtcal = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.txtcustname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbid = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtGamt = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtvat = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtnet = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btncancel = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Txtdate, Me.btnPrint, Me.btnCal, Me.Cal, Me.txtcal, Me.Label6, Me.DataGrid1, Me.txtcustname, Me.Label2, Me.cmbid, Me.Label1, Me.btnOk, Me.Panel3, Me.btncancel})
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(528, 306)
        Me.Panel1.TabIndex = 1
        '
        'Txtdate
        '
        Me.Txtdate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Txtdate.Location = New System.Drawing.Point(215, 271)
        Me.Txtdate.Name = "Txtdate"
        Me.Txtdate.Size = New System.Drawing.Size(96, 13)
        Me.Txtdate.TabIndex = 16
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
        Me.btnPrint.Location = New System.Drawing.Point(141, 232)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(60, 18)
        Me.btnPrint.TabIndex = 7
        Me.btnPrint.Text = "&Print"
        '
        'btnCal
        '
        Me.btnCal.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.btnCal.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCal.ForeColor = System.Drawing.Color.Navy
        Me.btnCal.Location = New System.Drawing.Point(490, 14)
        Me.btnCal.Name = "btnCal"
        Me.btnCal.Size = New System.Drawing.Size(18, 20)
        Me.btnCal.TabIndex = 3
        Me.btnCal.Text = "V"
        '
        'Cal
        '
        Me.Cal.BackColor = System.Drawing.Color.White
        Me.Cal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal.Location = New System.Drawing.Point(304, 34)
        Me.Cal.Name = "Cal"
        Me.Cal.TabIndex = 14
        Me.Cal.TabStop = False
        Me.Cal.Visible = False
        '
        'txtcal
        '
        Me.txtcal.BackColor = System.Drawing.Color.White
        Me.txtcal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcal.Enabled = False
        Me.txtcal.ForeColor = System.Drawing.Color.Navy
        Me.txtcal.Location = New System.Drawing.Point(412, 14)
        Me.txtcal.Name = "txtcal"
        Me.txtcal.ReadOnly = True
        Me.txtcal.Size = New System.Drawing.Size(80, 20)
        Me.txtcal.TabIndex = 2
        Me.txtcal.Text = ""
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(325, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Invoice Date"
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowNavigation = False
        Me.DataGrid1.AllowSorting = False
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.White
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.Silver
        Me.DataGrid1.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.ForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.Location = New System.Drawing.Point(14, 72)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.ParentRowsVisible = False
        Me.DataGrid1.RowHeadersVisible = False
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.DataGrid1.Size = New System.Drawing.Size(498, 148)
        Me.DataGrid1.TabIndex = 4
        '
        'txtcustname
        '
        Me.txtcustname.BackColor = System.Drawing.Color.White
        Me.txtcustname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustname.Enabled = False
        Me.txtcustname.ForeColor = System.Drawing.Color.Navy
        Me.txtcustname.Location = New System.Drawing.Point(116, 38)
        Me.txtcustname.Name = "txtcustname"
        Me.txtcustname.ReadOnly = True
        Me.txtcustname.Size = New System.Drawing.Size(222, 20)
        Me.txtcustname.TabIndex = 1
        Me.txtcustname.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(14, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Customer Name"
        '
        'cmbid
        '
        Me.cmbid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbid.Location = New System.Drawing.Point(116, 16)
        Me.cmbid.Name = "cmbid"
        Me.cmbid.Size = New System.Drawing.Size(100, 21)
        Me.cmbid.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(14, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer ID"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.Location = New System.Drawing.Point(14, 232)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 18)
        Me.btnOk.TabIndex = 5
        Me.btnOk.Text = "&Save"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtGamt, Me.Label7, Me.txtvat, Me.Label8, Me.txtnet, Me.Label9})
        Me.Panel3.Location = New System.Drawing.Point(328, 224)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(186, 70)
        Me.Panel3.TabIndex = 7
        '
        'txtGamt
        '
        Me.txtGamt.BackColor = System.Drawing.Color.White
        Me.txtGamt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGamt.Enabled = False
        Me.txtGamt.ForeColor = System.Drawing.Color.Navy
        Me.txtGamt.Location = New System.Drawing.Point(92, 44)
        Me.txtGamt.Name = "txtGamt"
        Me.txtGamt.ReadOnly = True
        Me.txtGamt.Size = New System.Drawing.Size(90, 20)
        Me.txtGamt.TabIndex = 8
        Me.txtGamt.Text = "0.00"
        Me.txtGamt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(2, 44)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 18)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Gross Amount"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtvat
        '
        Me.txtvat.BackColor = System.Drawing.Color.White
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Enabled = False
        Me.txtvat.ForeColor = System.Drawing.Color.Navy
        Me.txtvat.Location = New System.Drawing.Point(92, 22)
        Me.txtvat.Name = "txtvat"
        Me.txtvat.ReadOnly = True
        Me.txtvat.Size = New System.Drawing.Size(90, 20)
        Me.txtvat.TabIndex = 6
        Me.txtvat.Text = "0.00"
        Me.txtvat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(92, 2)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 18)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Total VAT"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtnet
        '
        Me.txtnet.BackColor = System.Drawing.Color.White
        Me.txtnet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtnet.Enabled = False
        Me.txtnet.ForeColor = System.Drawing.Color.Navy
        Me.txtnet.Location = New System.Drawing.Point(2, 22)
        Me.txtnet.Name = "txtnet"
        Me.txtnet.ReadOnly = True
        Me.txtnet.Size = New System.Drawing.Size(88, 20)
        Me.txtnet.TabIndex = 4
        Me.txtnet.Text = "0.00"
        Me.txtnet.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(2, 2)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 18)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Total Net"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btncancel
        '
        Me.btncancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.btncancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.Location = New System.Drawing.Point(78, 232)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.Size = New System.Drawing.Size(60, 18)
        Me.btncancel.TabIndex = 6
        Me.btncancel.Text = "&Close"
        '
        'SalesCreditInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(536, 315)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SalesCreditInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Sales Credit Notes"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCal.Click
        If Cal.Visible = False Then
            Cal.Visible = True
        Else
            Cal.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose(True)
    End Sub

    Private Sub SalesCreditInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ParentIDs = "" ' Initialization
        dr = clsSales.LoadCustomerID
        While dr.Read
            cmbid.Items.Add(dr.Item("Customerid"))
        End While

        Dim dt As New DataTable("mytable")
        Dim i As Integer
        dt.Columns.Add(New DataColumn("TNo", GetType(Integer)))
        dt.Columns.Add(New DataColumn("Sales Description", GetType(String)))
        dt.Columns.Add(New DataColumn("TaxID", GetType(String)))
        dt.Columns.Add(New DataColumn("Net", GetType(Double)))
        dt.Columns.Add(New DataColumn("VAT", GetType(Double)))
        dt.Columns.Add(New DataColumn("Amount", GetType(Double)))

        Dim tableStyle As New DataGridTableStyle()

        tableStyle.MappingName = "mytable"

        myds = GlbGlobal.TaxCodeSummary()

        i = 0
        Dim TextCol As New DataGridTextBoxColumn()
        TextCol.Width = 30
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 1
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 198
        TextCol.TextBox.Focus()
        TextCol.TextBox.MaxLength = 50
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 2 ' THE COMBOBOX CONTROL
        ComboTextCol.MappingName = "TaxID" 'must be from the grid table...
        ComboTextCol.HeaderText = "TaxID"
        ComboTextCol.Width = 50
        ComboTextCol.NullText = "T1"
        ComboTextCol.ColumnComboBox.DataSource = myds.Tables(0).DefaultView 'dv;
        ComboTextCol.ColumnComboBox.DisplayMember = "TaxID"
        ComboTextCol.ColumnComboBox.ValueMember = "CountID"
        tableStyle.PreferredRowHeight = ComboTextCol.ColumnComboBox.Height + 2
        tableStyle.PreferredColumnWidth = 50
        tableStyle.GridColumnStyles.Add(ComboTextCol)

        i = 3
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 70
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = 0
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 12
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName & " " & Chr(9)
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 4
        TextCol = New DataGridTextBoxColumn()
        TextCol.Width = 65
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = 0
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
        TextCol.TextBox.Text = 0
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 14
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
            dr(3) = "0.00"
            dr(4) = "0.00"
            dr(5) = "0.00"
            dt.Rows.Add(dr)
            i += 1
        End While

        DataGrid1.TableStyles.Clear()
        DataGrid1.TableStyles.Add(tableStyle)
        dt.DefaultView.AllowNew = False
        DataGrid1.DataSource = dt
        numLoad = 1

    End Sub

    Private Sub DataGrid1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGrid1.MouseMove

        Dim hti As DataGrid.HitTestInfo = DataGrid1.HitTest(New Point(e.X, e.Y))
        If hti.Type = DataGrid.HitTestType.ColumnResize Then
            Return
        End If

    End Sub

    Private Sub DataGrid1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGrid1.MouseDown

        Dim hti As DataGrid.HitTestInfo = DataGrid1.HitTest(New Point(e.X, e.Y))
        If hti.Type = DataGrid.HitTestType.ColumnResize Then
            Return
        End If

    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim TempRow As Int32, pInvno As Integer, Inc As Integer = 1, j As Integer
        Dim TempCol As Int32
        Dim Desc As String
        Dim TaxID As String
        Dim Net As Double, NetAmt As Double = 0
        Dim Vat As Double
        Dim Amt As Double
        Dim CustId As String, InvDate As String, TranDate As Date
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

            txtGamt.Text = Format(VatValue + NetValue, "0.00")
            txtvat.Text = Format(VatValue, "0.00")
            txtnet.Text = Format(NetValue, "0.00")
        End If

        If cmbid.Text = "" Then
            MsgBox("Select Customer ID first.    ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        If txtcal.Text = "" Then
            MsgBox("Select invoice date first.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        For TempRow = 0 To 4
            If DataGrid1.Item(TempRow, 3) <> 0 Then
                InvoiceExists = True
                Exit For
            End If
        Next

        If InvoiceExists = False Then
            MessageBox.Show("Atleast one sales credit note should be entered.   ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For TempRow = 0 To 4
            If CStr(DataGrid1.Item(TempRow, 1)).IndexOf("'"c) <> -1 Then
                MessageBox.Show("Character "" ' "" is not allowed.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        CustId = cmbid.Text
        InvDate = txtcal.Text
        TranDate = CDate(Txtdate.Text)

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        AutoNo = GlbGlobal.GetLedgerAutoNo
        ParentNo = GlbGlobal.GetTransactionNo()

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
                TransactionAndLedgerTable(CustId, TaxID, InvDate, Desc, Net, Vat, TranDate, DTime)
                ParentNo += 1
                AutoNo += 1
                Inc += 1
                NetAmt = Format(NetAmt + (Net + Vat), "0.00")
            End If
        Next

        'Get PdfInvno From Tblpdfinvoices
        pInvno = clsSales.getPdfinvNo()

        'Insert data in tblpdfinvoices table
        clsSales.Insertpdfinvoice(" ", TranDate, "SC", pInvno, CustId, NetAmt, DTime)
        If GlbGlobal.ExecuteCommand = False Then Exit Sub

        'Update Groupno field in Tblpdfinvoices table
        For j = 1 To PdfCollection.Count
            clsSales.updateGroupno(pInvno, CInt(PdfCollection.Item(j)))
        Next

        If (GlbGlobal.ExecuteCommand() = True) Then

            If IsService = True Then
                'Insert Data In Credit Card info Table
                clsSales.insertCreditCardInfoFORSC(pInvno, TranDate, NetAmt, DTime, "@ACC" & CustId, "Credit Notes From Infini Accounts Express", "Credit Note", "CB", cmbid.Text)
                If ClassGlobal.ExecuteCommand = False Then Exit Sub
            End If

            Dim MsgStr As String
            MsgStr = "Sales credit notes have been created. Now you can print it.    " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnPrint.Enabled = True
            btnOk.Enabled = False

        End If

    End Sub

    Sub TransactionAndLedgerTable(ByVal TransId As String, ByVal TaxId As String, ByVal InvDate As String, _
                               ByVal InvDetails As String, ByVal InvNet As Double, ByVal InvVat As Double, _
                               ByVal TranDate As Date, ByVal TranDateTime As String)

        If ParentIDs.Length = 0 Then
            ParentIDs = Convert.ToString(ParentNo)
        Else
            ParentIDs = ParentIDs & "," & Convert.ToString(ParentNo)
        End If

        'Insert Data In TblTransaction
        GlbGlobal.InsertTransactionData(ParentNo, "SC", TaxId, InvDate, InvDetails, InvNet, InvVat, TranDate, TranDateTime, "N")

        'Insert Data In TblLedger
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "10000", InvNet, 0, "U", TransId, "", AutoNo, TranDateTime)

        'Debtor Control Account For Ledger Transaction
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "70100", 0, InvNet + InvVat, "T", "10000", "", AutoNo, TranDateTime)

        'Sales Tax Control Account Ledger Id Transaction
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "71100", InvVat, 0, "T", "10000", "", AutoNo, TranDateTime)

        'Customer Account Transaction
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, TransId, 0, InvNet + InvVat, "C", "10000", "*", AutoNo, TranDateTime)

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

        txtGamt.Text = Format(VatValue + NetValue, "0.00")
        txtvat.Text = Format(VatValue, "0.00")
        txtnet.Text = Format(NetValue, "0.00")
        storeRowNo = DataGrid1.CurrentRowIndex
        numCallCurrentChaged = 0
        numLoad = 0
    End Sub

    Private Sub CmbID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbid.SelectedIndexChanged
        Dim cid As String = (cmbid.SelectedItem)
        Dr = clsSales.GetCustInfo(cid)
        Dr.Read()
        If IsDBNull(Dr("Customername")) = False Then txtcustname.Text = Dr("Customername")
    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim objShowInvoiceDetail As New ShowInvoiceDetail
        objShowInvoiceDetail.MakeSalesInvoiceDetail(Convert.ToString(cmbid.SelectedItem), ParentIDs, "Sales Credit Invoice")
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
