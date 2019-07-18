Imports System
Imports System.Data
Imports System.Windows.Forms.CurrencyManager
Imports InfiniExpress.BLL
Imports InfiniExpress.DataGridTextBoxCombo

Public Class SalesInvoice

    Inherits System.Windows.Forms.Form

#Region "Local Variables"
    Public Dr As IDataReader
    Private okToValidate As Boolean = True
    Public clsSales As New ClassSales()
    Dim GlbGlobal As New ClassGlobal()
    Dim ComboTextCol As New DataGridComboBoxColumn()
    Dim ParentIDs As String
    Dim myds As New DataSet()
    Dim AutoNo As Integer, ParentNo As Integer
    Dim storeRowNo As Integer
    Dim numCallCurrentChaged As Integer
    Dim numLoad As Integer
    Dim maintainCC As Integer
#End Region

    Dim pPtyinvno As Integer
    Dim pPtyAmt As Double
    Dim pPtydate As Date
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
    Protected WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents txtcustname As System.Windows.Forms.TextBox
    Protected WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Protected WithEvents Button1 As System.Windows.Forms.Button
    Protected WithEvents Button2 As System.Windows.Forms.Button
    Protected WithEvents Panel2 As System.Windows.Forms.Panel
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents txtnet As System.Windows.Forms.TextBox
    Protected WithEvents txtvat As System.Windows.Forms.TextBox
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents Cal As System.Windows.Forms.MonthCalendar
    Protected WithEvents btnCal As System.Windows.Forms.Button
    Protected WithEvents txtcal As System.Windows.Forms.TextBox
    Protected WithEvents txtGross As System.Windows.Forms.TextBox
    Protected WithEvents CmbID As System.Windows.Forms.ComboBox
    Protected WithEvents cmbFunctionArea As System.Windows.Forms.ComboBox ' combo box var              
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Txtdate As System.Windows.Forms.TextBox
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SalesInvoice))
        Me.cmbFunctionArea = New System.Windows.Forms.ComboBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Txtdate = New System.Windows.Forms.TextBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnCal = New System.Windows.Forms.Button
        Me.Cal = New System.Windows.Forms.MonthCalendar
        Me.txtcal = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtGross = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtvat = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtnet = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.txtcustname = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CmbID = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbFunctionArea
        '
        Me.cmbFunctionArea.Location = New System.Drawing.Point(0, 0)
        Me.cmbFunctionArea.Name = "cmbFunctionArea"
        Me.cmbFunctionArea.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Txtdate)
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Controls.Add(Me.btnCal)
        Me.Panel1.Controls.Add(Me.Cal)
        Me.Panel1.Controls.Add(Me.txtcal)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.DataGrid1)
        Me.Panel1.Controls.Add(Me.txtcustname)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.CmbID)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(528, 290)
        Me.Panel1.TabIndex = 0
        '
        'Txtdate
        '
        Me.Txtdate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Txtdate.Location = New System.Drawing.Point(216, 344)
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
        Me.btnPrint.Location = New System.Drawing.Point(143, 262)
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
        Me.btnCal.Location = New System.Drawing.Point(494, 10)
        Me.btnCal.Name = "btnCal"
        Me.btnCal.Size = New System.Drawing.Size(18, 20)
        Me.btnCal.TabIndex = 3
        Me.btnCal.Text = "V"
        '
        'Cal
        '
        Me.Cal.BackColor = System.Drawing.Color.White
        Me.Cal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal.Location = New System.Drawing.Point(312, 32)
        Me.Cal.Name = "Cal"
        Me.Cal.TabIndex = 10
        Me.Cal.Visible = False
        '
        'txtcal
        '
        Me.txtcal.AutoSize = False
        Me.txtcal.BackColor = System.Drawing.Color.White
        Me.txtcal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcal.Enabled = False
        Me.txtcal.ForeColor = System.Drawing.Color.Navy
        Me.txtcal.Location = New System.Drawing.Point(416, 10)
        Me.txtcal.MaxLength = 14
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
        Me.Label6.Location = New System.Drawing.Point(331, 11)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Invoice Date"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtGross)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtvat)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.txtnet)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Location = New System.Drawing.Point(324, 210)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(186, 70)
        Me.Panel2.TabIndex = 7
        '
        'txtGross
        '
        Me.txtGross.AutoSize = False
        Me.txtGross.BackColor = System.Drawing.Color.White
        Me.txtGross.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGross.Enabled = False
        Me.txtGross.ForeColor = System.Drawing.Color.Navy
        Me.txtGross.Location = New System.Drawing.Point(92, 44)
        Me.txtGross.Name = "txtGross"
        Me.txtGross.ReadOnly = True
        Me.txtGross.Size = New System.Drawing.Size(88, 20)
        Me.txtGross.TabIndex = 8
        Me.txtGross.Text = "0.00"
        Me.txtGross.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(2, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 18)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Gross Amount"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtvat
        '
        Me.txtvat.AutoSize = False
        Me.txtvat.BackColor = System.Drawing.Color.White
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Enabled = False
        Me.txtvat.ForeColor = System.Drawing.Color.Navy
        Me.txtvat.Location = New System.Drawing.Point(92, 22)
        Me.txtvat.Name = "txtvat"
        Me.txtvat.ReadOnly = True
        Me.txtvat.Size = New System.Drawing.Size(88, 20)
        Me.txtvat.TabIndex = 6
        Me.txtvat.Text = "0.00"
        Me.txtvat.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(92, 2)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 18)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Total VAT"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtnet
        '
        Me.txtnet.AutoSize = False
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
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(2, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 18)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Total Net"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(14, 262)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(60, 18)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "&Save"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(79, 262)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 18)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "&Close"
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowNavigation = False
        Me.DataGrid1.AllowSorting = False
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.White
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.Silver
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.ForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.GridLineColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.HeaderBackColor = System.Drawing.Color.Silver
        Me.DataGrid1.HeaderFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.LinkColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.Location = New System.Drawing.Point(14, 64)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsBackColor = System.Drawing.Color.Navy
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.ParentRowsVisible = False
        Me.DataGrid1.RowHeadersVisible = False
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.SelectionForeColor = System.Drawing.Color.Navy
        Me.DataGrid1.Size = New System.Drawing.Size(498, 144)
        Me.DataGrid1.TabIndex = 4
        '
        'txtcustname
        '
        Me.txtcustname.BackColor = System.Drawing.Color.White
        Me.txtcustname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustname.Enabled = False
        Me.txtcustname.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcustname.ForeColor = System.Drawing.Color.Navy
        Me.txtcustname.Location = New System.Drawing.Point(118, 36)
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
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(14, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Customer Name"
        '
        'CmbID
        '
        Me.CmbID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbID.Location = New System.Drawing.Point(118, 12)
        Me.CmbID.Name = "CmbID"
        Me.CmbID.Size = New System.Drawing.Size(100, 21)
        Me.CmbID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(14, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Customer ID"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(12, 216)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 18)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Payment Type *"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Items.AddRange(New Object() {"Bank / Cheque", "Credit Card"})
        Me.ComboBox1.Location = New System.Drawing.Point(111, 216)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(128, 21)
        Me.ComboBox1.TabIndex = 18
        '
        'SalesInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(536, 297)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SalesInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Sales Invoice"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose(True)
    End Sub

    Private Sub SalesInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PopulatingGrid()
        'lobal.taxRatesArray = New Double()
    End Sub

    Sub PopulatingGrid()
        ParentIDs = ""  ' Initialization
        dr = clsSales.LoadCustomerID
        While dr.Read
            CmbID.Items.Add(dr.Item("Customerid"))
        End While

        Dim dt As New DataTable("mytable")
        Dim i As Integer
        dt.Columns.Add(New DataColumn("TNo", GetType(Integer)))
        dt.Columns.Add(New DataColumn("Sales Description", GetType(String)))
        dt.Columns.Add(New DataColumn("TaxID", GetType(String)))
        dt.Columns.Add(New DataColumn("Net " & Chr(9), GetType(Double)))
        dt.Columns.Add(New DataColumn("VAT " & Chr(9), GetType(Double)))
        dt.Columns.Add(New DataColumn("Amount " & Chr(9), GetType(Double)))

        Dim tableStyle As New DataGridTableStyle

        tableStyle.MappingName = "mytable"
        myds = GlbGlobal.TaxCodeSummary()

        i = 0
        Dim TextCol As New DataGridTextBoxColumn
        TextCol.Width = 30
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 1
        TextCol = New DataGridTextBoxColumn
        TextCol.Width = 193
        TextCol.TextBox.Focus()
        TextCol.TextBox.MaxLength = 50
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 2 ' THE COMBOBOX CONTROL
        ComboTextCol.MappingName = "TaxID" 'must be from the grid table...
        ComboTextCol.HeaderText = "TaxID"
        ComboTextCol.Width = 67
        ComboTextCol.NullText = "T1"
        ComboTextCol.ColumnComboBox.DataSource = myds.Tables(0).DefaultView 'dv;
        ComboTextCol.ColumnComboBox.DisplayMember = "TaxID"
        ComboTextCol.ColumnComboBox.ValueMember = "CountId"
        ComboTextCol.ColumnComboBox.SelectedItem = ComboTextCol.ColumnComboBox.SelectedText '(ComboTextCol.ColumnComboBox.SelectedIndex)
        tableStyle.PreferredRowHeight = ComboTextCol.ColumnComboBox.Height + 2
        tableStyle.GridColumnStyles.Add(ComboTextCol)

        i = 3
        TextCol = New DataGridTextBoxColumn
        TextCol.Width = 67
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 12
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = False
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 4
        TextCol = New DataGridTextBoxColumn
        TextCol.Width = 67
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 10
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
        TextCol.ReadOnly = True
        tableStyle.GridColumnStyles.Add(TextCol)

        i = 5
        TextCol = New DataGridTextBoxColumn
        TextCol.Width = 69
        TextCol.Alignment = HorizontalAlignment.Right
        TextCol.TextBox.TextAlign = HorizontalAlignment.Right
        TextCol.TextBox.Text = "0.00"
        TextCol.Format = "0.00"
        TextCol.TextBox.MaxLength = 14
        TextCol.MappingName = dt.Columns(i).ColumnName
        TextCol.HeaderText = dt.Columns(i).ColumnName
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

        ComboBox1.SelectedIndex = 0

        If IsService = False Then
            Label7.Visible = False
            ComboBox1.Visible = False
        End If

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim TempRow As Int32, pInvno As Integer, Inc As Integer = 1, j As Integer
        Dim TempCol As Int32
        Dim Desc As String
        Dim TaxID As String
        Dim Net As Double, NetAmt As Double = 0
        Dim Vat As Double
        Dim Amt As Double = 0
        Dim CustId As String, InvDate As String, TranDate As Date
        Dim InvoiceExists As Boolean = False
        Dim PdfCollection As New Hashtable
        Dim CardInfo As New CreditCardInfo

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

            txtGross.Text = Format(VatValue + NetValue, "0.00")
            txtvat.Text = Format(VatValue, "0.00")
            txtnet.Text = Format(NetValue, "0.00")
            'MsgBox(DataGrid1.CurrentCell.RowNumber & "  " & taxRatesArray(DataGrid1.CurrentCell.RowNumber))
        End If

        If CmbID.Text = "" Then
            MsgBox("Select customer ID first.      ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        If txtcal.Text = "" Then
            MsgBox("Select invoie date first.      ", MsgBoxStyle.Information, "Infini Accounts Express")
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

        CustId = CmbID.Text
        InvDate = txtcal.Text
        TranDate = CDate(Txtdate.Text)

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        ParentNo = GlbGlobal.GetTransactionNo()
        AutoNo = GlbGlobal.GetLedgerAutoNo

        For TempRow = 0 To 4

            If DataGrid1.Item(TempRow, 3) <> 0 Then

                Desc = CStr(DataGrid1.Item(TempRow, 1))
                TaxID = Convert.ToString(myds.Tables(0).Rows(taxRatesArray(TempRow))(1))
                Net = CDbl(DataGrid1.Item(TempRow, 3))
                Vat = CDbl(DataGrid1.Item(TempRow, 4))
                Amt = CDbl(DataGrid1.Item(TempRow, 5))

                'Adding item in collection
                PdfCollection.Add(Inc, ParentNo)

                'Inserting data in ledger, transaction & outstanding table
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
        clsSales.Insertpdfinvoice(" ", TranDate, "SI", pInvno, CustId, NetAmt, DTime)
        If GlbGlobal.ExecuteCommand = False Then Exit Sub

        'Update Groupno field in Tblpdfinvoices table
        For j = 1 To PdfCollection.Count
            clsSales.updateGroupno(pInvno, CInt(PdfCollection.Item(j)))
        Next

        If (GlbGlobal.ExecuteCommand() = True) Then

            If IsService = True Then

                If ComboBox1.SelectedItem <> "Credit Card" Then
                    clsSales.insertCreditCardInfo(pInvno, TranDate, NetAmt, DTime, "@ACC" & CustId, "Invoice From Infini Accounts Express", "Invoice", "CB", CmbID.SelectedItem)
                    If ClassGlobal.ExecuteCommand = False Then Exit Sub
                Else
                    'CardInfo.MdiParent = Me.MdiParent
                    'CardInfo.Show()
                    CardInfo.Label13.Text = CStr(pInvno)
                    CardInfo.Label14.Text = CDbl(NetAmt)
                    CardInfo.Label16.Text = CDate(TranDate)
                    CardInfo.Label2.Text = ComboBox1.Text
                    CardInfo.Label17.Text = CustId
                    CardInfo.ShowDialog(Me.ParentForm)
                End If

            End If

            Dim MsgStr As String
            MsgStr = "Sales invoices have been created. Now you can print it.    " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnPrint.Enabled = True
            Button2.Enabled = False

        End If

    End Sub

    Sub TransactionAndLedgerTable(ByVal TransId As String, ByVal TaxId As String, ByVal InvDate As String, _
                               ByVal InvDetails As String, ByVal InvNet As Double, ByVal InvVat As Double, _
                               ByVal TranDate As Date, ByVal TranDateTime As String)

        'Dim ParentNo As Integer
        'ParentNo = GlbGlobal.GetTransactionNo()
        If ParentIDs.Length = 0 Then
            ParentIDs = Convert.ToString(ParentNo)
        Else
            ParentIDs = ParentIDs & "," & Convert.ToString(ParentNo)
        End If

        'Insert Data In TblTransaction
        GlbGlobal.InsertTransactionData(ParentNo, "SI", TaxId, InvDate, InvDetails, InvNet, InvVat, TranDate, TranDateTime, "N")

        'Insert Data In TblLedger
        '("SALESACCOUNT") = "10000"
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "10000", 0, InvNet, "U", TransId, "", AutoNo, TranDateTime)

        'Debtor Control Account For Ledger Transaction
        '("DEBTORSCONTROLACCOUNT") = "70100" & '("SALESACCOUNT") = "10000"
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "70100", InvNet + InvVat, 0, "T", "10000", "", AutoNo, TranDateTime)

        'Sales Tax Control Account Ledger Id Transaction
        '("SALESTAXCONTROLACCOUNT") = "71100" & '("SALESACCOUNT") = "10000"
        'AutoNo = ClsGlobal.GetLedgerAutoNo
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, "71100", 0, InvVat, "T", "10000", "", AutoNo, TranDateTime)

        'Customer Account Transaction
        'AutoNo = ClsGlobal.GetLedgerAutoNo
        AutoNo = AutoNo + 1
        GlbGlobal.InsertLedger(ParentNo, ParentNo, TransId, InvNet + InvVat, 0, "C", "10000", "*", AutoNo, TranDateTime)

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

        'MsgBox(DataGrid1.CurrentCell.RowNumber)
        maintainCC = DataGrid1.CurrentRowIndex
        DataGrid1.Item(storeRowNo, 4) = Format(DataGrid1.Item(storeRowNo, 3) * CType(Convert.ToDouble(myds.Tables(0).Rows(taxRatesArray(storeRowNo))(2)) / 100, Double), "0.00")
        DataGrid1.Item(storeRowNo, 5) = Format(DataGrid1.Item(storeRowNo, 4) + DataGrid1.Item(storeRowNo, 3), "0.00")

        numCallCurrentChaged += 1
        DataGrid1.Item(maintainCC, 4) = Format(DataGrid1.Item(maintainCC, 3) * CType(Convert.ToDouble(myds.Tables(0).Rows(taxRatesArray(maintainCC))(2)) / 100, Double), "0.00")
        DataGrid1.Item(maintainCC, 5) = Format(DataGrid1.Item(maintainCC, 4) + DataGrid1.Item(maintainCC, 3), "0.00")

        Dim i As Integer = 0
        Dim NetValue As Double = 0
        Dim VatValue As Single = 0
        Dim Amt As Double = 0

        For i = 0 To 4
            NetValue += DataGrid1.Item(i, 3)
            VatValue += DataGrid1.Item(i, 4)
            Amt += DataGrid1.Item(i, 5)
        Next

        txtGross.Text = Format(VatValue + NetValue, "0.00")
        txtvat.Text = Format(VatValue, "0.00")
        txtnet.Text = Format(NetValue, "0.00")
        storeRowNo = DataGrid1.CurrentRowIndex
        numCallCurrentChaged = 0
        numLoad = 0
    End Sub

    Private Sub Cal_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Cal.DateSelected
        txtcal.Text = Format(Cal.SelectionStart, "dd/MM/yyyy")
        Txtdate.Text = Cal.SelectionStart
        Cal.Visible = False
    End Sub

    Private Sub CmbID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbID.SelectedIndexChanged
        Dim cid As String = (CmbID.SelectedItem)
        Dr = clsSales.GetCustInfo(cid)
        Dr.Read()
        If IsDBNull(Dr("Customername")) = False Then txtcustname.Text = Dr("Customername")
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim objShowInvoiceDetail As New ShowInvoiceDetail
        objShowInvoiceDetail.MakeSalesInvoiceDetail(Convert.ToString(CmbID.SelectedItem), ParentIDs, "Sales Invoice")
        objShowInvoiceDetail.Show()
        Me.Hide()
        objShowInvoiceDetail.Activate()
    End Sub

End Class

