Imports System
Imports InfiniExpress.BLL

Public Class VendorPayment

    Inherits System.Windows.Forms.Form

#Region "Local Variables"
    Dim ClsPurchase As New ClassPurchases()
    Dim ClsGlobal As New ClassGlobal()
    Public Dr As IDataReader
    Public Ds As DataSet
    Public oldCRow As Integer, oldCCol As Integer, Trow As Integer
    Public CCol As Integer, CRow As Integer
    Private Custid As String, Tdate As String
    Dim row As Integer, i As Integer, PaymentValue As Double, DiscountValue As Double
    Dim ParentNo As Integer, InvDetails As String, InvDate As String, TranDate As Date
    Dim InvTP As String, InvValue As Double, OSDet As String, OSDet2 As String
    Public InvNo As Integer, AutoNo As Integer, MyAuto As Integer, OSMaxNo As Integer
    Dim VendID As String, TranNo As Double
    Dim ReCursiveCall As Boolean
    Dim DTime As String, dLast As Date
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
        MDI.numOfForms = 0
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Protected WithEvents Panel2 As System.Windows.Forms.Panel
    Protected WithEvents Button4 As System.Windows.Forms.Button
    Protected WithEvents Cal1 As System.Windows.Forms.MonthCalendar
    Protected WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Protected WithEvents txtRval As System.Windows.Forms.TextBox
    Protected WithEvents Label9 As System.Windows.Forms.Label
    Protected WithEvents txtbval As System.Windows.Forms.TextBox
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Button1 As System.Windows.Forms.Button
    Protected WithEvents TextBox3 As System.Windows.Forms.TextBox
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents TextBox4 As System.Windows.Forms.TextBox
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents Label8 As System.Windows.Forms.Label
    Protected WithEvents CmbVendorID As System.Windows.Forms.ComboBox
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents SCBOX As System.Windows.Forms.TextBox
    Friend WithEvents Txtdate As System.Windows.Forms.TextBox
    Friend WithEvents PCBox As System.Windows.Forms.TextBox
    Friend WithEvents PBox As System.Windows.Forms.TextBox
    Friend WithEvents DBox As System.Windows.Forms.TextBox
    Friend WithEvents CBox As System.Windows.Forms.TextBox
    Friend WithEvents Bar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VendorPayment))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Bar = New System.Windows.Forms.ProgressBar()
        Me.Txtdate = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Cal1 = New System.Windows.Forms.MonthCalendar()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.txtRval = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtbval = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.CmbVendorID = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.txt2 = New System.Windows.Forms.TextBox()
        Me.SCBOX = New System.Windows.Forms.TextBox()
        Me.PCBox = New System.Windows.Forms.TextBox()
        Me.PBox = New System.Windows.Forms.TextBox()
        Me.DBox = New System.Windows.Forms.TextBox()
        Me.CBox = New System.Windows.Forms.TextBox()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Bar, Me.Txtdate, Me.Button4, Me.btnOk, Me.Cal1, Me.DataGrid1, Me.txtRval, Me.Label9, Me.txtbval, Me.Label4, Me.Button1, Me.TextBox3, Me.Label5, Me.TextBox4, Me.Label7, Me.CmbVendorID, Me.Label8})
        Me.Panel2.Location = New System.Drawing.Point(6, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(550, 256)
        Me.Panel2.TabIndex = 27
        '
        'Bar
        '
        Me.Bar.Location = New System.Drawing.Point(194, 230)
        Me.Bar.Name = "Bar"
        Me.Bar.Size = New System.Drawing.Size(190, 16)
        Me.Bar.Step = 5
        Me.Bar.TabIndex = 34
        Me.Bar.Visible = False
        '
        'Txtdate
        '
        Me.Txtdate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Txtdate.Location = New System.Drawing.Point(226, 231)
        Me.Txtdate.Name = "Txtdate"
        Me.Txtdate.Size = New System.Drawing.Size(96, 13)
        Me.Txtdate.TabIndex = 33
        Me.Txtdate.Text = ""
        Me.Txtdate.Visible = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(78, 228)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(60, 18)
        Me.Button4.TabIndex = 32
        Me.Button4.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(6, 228)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 18)
        Me.btnOk.TabIndex = 31
        Me.btnOk.Text = "Save"
        '
        'Cal1
        '
        Me.Cal1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal1.Location = New System.Drawing.Point(170, 60)
        Me.Cal1.Name = "Cal1"
        Me.Cal1.TabIndex = 29
        Me.Cal1.Visible = False
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowSorting = False
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.White
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.Silver
        Me.DataGrid1.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(6, 64)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.DataGrid1.SelectionForeColor = System.Drawing.Color.White
        Me.DataGrid1.Size = New System.Drawing.Size(536, 158)
        Me.DataGrid1.TabIndex = 28
        '
        'txtRval
        '
        Me.txtRval.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRval.Location = New System.Drawing.Point(456, 40)
        Me.txtRval.MaxLength = 8
        Me.txtRval.Name = "txtRval"
        Me.txtRval.Size = New System.Drawing.Size(80, 20)
        Me.txtRval.TabIndex = 27
        Me.txtRval.Text = "0.00"
        Me.txtRval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label9.Location = New System.Drawing.Point(454, 23)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(90, 18)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Payment Value"
        '
        'txtbval
        '
        Me.txtbval.BackColor = System.Drawing.Color.White
        Me.txtbval.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtbval.Enabled = False
        Me.txtbval.Location = New System.Drawing.Point(368, 40)
        Me.txtbval.MaxLength = 8
        Me.txtbval.Name = "txtbval"
        Me.txtbval.Size = New System.Drawing.Size(82, 20)
        Me.txtbval.TabIndex = 25
        Me.txtbval.Text = "0.00"
        Me.txtbval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(366, 23)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 18)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Bank Value"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.ForeColor = System.Drawing.Color.Navy
        Me.Button1.Location = New System.Drawing.Point(342, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(18, 19)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "V"
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.Color.White
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(264, 40)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(80, 20)
        Me.TextBox3.TabIndex = 21
        Me.TextBox3.Text = ""
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label5.Location = New System.Drawing.Point(262, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 18)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Payment Date"
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.Color.White
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox4.Enabled = False
        Me.TextBox4.Location = New System.Drawing.Point(108, 40)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(148, 20)
        Me.TextBox4.TabIndex = 19
        Me.TextBox4.Text = ""
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 42)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 18)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Vendor Name"
        '
        'CmbVendorID
        '
        Me.CmbVendorID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbVendorID.ItemHeight = 13
        Me.CmbVendorID.Location = New System.Drawing.Point(108, 16)
        Me.CmbVendorID.Name = "CmbVendorID"
        Me.CmbVendorID.Size = New System.Drawing.Size(116, 21)
        Me.CmbVendorID.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(98, 18)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Vendor ID"
        '
        'txt1
        '
        Me.txt1.Location = New System.Drawing.Point(204, 294)
        Me.txt1.Name = "txt1"
        Me.txt1.Size = New System.Drawing.Size(54, 20)
        Me.txt1.TabIndex = 30
        Me.txt1.Text = ""
        '
        'txt2
        '
        Me.txt2.Location = New System.Drawing.Point(272, 294)
        Me.txt2.Name = "txt2"
        Me.txt2.Size = New System.Drawing.Size(54, 20)
        Me.txt2.TabIndex = 29
        Me.txt2.Text = "0.00"
        '
        'SCBOX
        '
        Me.SCBOX.Location = New System.Drawing.Point(344, 292)
        Me.SCBOX.Name = "SCBOX"
        Me.SCBOX.Size = New System.Drawing.Size(54, 20)
        Me.SCBOX.TabIndex = 28
        Me.SCBOX.Text = "0"
        '
        'PCBox
        '
        Me.PCBox.Location = New System.Drawing.Point(440, 320)
        Me.PCBox.MaxLength = 20
        Me.PCBox.Name = "PCBox"
        Me.PCBox.Size = New System.Drawing.Size(54, 20)
        Me.PCBox.TabIndex = 31
        Me.PCBox.TabStop = False
        Me.PCBox.Text = "0"
        Me.PCBox.Visible = False
        '
        'PBox
        '
        Me.PBox.Location = New System.Drawing.Point(504, 320)
        Me.PBox.MaxLength = 20
        Me.PBox.Name = "PBox"
        Me.PBox.Size = New System.Drawing.Size(54, 20)
        Me.PBox.TabIndex = 32
        Me.PBox.TabStop = False
        Me.PBox.Text = "0"
        Me.PBox.Visible = False
        '
        'DBox
        '
        Me.DBox.Location = New System.Drawing.Point(440, 344)
        Me.DBox.MaxLength = 20
        Me.DBox.Name = "DBox"
        Me.DBox.Size = New System.Drawing.Size(54, 20)
        Me.DBox.TabIndex = 33
        Me.DBox.TabStop = False
        Me.DBox.Text = "0"
        Me.DBox.Visible = False
        '
        'CBox
        '
        Me.CBox.Location = New System.Drawing.Point(504, 344)
        Me.CBox.MaxLength = 20
        Me.CBox.Name = "CBox"
        Me.CBox.Size = New System.Drawing.Size(54, 20)
        Me.CBox.TabIndex = 34
        Me.CBox.TabStop = False
        Me.CBox.Text = "0"
        Me.CBox.Visible = False
        '
        'VendorPayment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(562, 267)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CBox, Me.DBox, Me.PBox, Me.PCBox, Me.txt1, Me.txt2, Me.SCBOX, Me.Panel2})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VendorPayment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Vendor Payment"
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub VendorPayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dr = ClsPurchase.LoadVendorID
        While Dr.Read
            CmbVendorID.Items.Add(Dr.Item("Vendorid"))
        End While
        Dr.Close()
        populating("")
        ReCursiveCall = False
    End Sub

    Private Sub CmbVendorID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbVendorID.SelectedIndexChanged
        VendID = (CmbVendorID.SelectedItem)
        Dr = ClsPurchase.GetVendorDetails(VendID)
        Dr.Read()
        If IsDBNull(Dr("Vendorname")) = False Then
            TextBox4.Text = Dr("Vendorname")
        End If
        Dr.Close()

        'Get Bank Balance Against Particuler User
        txtbval.Text = Format(ClsGlobal.GetBankBalance("70200"), "0.00")
        txt2.Text = CDbl(txtbval.Text)
        populating(VendID)
    End Sub

    Public Function GetTable(ByVal tempCId As String) As DataTable

        Dim Tinvtp As String, Tinvdate As String, Tinvdetails As String
        Dim Tmyauto As Double, Texpr2 As Double, Texpr1 As Double
        Dim Ttaxid As String, Tpaystatus As String, Ttransid As String
        Dim R3 As String, Paystatus As String, Tid As Integer, TAmt As Double
        Dim Tparentid As Integer
        Dim Rd As IDataReader

        Dim DT As New DataTable   'Creating DataTable
        Dim DCTNO As New DataColumn("Parentid", GetType(Integer))
        Dim DCType As New DataColumn("Invtp", GetType(String))
        Dim DCTaxID As New DataColumn("TaxID", GetType(String))
        Dim DCDate As New DataColumn("InvDate", GetType(String))
        Dim DCTDetail As New DataColumn("InvDetails", GetType(String))
        Dim DCmyauto As New DataColumn("myautono", GetType(Double))
        Dim DCPaystatus As New DataColumn("Paystatus", GetType(String))
        Dim DCExpr1 As New DataColumn("Expr1", GetType(Double))
        Dim DCExpr2 As New DataColumn("Expr2", GetType(Double))
        Dim DCNull1 As New DataColumn("NullCol1", GetType(Double))
        Dim DCNull2 As New DataColumn("NullCol2", GetType(Double))
        Dim DCAmt As New DataColumn("Amount", GetType(Double))


        'Add The Column to the Datatable's Columns Collection

        DT.Columns.Add(DCNull1)
        DT.Columns.Add(DCNull2)
        DT.Columns.Add(DCTNO)
        DT.Columns.Add(DCType)
        DT.Columns.Add(DCTaxID)
        DT.Columns.Add(DCDate)
        DT.Columns.Add(DCTDetail)
        DT.Columns.Add(DCmyauto)
        DT.Columns.Add(DCPaystatus)
        DT.Columns.Add(DCExpr1)
        DT.Columns.Add(DCExpr2)
        DT.Columns.Add(DCAmt)

        'Creating DataRow
        Dim DR As DataRow
        Rd = ClsPurchase.LoadVendorPayments(tempCId)

        While Rd.Read
            Dim amt As Double = Format(CDbl(Rd("Amount")), "0.00")
            If amt <> 0.0 Then
                DR = DT.NewRow()

                DR("NullCol1") = "0.00" 'Rd("Expr1")
                DR("NullCol2") = "0.00" 'Rd("Expr2")
                DR("parentid") = Rd("parentid")
                DR("Invtp") = Rd("Invtp")
                DR("Taxid") = Rd("Taxid")
                DR("Invdate") = Rd("Invdate")
                DR("Invdetails") = Rd("Invdetails")
                DR("myautono") = Rd("myautono")
                DR("paystatus") = Rd("paystatus")
                DR("Expr1") = Rd("Expr1")
                DR("Expr2") = Rd("Expr2")
                DR("Amount") = Rd("Amount")

                DT.Rows.Add(DR)

            End If
        End While
        Rd.Close()
        Return DT

    End Function


    Sub populating(ByVal tempCid As String)
        Ds = New DataSet
        Dim Dtab As DataTable
        Dtab = GetTable(tempCid)
        Ds.Merge(Dtab)
        'Ds = ClsPurchase.LoadVendorPayments(tempCid)
        DataGrid1.TableStyles.Clear()

        'Formating The DataGrid Bt Adding a DataGrid Table Style
        Dim DgridTstyle As New DataGridTableStyle
        Dim Type As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Invdate As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Details As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Value As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Payment As New DataGridTextBoxColumn
        Dim Discount As New DataGridTextBoxColumn
        Dim MyAuto As New DataGridTextBoxColumn
        Dim ParentID As New DataGridTextBoxColumn

        With Type
            .HeaderText = "InvTP"
            .MappingName = "invtp"
            .Width = 40
            .ReadOnly = True
        End With

        With Invdate
            .HeaderText = "Date"
            .MappingName = "invdate"
            .Width = 60
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Details
            .HeaderText = "Details"
            .MappingName = "invdetails"
            .Width = 205
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Value
            .HeaderText = "Value"
            .MappingName = "Amount"
            .Width = 50
            .ReadOnly = True
            .TextBox.TextAlign = HorizontalAlignment.Right
            .TextBox.ReadOnly = True
            .Format = "0.00"
        End With

        With Payment
            .HeaderText = "Payment " & Chr(9)
            .Width = 80
            .MappingName = "NullCol1"
            .TextBox.MaxLength = 12
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .Format = "0.00"
        End With

        With Discount
            .HeaderText = "Discount" & " " & Chr(9)
            .Width = 80
            .MappingName = "NullCol2"
            .TextBox.MaxLength = 10
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .Format = "0.00"
        End With

        With MyAuto
            .HeaderText = ""
            .Width = 0
            .MappingName = "MyAutoNo"
        End With

        With ParentID
            .HeaderText = ""
            .Width = 0
            .MappingName = "Parentid"
        End With

        'Add The GridColumnStyle to Datagrid Column Style
        With DgridTstyle.GridColumnStyles
            .Add(Type)
            .Add(Invdate)
            .Add(Details)
            .Add(Value)
            .Add(Payment)
            .Add(Discount)
            .Add(MyAuto)
            .Add(ParentID)
        End With

        'Bind The Datagrid to Dataset
        Trow = Ds.Tables(0).Rows.Count
        If Ds.Tables(0).Rows.Count > 0 Then
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
            DataGrid1.DataSource = Ds.Tables(0)

            'Stoping Extra Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False

            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = Ds.Tables(0).TableName '("TblTransaction")
            DataGrid1.TableStyles.Add(DgridTstyle)
            btnOk.Enabled = False

        End If

        oldCRow = 0
        oldCCol = 0

    End Sub

    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged

        If ReCursiveCall = True Then
            Exit Sub
        End If
        'Dim Trows As Integer
        Dim CustId As String, PaymentDate As Date, r As Integer, TDate As String
        Dim InvNo As Double, InvTp As String, InvDate As Date, InvDetails As String
        Dim InvValue As Double, InvPayment As Double, InvDiscount As Double
        Dim i As Integer, DisValue As Double, j As Integer
        Dim rt As Integer
        Dim RValue As Double = CDbl("0.00")

        CCol = DataGrid1.CurrentCell.ColumnNumber
        CRow = DataGrid1.CurrentCell.RowNumber

        If oldCCol = 4 Or oldCCol = 5 Then

            If IsNumeric(DataGrid1.Item(oldCRow, oldCCol)) = False Then
                MsgBox("Enter numeric value in payment.   ", MsgBoxStyle.Information, "Infini Accounts Express")
                DataGrid1.Item(oldCRow, oldCCol) = "0.00"
                Exit Sub
            End If

            ReCursiveCall = True

            If DataGrid1.Item(oldCRow, oldCCol) = 0 Or DataGrid1.Item(oldCRow, oldCCol) = 0.0 Then
                DataGrid1.Item(oldCRow, oldCCol) = "0.00"
            End If
            ReCursiveCall = False
            'Trows = DataGrid1.VisibleRowCount()

            For i = 0 To Trow - 1

                InvTp = Trim(DataGrid1.Item(i, 0)) ' InvTP
                If InvTp = "" Then Exit For
                TDate = DataGrid1.Item(i, 1) ' InvDate
                InvDetails = DataGrid1.Item(i, 2) ' InvDetails
                InvValue = Format((DataGrid1.Item(i, 3)), "0.00") ' InvValue

                InvPayment = CDbl(DataGrid1.Item(i, 4)) ' Receipt
                If IsNumeric(InvPayment) = False Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    Exit For
                End If
                InvDiscount = CType(DataGrid1.Item(i, 5), Double) ' Discount
                If IsNumeric(InvDiscount) = False Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    Exit For
                End If

                If InvTp <> "PI" Then
                    If InvDiscount > 0 Then
                        'MsgBox("Negative values are not allowed.     ", MsgBoxStyle.Information, "Infini Accounts Express")
                        DataGrid1.Item(i, 4) = "0.00"
                        DataGrid1.Item(i, 5) = "0.00"
                        Exit For
                    End If
                End If

                DisValue = CDbl(InvPayment) + CDbl(InvDiscount)
                If DisValue > InvValue Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    MessageBox.Show("Payment amount is greater then amount of settlement.      ", "Infini Accounts Express", MessageBoxButtons.OK)
                    txtbval.Text = "0.00"
                    txtRval.Text = "0.00"
                    Exit For
                End If

                'If InvPayment = 0.0 And InvDiscount > 0 And InvDiscount <= InvValue Then
                '    InvPayment = InvValue - InvDiscount
                '    DataGrid1.Item(i, 4) = Format(InvPayment, "0.00")
                'End If

                If InvTp = "PI" Then
                    RValue = CDbl(RValue) + CDbl(InvPayment)
                    RValue = Format(RValue, "0.00")
                Else
                    RValue = CDbl(RValue) - CDbl(InvPayment)
                    RValue = Format(RValue, "0.00")
                End If
            Next

            txtbval.Text = Format(CDbl(txt2.Text) + CDbl(RValue), "0.00")
            txtRval.Text = Format(RValue, "0.00")
        End If
        oldCRow = CRow
        oldCCol = CCol
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose(True)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Cal1.Visible = False Then
            Cal1.Visible = True
        Else
            Cal1.Visible = False
        End If
    End Sub

    Public Sub IncreaseProgressBar(ByVal Value As Integer)
        ' Increment the value of the ProgressBar a value of one each time.
        Bar.Increment(Value)

        ' Determine if we have completed by comparing the value of the Value property to the Maximum value.
        If Bar.Value = Bar.Maximum Then
            ' Stop the timer.
            Exit Sub
        End If
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim A As Object, B As Object
        DataGrid1_CurrentCellChanged(A, B)

        Dim TFlag As Boolean, OSAutoNo As Double
        Dim PPValue As Double, PAValue As Double, Amount As Double
        Dim VendorID As String = CmbVendorID.SelectedItem
        Dim PaymentValue As Double

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Select payment date first.   ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        TFlag = False
        InvDetails = "Purchase Payment"
        InvDate = TextBox3.Text
        TranDate = CDate(Txtdate.Text)

        'Payment Amount
        Amount = txtRval.Text

        If Amount < 0 Then
            MsgBox("Unsettled cheque balance for this account.      ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtbval.Text = "0.00"
            txtRval.Text = "0.00"
            Exit Sub
        End If

        row = DataGrid1.VisibleRowCount
        For i = 0 To row - 1
            PaymentValue = DataGrid1.Item(i, 4)
            DiscountValue = DataGrid1.Item(i, 5)
            If PaymentValue > 0 Or DiscountValue > 0 Then
                PaymentValue = 0
                DiscountValue = 0
                TFlag = True
                Exit For
            End If
        Next i

        If Amount > 0 Then TFlag = True
        If TFlag = False Then Exit Sub

        'Calculate PP Payment Value
        PPValue = Format(CalculatePP(), "0.00")

        If PPValue < 0 Then
            MsgBox("Unsettled cheque balance for this account.     ", MsgBoxStyle.Exclamation, "Infini Accounts Express")
            Exit Sub
        End If

        PAValue = Format(Amount - PPValue, "0.00")
        If PAValue > 0 Then
            MsgBox("Unsettled cheque balance of " & Format(PAValue, "0.00") & " found for this account.      ", MsgBoxStyle.Exclamation, "Infini Accounts Express")
        End If

        DTime = dLast.Now 'Transaction DateTime

        'Get New Transaction No
        ParentNo = ClsGlobal.GetTransactionNo()
        'Get New Ledger Auto No
        AutoNo = ClsGlobal.GetLedgerAutoNo
        'Get New Outstanding Max. No
        OSMaxNo = ClsGlobal.GetOutStandingMaxOSNo()

        'Transaction For PP
        If PPValue > 0 Then

            ''Get New Transaction No
            'ParentNo = ClsGlobal.GetTransactionNo()

            'Payment PP Transaction
            ClsGlobal.InsertTransactionData(ParentNo, "PP", "T9", InvDate, InvDetails, PPValue, 0, TranDate, DTime)

            'Debtors Control Account Transaction
            '("CREDITORSCONTROLACCOUNT") = "71000"
            '("DEFAULTBANK") = "70200"
            ClsGlobal.InsertLedger(ParentNo, ParentNo, "71000", PPValue, 0, "T", "70200", "f", AutoNo, DTime)

        End If

        row = DataGrid1.VisibleRowCount

        For i = 0 To row - 1

            With DataGrid1

                InvTP = Trim(.Item(i, 0)) ' InvTP
                If InvTP = "" Then Exit For
                InvValue = Format(CDbl(.Item(i, 3)), "0.00") ' InvValue
                PaymentValue = Format(CDbl(.Item(i, 4)), "0.00") ' PaymentValue
                DiscountValue = Format(CDbl(.Item(i, 5)), "0.00") ' Discount
                OSAutoNo = CInt(.Item(i, 6)) ' MyAutoNo

            End With

            If PaymentValue > 0 Then
                If InvValue = (PaymentValue + DiscountValue) Then 'Full PaymentValue
                    ClsGlobal.UpdateLedgerPayStatus(OSAutoNo, "f", DTime)
                ElseIf PaymentValue > 0 Then 'Partial PaymentValue
                    ClsGlobal.UpdateLedgerPayStatus(OSAutoNo, "p", DTime)
                End If
            End If
            PaymentValue = 0

        Next i

        'Calculation Of Purchase Invoice "PI"
        CalculationPI()

        If PAValue > 0 Then 'Transaction For PA

            'Get New Transaction No. 
            InvNo = ParentNo + 1

            'Purchases On Account PA Transaction
            InvDetails = "Payment On Account"
            ClsGlobal.InsertTransactionData(InvNo, "PA", "T9", InvDate, "Payment On Account", PAValue, 0, TranDate, DTime)

            'Bank Credit Transaction
            'Application("DEFAULTBANK") = "70200"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, "70200", 0, PAValue, "B", VendID, "", AutoNo, DTime)

            'Vendor Debit Transaction
            'Application("DEFAULTBANK") = "70200"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, VendID, PAValue, 0, "S", "70200", "*", AutoNo, DTime)

            'Debtors Control Account Debit Transaction
            'Application("DEFAULTBANK") = "70200"
            '"CREDITORSCONTROLACCOUNT" = "71000"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, "71000", PAValue, 0, "T", "70200", "", AutoNo, DTime)

        End If
        If ClsGlobal.ExecuteCommand() = True Then

            Dim MsgStr As String
            MsgStr = "Payment has been created.   " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose(True)

        End If


    End Sub

    Private Function CalculatePP() As Double

        Dim PValue As Double
        Dim PIValue As Double, PCValue As Double, InvTP As String
        Dim row As Integer, r As Integer

        row = DataGrid1.VisibleRowCount
        PIValue = 0
        PCValue = 0

        For r = 0 To row - 1

            InvTP = Trim(DataGrid1.Item(r, 0)) ' InvTP
            PValue = DataGrid1.Item(r, 4)
            If InvTP = "" Then Exit For

            If InvTP = "PI" Then
                PIValue = PIValue + CDbl(PValue)
            Else
                PCValue = PCValue + CDbl(PValue)
            End If

        Next r

        Return (PIValue - PCValue)

    End Function

    Public Sub CalculationPI()

        Dim Balance As Double
        Dim K As Integer, OSNo As Double
        Dim OSAmount As Double, OSTP As String
        Dim KRow As Integer
        Dim dt As DataTable = CType(DataGrid1.DataSource, DataTable)
        Dim dr As DataRow
        Dim tr As DataRow

        For Each dr In dt.Rows

            InvTP = dr(3)  'InvTP
            If InvTP = "" Then Exit For
            InvNo = dr(2) 'InvNo.
            Tdate = dr(5) 'InvDate
            InvDetails = dr(6) 'InvDetails
            PBox.Text = dr(0) 'Payment
            PaymentValue = PBox.Text

            If InvTP = "PI" Then

                While PaymentValue > 0

                    Balance = 0
                    For Each tr In dt.Rows

                        OSTP = tr(3) 'InvTP
                        If OSTP = "" Then Exit For
                        OSNo = tr(2) 'InvNo.
                        CBox.Text = tr(0) 'Payment
                        Balance = CDbl(CBox.Text)
                        If Balance > 0 Then
                            If OSTP = "PC" Or OSTP = "PA" Then
                                CBox.Text = "0.00"
                                Exit For
                            Else
                                Balance = 0
                            End If
                        End If
                    Next

                    'PC Transactions
                    If Balance > 0 Then
                        OSAmount = Balance
                        If OSAmount > PaymentValue Then
                            OSDet = Format(PaymentValue, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(PaymentValue, "0.00") & " To " & InvTP & " " & InvNo
                            OSAmount = PaymentValue
                            PCBox.Text = Format(OSAmount - PaymentValue, "0.00")
                            tr(0) = Format(Balance - PaymentValue, "0.00")
                            PaymentValue = 0
                            PBox.Text = "0.00"
                        ElseIf OSAmount = PaymentValue Then
                            OSDet = Format(PaymentValue, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(PaymentValue, "0.00") & " To " & InvTP & " " & InvNo
                            PaymentValue = 0
                            tr(0) = 0
                            PBox.Text = "0.00"
                            PCBox.Text = "0.00"
                        ElseIf OSAmount < PaymentValue Then
                            OSDet = Format(OSAmount, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(OSAmount, "0.00") & " To " & InvTP & " " & InvNo
                            PaymentValue = PaymentValue - OSAmount
                            PBox.Text = Format(PaymentValue, "0.00")
                            PCBox.Text = PaymentValue
                            tr(0) = 0
                        End If

                        'OutStanding Credit Vendor Transaction
                        MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "S")
                        ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, VendID, 0, OSAmount, "S", OSDet, "", InvDate, OSNo, 0, "F", DTime, OSMaxNo)

                        'OutStanding Debit Vendor Transaction
                        OSMaxNo = OSMaxNo + 1
                        MyAuto = ClsGlobal.GetExistingLedgerAutoNo(OSNo, "S")
                        ClsGlobal.InsertOutStanding(MyAuto, OSNo, OSNo, VendID, OSAmount, 0, "S", OSDet2, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                        If OSTP = "PA" Then
                            'OutStanding Debit Bank Transaction
                            'Application("DEFAULTBANK")="70200"
                            'Get ledger MyAuto
                            OSMaxNo = OSMaxNo + 1
                            ClsGlobal.InsertOutStanding(AutoNo, OSNo, OSNo, "70200", 0, OSAmount, "B", OSDet2, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)
                        End If
                        Balance = 0
                        PBox.Text = PCBox.Text
                        PaymentValue = CDbl(PCBox.Text)
                        OSMaxNo = OSMaxNo + 1

                    Else

                        'PI Transactions
                        If PaymentValue > 0 Then

                            'Bank Transaction
                            'Application("DEFAULTBANK")="70200"
                            AutoNo = AutoNo + 1
                            ClsGlobal.InsertLedger(ParentNo, ParentNo, "70200", 0, PaymentValue, "B", VendID, "", AutoNo, DTime)

                            'OutStanding Bank Transaction
                            'Application("DEFAULTBANK") = "70200"
                            OSDet = Format(PaymentValue, "0.00") & " To PI " & InvNo

                            ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, "70200", 0, PaymentValue, "B", OSDet, VendID, InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                            'Vendor Transaction
                            'Application("DEFAULTBANK")="70200"
                            AutoNo = AutoNo + 1
                            ClsGlobal.InsertLedger(ParentNo, ParentNo, VendID, PaymentValue, 0, "S", "70200", "f", AutoNo, DTime)

                            'OutStanding Vendor Transaction
                            'Application("DEFAULTBANK")="70200"
                            OSMaxNo = OSMaxNo + 1
                            ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, VendID, PaymentValue, 0, "S", OSDet, "70200", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                            'OutStanding Vendor Transaction
                            OSDet = Format(PaymentValue, "0.00") & " From PP " & ParentNo

                            OSMaxNo = OSMaxNo + 1
                            MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "S")
                            ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, VendID, 0, PaymentValue, "S", OSDet, "", InvDate, InvNo, 0, "F", DTime, OSMaxNo)
                        End If
                        PBox.Text = "0.00"
                        PaymentValue = 0
                        OSMaxNo = OSMaxNo + 1
                    End If

                End While

            Else
                'PI Transactions
                PaymentValue = CDbl(SCBOX.Text)
                If PaymentValue > 0 Then
                    'Bank Transaction
                    'Application("DEFAULTBANK")="70200"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(ParentNo, ParentNo, "70200", 0, PaymentValue, "B", VendID, "", AutoNo, DTime)

                    'OutStanding Bank Transaction
                    'Application("DEFAULTBANK") = "70200"
                    OSDet = Format(PaymentValue, "0.00") & " To PI " & InvNo

                    ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, "70200", 0, PaymentValue, "B", OSDet, VendID, InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'Vendor Transaction
                    'Application("DEFAULTBANK")="70200"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(ParentNo, ParentNo, VendID, PaymentValue, 0, "S", "70200", "f", AutoNo, DTime)

                    'OutStanding Vendor Transaction
                    'Application("DEFAULTBANK")="70200"
                    OSMaxNo = OSMaxNo + 1
                    ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, VendID, PaymentValue, 0, "S", OSDet, "70200", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'OutStanding Vendor Transaction
                    OSDet = Format(PaymentValue, "0.00") & " From PP " & ParentNo

                    'Get ledger MyAuto
                    OSMaxNo = OSMaxNo + 1
                    MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "S")
                    ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, VendID, 0, PaymentValue, "S", OSDet, "", InvDate, InvNo, 0, "F", DTime, OSMaxNo)
                End If
                PaymentValue = 0
                PBox.Text = "0.00"
                OSMaxNo = OSMaxNo + 1
            End If

        Next

        'Get New Transaction No.
        TranNo = ParentNo + 1
        Dim DCheck As Boolean = False

        For Each dr In dt.Rows

            InvTP = dr(3) 'InvTP
            If InvTP = "" Then Exit For
            InvNo = dr(2) 'InvNo.
            Tdate = dr(5) 'InvDate
            InvDetails = dr(6) 'InvDetails
            DBox.Text = dr(1) 'Discount
            DiscountValue = DBox.Text

            If InvTP = "PI" Then

                ' Discount Transactions
                If DiscountValue > 0 Then 'Transaction For PD

                    DCheck = True

                    'Discount SD Transaction
                    InvDetails = "Purchase Discount"
                    ClsGlobal.InsertTransactionData(TranNo, "PD", "T9", InvDate, InvDetails, DiscountValue, 0, TranDate, DTime)

                    'Sales Discount Control Account Transaction
                    '("PURCHASEDISCOUNT") = "20010"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(TranNo, TranNo, "20010", 0, DiscountValue, "T", VendID, "", AutoNo, DTime)

                    'Debtors Control Account Transaction
                    'Application("CREDITORSCONTROLACCOUNT") = "71000"
                    'Application("PURCHASEDISCOUNT") = "20010"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(TranNo, TranNo, "71000", DiscountValue, 0, "T", "20010", "", AutoNo, DTime)

                    'Customer Account Transaction
                    'Application("PURCHASEDISCOUNT") = "20010"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(TranNo, TranNo, VendID, DiscountValue, 0, "S", "20010", "f", AutoNo, DTime)

                    'OutStanding Credit Customer Transaction
                    OSDet = Format(DiscountValue, "0.00") & " Payed To PI " & InvNo

                    'Get ledger MyAuto
                    ClsGlobal.InsertOutStanding(AutoNo, TranNo, TranNo, VendID, DiscountValue, 0, "S", OSDet, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'OutStanding Debit Customer Transaction
                    OSDet = Format(DiscountValue, "0.00") & " From PD " & TranNo

                    'Get ledger MyAuto
                    OSMaxNo = OSMaxNo + 1
                    MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "S")
                    ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, VendID, 0, DiscountValue, "S", OSDet, "", InvDate, TranNo, 0, "F", DTime, OSMaxNo)
                    DBox.Text = "0.00"
                    TranNo = TranNo + 1
                    OSMaxNo = OSMaxNo + 1
                End If
            End If
        Next
        If DCheck = True Then
            ParentNo = TranNo - 1
        End If

    End Sub

    Private Sub txtRVal_TextLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRval.Leave ' .TextChanged

        If IsNumeric(txtRval.Text) = False Then
            txtRval.Text = "0.00"
        End If
        txtRval.Text = Format(Val(txtRval.Text), "0.00")

    End Sub

    Private Sub Cal1_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Cal1.DateSelected
        TextBox3.Text = Format(Cal1.SelectionStart, "dd/MM/yyyy")
        Txtdate.Text = Cal1.SelectionStart
        Cal1.Visible = False
    End Sub

End Class
