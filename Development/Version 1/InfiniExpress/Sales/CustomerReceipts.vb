Imports System
Imports InfiniExpress.BLL

Public Class CustomerReceipts

    Inherits System.Windows.Forms.Form

#Region "Local Variables"
    Public ClsSales As New ClassSales()
    Dim ClsGlobal As New ClassGlobal()
    Public Dr As IDataReader
    Public Ds As DataSet
    Public oldCRow As Integer, oldCCol As Integer
    Public CCol As Integer, CRow As Integer
    Private Custid As String, Tdate As String
    Dim row As Integer, i As Integer, ReceiptValue As Double, DiscountValue As Double
    Dim ParentNo As Integer, InvDetails As String, InvDate As String, TranDate As Date
    Dim invtp As String, InvValue As Double, OSDet As String, OSDet2 As String
    Public InvNo As Integer, AutoNo As Integer, MyAuto As Integer, OSMaxNo As Integer
    Dim recursiveCall As Boolean, GRow As Integer
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
    Protected WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents btnCal As System.Windows.Forms.Button
    Protected WithEvents Cal As System.Windows.Forms.MonthCalendar
    Protected WithEvents txtcal As System.Windows.Forms.TextBox
    Protected WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents txtcustname As System.Windows.Forms.TextBox
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents cmbcid As System.Windows.Forms.ComboBox
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents SCBOX As System.Windows.Forms.TextBox
    Friend WithEvents txt2 As System.Windows.Forms.TextBox
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents CBOX As System.Windows.Forms.TextBox
    Friend WithEvents Txtdate As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CustomerReceipts))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Txtdate = New System.Windows.Forms.TextBox()
        Me.Cal1 = New System.Windows.Forms.MonthCalendar()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.txtRval = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtbval = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmbcid = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnCal = New System.Windows.Forms.Button()
        Me.Cal = New System.Windows.Forms.MonthCalendar()
        Me.txtcal = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtcustname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SCBOX = New System.Windows.Forms.TextBox()
        Me.txt2 = New System.Windows.Forms.TextBox()
        Me.txt1 = New System.Windows.Forms.TextBox()
        Me.CBOX = New System.Windows.Forms.TextBox()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Txtdate, Me.Cal1, Me.Button4, Me.btnOk, Me.txtRval, Me.Label9, Me.txtbval, Me.Label4, Me.Button1, Me.TextBox3, Me.Label5, Me.TextBox4, Me.Label7, Me.cmbcid, Me.Label8, Me.DataGrid1})
        Me.Panel2.Location = New System.Drawing.Point(-2, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(550, 256)
        Me.Panel2.TabIndex = 26
        '
        'Txtdate
        '
        Me.Txtdate.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Txtdate.Location = New System.Drawing.Point(226, 230)
        Me.Txtdate.Name = "Txtdate"
        Me.Txtdate.Size = New System.Drawing.Size(96, 13)
        Me.Txtdate.TabIndex = 34
        Me.Txtdate.Text = ""
        Me.Txtdate.Visible = False
        '
        'Cal1
        '
        Me.Cal1.BackColor = System.Drawing.Color.White
        Me.Cal1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal1.ForeColor = System.Drawing.Color.Black
        Me.Cal1.Location = New System.Drawing.Point(157, 58)
        Me.Cal1.Name = "Cal1"
        Me.Cal1.TabIndex = 29
        Me.Cal1.TitleBackColor = System.Drawing.Color.Gray
        Me.Cal1.TrailingForeColor = System.Drawing.SystemColors.Highlight
        Me.Cal1.Visible = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(76, 228)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(60, 18)
        Me.Button4.TabIndex = 32
        Me.Button4.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.Enabled = False
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(6, 228)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(60, 18)
        Me.btnOk.TabIndex = 31
        Me.btnOk.Text = "Save"
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
        Me.Label9.Location = New System.Drawing.Point(454, 20)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(90, 18)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Receipts Value"
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
        Me.Label4.Location = New System.Drawing.Point(366, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 18)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Bank Value"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
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
        Me.TextBox3.ForeColor = System.Drawing.Color.Navy
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
        Me.Label5.Location = New System.Drawing.Point(262, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 18)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Receipts Date"
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
        Me.Label7.Text = "Customer Name"
        '
        'cmbcid
        '
        Me.cmbcid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbcid.ItemHeight = 13
        Me.cmbcid.Location = New System.Drawing.Point(108, 16)
        Me.cmbcid.Name = "cmbcid"
        Me.cmbcid.Size = New System.Drawing.Size(108, 21)
        Me.cmbcid.TabIndex = 17
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
        Me.Label8.Text = "Customer ID"
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowNavigation = False
        Me.DataGrid1.AllowSorting = False
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.White
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.Silver
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.ForeColor = System.Drawing.Color.Black
        Me.DataGrid1.HeaderFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.Black
        Me.DataGrid1.Location = New System.Drawing.Point(6, 64)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.Black
        Me.DataGrid1.ParentRowsVisible = False
        Me.DataGrid1.RowHeadersVisible = False
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.DataGrid1.Size = New System.Drawing.Size(536, 158)
        Me.DataGrid1.TabIndex = 28
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(342, 40)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(82, 20)
        Me.TextBox1.TabIndex = 25
        Me.TextBox1.Text = ""
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Navy
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(342, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 18)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Invoice Date"
        '
        'btnCal
        '
        Me.btnCal.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.btnCal.Location = New System.Drawing.Point(556, 18)
        Me.btnCal.Name = "btnCal"
        Me.btnCal.Size = New System.Drawing.Size(18, 20)
        Me.btnCal.TabIndex = 23
        '
        'Cal
        '
        Me.Cal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cal.Location = New System.Drawing.Point(84, 94)
        Me.Cal.Name = "Cal"
        Me.Cal.TabIndex = 22
        '
        'txtcal
        '
        Me.txtcal.Location = New System.Drawing.Point(502, 142)
        Me.txtcal.Name = "txtcal"
        Me.txtcal.Size = New System.Drawing.Size(88, 20)
        Me.txtcal.TabIndex = 21
        Me.txtcal.Text = ""
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Navy
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(508, 86)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 18)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Invoice Date"
        '
        'txtcustname
        '
        Me.txtcustname.Enabled = False
        Me.txtcustname.Location = New System.Drawing.Point(114, 40)
        Me.txtcustname.Name = "txtcustname"
        Me.txtcustname.Size = New System.Drawing.Size(222, 20)
        Me.txtcustname.TabIndex = 19
        Me.txtcustname.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Navy
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(12, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 18)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Customer Name"
        '
        'ComboBox1
        '
        Me.ComboBox1.ItemHeight = 13
        Me.ComboBox1.Location = New System.Drawing.Point(114, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(100, 21)
        Me.ComboBox1.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Navy
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 18)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Customer ID"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.TextBox1, Me.Label3, Me.btnCal, Me.Cal, Me.txtcal, Me.Label6, Me.txtcustname, Me.Label2, Me.ComboBox1, Me.Label1})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(550, 256)
        Me.Panel1.TabIndex = 0
        '
        'SCBOX
        '
        Me.SCBOX.Location = New System.Drawing.Point(398, 282)
        Me.SCBOX.Name = "SCBOX"
        Me.SCBOX.Size = New System.Drawing.Size(54, 20)
        Me.SCBOX.TabIndex = 1
        Me.SCBOX.Text = "0"
        '
        'txt2
        '
        Me.txt2.Location = New System.Drawing.Point(326, 284)
        Me.txt2.Name = "txt2"
        Me.txt2.Size = New System.Drawing.Size(54, 20)
        Me.txt2.TabIndex = 2
        Me.txt2.Text = "0.00"
        '
        'txt1
        '
        Me.txt1.Location = New System.Drawing.Point(258, 284)
        Me.txt1.Name = "txt1"
        Me.txt1.Size = New System.Drawing.Size(54, 20)
        Me.txt1.TabIndex = 3
        Me.txt1.Text = ""
        '
        'CBOX
        '
        Me.CBOX.Location = New System.Drawing.Point(332, 312)
        Me.CBOX.Name = "CBOX"
        Me.CBOX.Size = New System.Drawing.Size(76, 20)
        Me.CBOX.TabIndex = 4
        Me.CBOX.Text = ""
        '
        'CustomerReceipts
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(558, 263)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CBOX, Me.txt1, Me.txt2, Me.SCBOX, Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustomerReceipts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Receipts"
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Cal1.Visible = False Then
            Cal1.Visible = True
        Else
            Cal1.Visible = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose(True)
    End Sub

    Private Sub CustomerReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dr = ClsSales.LoadCustomerID
        While Dr.Read
            cmbcid.Items.Add(Dr.Item("Customerid"))
        End While
        populating("")
        recursiveCall = False
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

    Private Sub CmbcID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbcid.SelectedIndexChanged

        Custid = (cmbcid.SelectedItem)
        Dr = ClsSales.GetCustInfo(Custid)
        Dr.Read()
        If IsDBNull(Dr("Customername")) = False Then
            TextBox4.Text = Dr("Customername")
        End If
        Dr.Close()

        'Get Bank Balance Against Particuler User
        txtbval.Text = Format(ClsGlobal.GetBankBalance("70200"), "0.00")
        txt2.Text = txtbval.Text

        populating(Custid)
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
        Rd = ClsSales.LoadCustomerReceipts(tempCid)

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
        'Ds = ClsSales.LoadCustomerReceipts(tempCid)
        DataGrid1.TableStyles.Clear()

        'Formating The DataGrid Bt Adding a DataGrid Table Style
        Dim DgridTstyle As New DataGridTableStyle
        Dim Type As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Invdate As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Details As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Value As New DataGridBrowser.DataGridNoActiveCellColumn
        Dim Receipts As New DataGridTextBoxColumn
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
            .Width = 65
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Details
            .HeaderText = "Details"
            .MappingName = "invdetails"
            .Width = 200
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With Value
            .HeaderText = "Value"
            .MappingName = "Amount"
            .Width = 50
            .ReadOnly = True
            .TextBox.ReadOnly = True
            .Format = "#0.00"
        End With

        With Receipts
            .HeaderText = "Receipts" & "  " & Chr(9)
            .Width = 80
            .TextBox.Text = "0.00"
            .Format = "0.00"
            .MappingName = "NullCol1"
            .TextBox.MaxLength = 10
        End With

        With Discount
            .HeaderText = "Discount" & "  " & Chr(9)
            .Width = 80
            .TextBox.Text = "0.00"
            .Format = "0.00"
            .MappingName = "NullCol2"
            .TextBox.MaxLength = 10
        End With

        With MyAuto
            .HeaderText = ""
            .Width = 0
            .Format = "0.00"
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
            .Add(Receipts)
            .Add(Discount)
            .Add(MyAuto)
            .Add(ParentID)
        End With

        GRow = Ds.Tables(0).Rows.Count

        'Bind The Datagrid to Dataset
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
            DgridTstyle.MappingName = Ds.Tables(0).TableName
            DataGrid1.TableStyles.Add(DgridTstyle)
            btnOk.Enabled = False
        End If

        oldCRow = 0
        oldCCol = 0

    End Sub

    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged

        Dim Trows As Integer
        Dim CustId As String, ReceiptDate As Date, r As Integer, TDate As String
        Dim InvNo As Double, InvTp As String, InvDate As Date, InvDetails As String
        Dim InvValue As Double, InvReceipt As Double, InvDiscount As Double
        Dim i As Integer, DisValue As Double, j As Integer
        Dim rt As Integer
        Dim RValue As Double = 0

        If recursiveCall = True Then Exit Sub

        CCol = DataGrid1.CurrentCell.ColumnNumber
        CRow = DataGrid1.CurrentCell.RowNumber

        If oldCCol = 4 Or oldCCol = 5 Then

            If IsNumeric(DataGrid1.Item(oldCRow, oldCCol)) = False Then
                MsgBox("Enter numeric value in receipts.     ", MsgBoxStyle.Information, "Infini Accounts Express")
                DataGrid1.Item(oldCRow, oldCCol) = "0.00"
                Exit Sub
            End If
            recursiveCall = True
            If DataGrid1.Item(oldCRow, oldCCol) = 0 Or DataGrid1.Item(oldCRow, oldCCol) = 0.0 Then
                DataGrid1.Item(oldCRow, oldCCol) = "0.00"
            End If
            recursiveCall = False

            Trows = GRow

            For i = 0 To Trows - 1

                InvTp = Trim(DataGrid1.Item(i, 0)) ' InvTP
                If InvTp = "" Then Exit For
                TDate = DataGrid1.Item(i, 1) ' InvDate
                InvDetails = DataGrid1.Item(i, 2) ' InvDetails
                InvValue = Format((DataGrid1.Item(i, 3)), "0.00") ' InvValue

                InvReceipt = DataGrid1.Item(i, 4) ' Receipt
                If IsNumeric(InvReceipt) = False Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    Exit For
                End If
                InvDiscount = DataGrid1.Item(i, 5) ' Discount
                If IsNumeric(InvDiscount) = False Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    Exit For
                End If

                If InvTp <> "SI" Then
                    If InvDiscount > 0 Then
                        DataGrid1.Item(i, 4) = "0.00"
                        DataGrid1.Item(i, 5) = "0.00"
                        Exit For
                    End If
                End If

                DisValue = CDbl(InvReceipt) + CDbl(InvDiscount)
                If DisValue > InvValue Then
                    DataGrid1.Item(i, 4) = "0.00"
                    DataGrid1.Item(i, 5) = "0.00"
                    MessageBox.Show("Receipt amount is greater then amount of settlement.     ", "Infini Accounts Express", MessageBoxButtons.OK)
                    txtbval.Text = "0.00"
                    txtRval.Text = "0.00"
                    Exit For
                End If

                'If InvReceipt = 0.0 And InvDiscount > 0 And InvDiscount <= InvValue Then
                'If InvDiscount <= InvValue Then
                '    InvReceipt = InvValue - InvDiscount
                '    DataGrid1.Item(i, 4) = Format(InvReceipt, "0.00")
                '    DataGrid1.Item(i, 5) = Format(InvDiscount, "0.00")
                'End If

                If InvTp = "SI" Then
                    RValue = Format(CDbl(RValue) + CDbl(InvReceipt), "0.00")
                Else
                    RValue = Format(CDbl(RValue) - CDbl(InvReceipt), "0.00")
                End If
            Next

            txtbval.Text = Format(CDbl(txt2.Text) + CDbl(RValue), "0.00")
            txtRval.Text = Format(RValue, "0.00")
        End If

        oldCRow = CRow
        oldCCol = CCol

    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim TFlag As Boolean, OSAutoNo As Double
        Dim SRValue As Double, SAValue As Double, Amount As Double
        Dim CustomerID As String = cmbcid.SelectedItem
        Dim a As Object, b As Object

        If TextBox3.Text = "" Then
            MsgBox("Select receipts date first.    ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        DataGrid1_CurrentCellChanged(a, b) 'For calculation

        TFlag = False
        InvDetails = "Sales Receipt"
        InvDate = TextBox3.Text
        TranDate = CDate(Txtdate.Text)

        'Receipts Amount
        Amount = txtRval.Text

        If Amount < 0 Then
            MsgBox("Unsettled cheque balance for this account.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtbval.Text = "0.00"
            txtRval.Text = "0.00"
            Exit Sub
        End If

        row = GRow
        For i = 0 To row - 1
            ReceiptValue = DataGrid1.Item(i, 4)
            DiscountValue = DataGrid1.Item(i, 5)
            If ReceiptValue > 0 Or DiscountValue > 0 Then
                ReceiptValue = 0
                DiscountValue = 0
                TFlag = True
                Exit For
            End If
        Next i

        If Amount > 0 Then TFlag = True

        If TFlag = False Then Exit Sub

        'Calculate Sales Receipts Value
        SRValue = Format(CalculateSR(), "0.00")

        If SRValue < 0 Then
            MsgBox("Unsettled cheque balance for this account.    ", MsgBoxStyle.Exclamation, "Infini Accounts Express")
            Exit Sub
        End If

        SAValue = Format(Amount - SRValue, "0.00")
        If SAValue > 0 Then
            MsgBox("Unsettled cheque balance of " & Format(SAValue, "0.00") & " found for this account.     ", MsgBoxStyle.Exclamation, "Infini Accounts Express")
        End If

        Me.Enabled = False 'Form Enabled False

        DTime = dLast.Now 'Transaction DateTime

        'Get New Transaction No
        ParentNo = ClsGlobal.GetTransactionNo()
        'Get New Ledger Auto No
        AutoNo = ClsGlobal.GetLedgerAutoNo
        'Get New Outstanding Max. No
        OSMaxNo = ClsGlobal.GetOutStandingMaxOSNo()

        'Transaction For SR
        If SRValue > 0 Then

            'Receipt SR Transaction
            ClsGlobal.InsertTransactionData(ParentNo, "SR", "T9", InvDate, InvDetails, SRValue, 0, TranDate, DTime)

            'Debtors Control Account Transaction
            '("DEBTORSCONTROLACCOUNT") = "70100"
            '("DEFAULTBANK") = "70200"
            ClsGlobal.InsertLedger(ParentNo, ParentNo, "70100", 0, SRValue, "T", "70200", "", AutoNo, DTime)

        End If

        row = GRow
        For i = 0 To row - 1

            With DataGrid1

                invtp = Trim(.Item(i, 0)) ' InvTP
                If invtp = "" Then Exit For
                InvValue = Format(CDbl(.Item(i, 3)), "0.00") ' InvValue
                ReceiptValue = .Item(i, 4) ' Receipt
                DiscountValue = .Item(i, 5) ' Discount
                OSAutoNo = .Item(i, 6) ' MyAutoNo

            End With

            If ReceiptValue > 0 Then
                If InvValue = (ReceiptValue + DiscountValue) Then 'Full Receipt
                    ClsGlobal.UpdateLedgerPayStatus(OSAutoNo, "f", DTime)
                ElseIf ReceiptValue > 0 Then 'Partial Receipt
                    ClsGlobal.UpdateLedgerPayStatus(OSAutoNo, "p", DTime)
                End If
            End If
            ReceiptValue = 0

        Next i

        'Calculation Of Sales Invoice "SI"
        CalculationSI()

        If SAValue > 0 Then 'Transaction For SA

            'Get New Transaction No. 
            InvNo = ParentNo + 1

            'Sales On Account SA Transaction
            InvDetails = "Sales On Account"
            ClsGlobal.InsertTransactionData(InvNo, "SA", "T9", InvDate, "Payment On Account", SAValue, 0, TranDate, DTime)

            'Bank Debit Transaction
            'Application("DEFAULTBANK") = "70200"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, "70200", SAValue, 0, "B", Custid, "f", AutoNo, DTime)

            'Customer Credit Transaction
            'Application("DEFAULTBANK") = "70200"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, Custid, 0, SAValue, "C", "70200", "*", AutoNo, DTime)

            'Debtors Control Account Credit Transaction
            'Application("DEFAULTBANK") = "70200"
            'Application("DEBTORSCONTROLACCOUNT") = "70100"
            AutoNo = AutoNo + 1
            ClsGlobal.InsertLedger(InvNo, InvNo, "70100", 0, SAValue, "T", "70200", "", AutoNo, DTime)

        End If

        Me.Enabled = True 'Form Enabled True
        If ClsGlobal.ExecuteCommand() = True Then

            Dim MsgStr As String
            MsgStr = "Receipt has been created.   " & vbNewLine & "Please Synchronize your transactions."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose(True)

        End If

    End Sub

    Private Function CalculateSR() As Double

        Dim RValue As Double
        Dim SIValue As Double, SCValue As Double, InvTP As String
        Dim row As Integer, r As Integer

        row = GRow
        SIValue = 0
        SCValue = 0

        For r = 0 To row - 1

            InvTP = Trim(DataGrid1.Item(r, 0)) ' InvTP
            RValue = DataGrid1.Item(r, 4)
            If InvTP = "" Then Exit For

            If InvTP = "SI" Then
                SIValue = SIValue + CDbl(RValue)
            Else
                SCValue = SCValue + CDbl(RValue)
            End If

        Next r

        Return (SIValue - SCValue)

    End Function

    Public Sub CalculationSI()

        Dim Balance As Double, ReceiptValue As Double
        Dim K As Integer, OSNo As Double, Dtran As Double
        Dim OSAmount As Double, OSTP As String
        Dim KRow As Integer, TAuto As Integer
        Dim dt As DataTable = CType(DataGrid1.DataSource, DataTable)
        Dim dr As DataRow, tr As DataRow
        Dim PBOX As New TextBox, PCBOX As New TextBox

        For Each dr In dt.Rows

            invtp = dr(3)  'InvTP
            If invtp = "" Then Exit For
            Dim r As Object = dr(2)
            InvNo = dr(2) 'InvNo.
            Tdate = dr(5) 'InvDate
            InvDetails = dr(6) 'InvDetails
            PBOX.Text = dr(0) 'Receipts
            ReceiptValue = PBOX.Text

            If invtp = "SI" Then

                While ReceiptValue > 0

                    Balance = 0
                    For Each tr In dt.Rows

                        OSTP = tr(3) 'InvTP
                        If OSTP = "" Then Exit For
                        OSNo = tr(2) 'InvNo.
                        CBOX.Text = tr(0) 'Payment
                        Balance = CDbl(CBOX.Text)
                        If Balance > 0 Then
                            If OSTP = "SC" Or OSTP = "SA" Then
                                CBOX.Text = "0.00"
                                Exit For
                            Else
                                Balance = 0
                            End If
                        End If
                    Next

                    ''''SC Transactions
                    If Balance > 0 Then
                        OSAmount = Balance
                        If OSAmount > ReceiptValue Then
                            OSDet = Format(ReceiptValue, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(ReceiptValue, "0.00") & " To " & invtp & " " & InvNo
                            OSAmount = ReceiptValue
                            PCBOX.Text = Format(OSAmount - ReceiptValue, "0.00")
                            tr(0) = Format(Balance - ReceiptValue, "0.00")
                            ReceiptValue = 0
                            PBOX.Text = "0.00"
                        ElseIf OSAmount = ReceiptValue Then
                            OSDet = Format(ReceiptValue, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(ReceiptValue, "0.00") & " To " & invtp & " " & InvNo
                            tr(0) = 0
                            ReceiptValue = 0
                            PBOX.Text = "0.00"
                            PCBOX.Text = "0.00"
                        ElseIf OSAmount < ReceiptValue Then
                            OSDet = Format(OSAmount, "0.00") & " From " & OSTP & " " & OSNo
                            OSDet2 = Format(OSAmount, "0.00") & " To " & invtp & " " & InvNo
                            ReceiptValue = ReceiptValue - OSAmount
                            tr(0) = 0
                            PBOX.Text = Format(ReceiptValue, "0.00")
                            PCBOX.Text = ReceiptValue
                        End If

                        'OutStanding Debit Customer Transaction
                        'Get ledger MyAuto
                        MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "C")
                        ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, Custid, OSAmount, 0, "C", OSDet, "", InvDate, OSNo, 0, "F", DTime, OSMaxNo)

                        'OutStanding Credit Customer Transaction
                        'Get ledger MyAuto
                        OSMaxNo = OSMaxNo + 1
                        MyAuto = ClsGlobal.GetExistingLedgerAutoNo(OSNo, "C")
                        ClsGlobal.InsertOutStanding(MyAuto, OSNo, OSNo, Custid, 0, OSAmount, "C", OSDet2, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                        If OSTP = "SA" Then
                            'OutStanding Debit Bank Transaction
                            'Application("DEFAULTBANK")="70200"
                            'Get ledger MyAuto
                            OSMaxNo = OSMaxNo + 1
                            ClsGlobal.InsertOutStanding(AutoNo, OSNo, OSNo, "70200", OSAmount, 0, "B", OSDet2, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)
                        End If
                        Balance = 0
                        PBOX.Text = PCBOX.Text
                        ReceiptValue = CDbl(PCBOX.Text)
                        OSMaxNo = OSMaxNo + 1

                    Else
                        'SI Transactions
                        If ReceiptValue > 0 Then
                            'Bank Transaction
                            'Application("DEFAULTBANK")="70200"
                            AutoNo = AutoNo + 1
                            ClsGlobal.InsertLedger(ParentNo, ParentNo, "70200", ReceiptValue, 0, "B", Custid, "f", AutoNo, DTime)

                            'OutStanding Bank Transaction
                            'Application("DEFAULTBANK") = "70200"
                            OSDet = Format(ReceiptValue, "0.00") & " To SI " & InvNo

                            'Get ledger MyAuto
                            ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, "70200", ReceiptValue, 0, "B", OSDet, Custid, InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                            'Customer Transaction
                            'Application("DEFAULTBANK")="70200"
                            AutoNo = AutoNo + 1
                            ClsGlobal.InsertLedger(ParentNo, ParentNo, Custid, 0, ReceiptValue, "C", "70200", "f", AutoNo, DTime)

                            'OutStanding Customer Transaction
                            'Application("DEFAULTBANK")="70200"
                            'Get ledger MyAuto
                            OSMaxNo = OSMaxNo + 1
                            ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, Custid, 0, ReceiptValue, "C", OSDet, "70200", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                            'OutStanding Customer Transaction
                            OSDet = Format(ReceiptValue, "0.00") & " From SR " & ParentNo

                            'Get ledger MyAuto
                            OSMaxNo = OSMaxNo + 1
                            MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "C")
                            ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, Custid, ReceiptValue, 0, "C", OSDet, "", InvDate, InvNo, 0, "F", DTime, OSMaxNo)
                        End If
                        PBOX.Text = "0.00"
                        ReceiptValue = 0
                        OSMaxNo = OSMaxNo + 1
                    End If
                End While

            Else
                ReceiptValue = CDbl(SCBOX.Text)
                If ReceiptValue > 0 Then
                    'Bank Transaction
                    'Application("DEFAULTBANK")="70200"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(ParentNo, ParentNo, "70200", ReceiptValue, 0, "B", Custid, "", AutoNo, DTime)

                    'OutStanding Bank Transaction
                    'Application("DEFAULTBANK") = "70200"
                    OSDet = Format(ReceiptValue, "0.00") & " To SI " & InvNo

                    'Get ledger MyAuto
                    ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, "70200", ReceiptValue, 0, "B", OSDet, Custid, InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'Customer Transaction
                    'Application("DEFAULTBANK")="70200"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(ParentNo, ParentNo, Custid, 0, ReceiptValue, "C", "70200", "f", AutoNo, DTime)

                    'OutStanding Customer Transaction
                    'Application("DEFAULTBANK")="70200"
                    'Get ledger MyAuto
                    OSMaxNo = OSMaxNo + 1
                    ClsGlobal.InsertOutStanding(AutoNo, ParentNo, ParentNo, Custid, 0, ReceiptValue, "C", OSDet, "70200", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'OutStanding Customer Transaction
                    OSDet = Format(ReceiptValue, "0.00") & " From SR " & ParentNo

                    'Get ledger MyAuto
                    OSMaxNo = OSMaxNo + 1
                    MyAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "C")
                    ClsGlobal.InsertOutStanding(MyAuto, InvNo, InvNo, Custid, ReceiptValue, 0, "C", OSDet, "", InvDate, InvNo, 0, "F", DTime, OSMaxNo)
                End If
                ReceiptValue = 0
                PBOX.Text = "0.00"
                OSMaxNo = OSMaxNo + 1
            End If
        Next

        'Get New Transaction No.
        Dtran = ParentNo + 1
        Dim DCheck As Boolean = False

        For Each tr In dt.Rows 'Discount Transactions

            invtp = tr(3) 'InvTP
            If invtp = "" Then Exit For
            Tdate = tr(5) 'InvDate
            InvDetails = tr(6) 'Details
            DiscountValue = tr(1) 'Discount
            InvNo = tr(2) 'Parentid

            If invtp = "SI" Then

                'Discount Transactions
                If DiscountValue > 0 Then 'Transaction For SD

                    DCheck = True

                    'Discount SD Transaction
                    InvDetails = "Sales Discount"
                    ClsGlobal.InsertTransactionData(Dtran, "SD", "T9", InvDate, InvDetails, DiscountValue, 0, TranDate, DTime)

                    'Sales Discount Control Account Transaction
                    'Application("SALESDISCOUNT") = "10100"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(Dtran, Dtran, "10100", DiscountValue, 0, "T", Custid, "", AutoNo, DTime)

                    'Debtors Control Account Transaction
                    'Application("DEBTORSCONTROLACCOUNT") = "70100"
                    'Application("SALESDISCOUNT") = "10100"
                    AutoNo = AutoNo + 1
                    ClsGlobal.InsertLedger(Dtran, Dtran, "70100", 0, DiscountValue, "T", "10100", "", AutoNo, DTime)
                    'Thread.Sleep(200)

                    'Customer Account Transaction
                    'Application("SALESDISCOUNT") = "10100"
                    TAuto = AutoNo + 1
                    ClsGlobal.InsertLedger(Dtran, Dtran, Custid, 0, DiscountValue, "C", "10100", "f", TAuto, DTime)

                    'OutStanding Credit Customer Transaction
                    OSDet = Format(DiscountValue, "0.00") & " Payed To SI " & InvNo

                    'Get ledger MyAuto
                    ClsGlobal.InsertOutStanding(TAuto, Dtran, Dtran, Custid, 0, DiscountValue, "C", OSDet, "", InvDate, 0, InvNo, "T", DTime, OSMaxNo)

                    'OutStanding Debit Customer Transaction
                    OSDet = Format(DiscountValue, "0.00") & " From SD " & Dtran
                    AutoNo = TAuto

                    'Get ledger MyAuto
                    OSMaxNo = OSMaxNo + 1
                    TAuto = ClsGlobal.GetExistingLedgerAutoNo(InvNo, "C")
                    ClsGlobal.InsertOutStanding(TAuto, InvNo, InvNo, Custid, DiscountValue, 0, "C", OSDet, "", InvDate, Dtran, 0, "F", DTime, OSMaxNo)
                    MyAuto = TAuto
                    Dtran = Dtran + 1
                    OSMaxNo = OSMaxNo + 1
                End If
                DiscountValue = 0
            End If
        Next
        If DCheck = True Then
            ParentNo = Dtran - 1
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
