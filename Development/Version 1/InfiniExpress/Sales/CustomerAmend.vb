'Option Strict On
Imports System
Imports InfiniExpress.BLL

Public Class CustomerAmend

    Inherits System.Windows.Forms.Form
    Public ClsSales As New ClassSales()
    Public _dr As IDataReader
    Dim CustID As String

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
    Protected WithEvents txtfax As System.Windows.Forms.TextBox
    Protected WithEvents txttelephone As System.Windows.Forms.TextBox
    Protected WithEvents txtpcode As System.Windows.Forms.TextBox
    Protected WithEvents txtcity As System.Windows.Forms.TextBox
    Protected WithEvents txtstreet As System.Windows.Forms.TextBox
    Protected WithEvents Label9 As System.Windows.Forms.Label
    Protected WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Protected WithEvents btnclear As System.Windows.Forms.Button
    Protected WithEvents btndelete As System.Windows.Forms.Button
    Protected WithEvents btncancel As System.Windows.Forms.Button
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents cmbID As System.Windows.Forms.ComboBox
    Protected WithEvents Label8 As System.Windows.Forms.Label
    Protected WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents Label11 As System.Windows.Forms.Label
    Protected WithEvents Label13 As System.Windows.Forms.Label
    Protected WithEvents txtclimit As System.Windows.Forms.TextBox
    Protected WithEvents txtvat As System.Windows.Forms.TextBox
    Protected WithEvents txtcustomername As System.Windows.Forms.TextBox
    Protected WithEvents txtemail As System.Windows.Forms.TextBox
    Protected WithEvents txtCname As System.Windows.Forms.TextBox
    Protected WithEvents txtcountry As System.Windows.Forms.TextBox
    Protected WithEvents TxtNo As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CustomerAmend))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtcountry = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtclimit = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtvat = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtcustomername = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbID = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtfax = New System.Windows.Forms.TextBox
        Me.txttelephone = New System.Windows.Forms.TextBox
        Me.txtemail = New System.Windows.Forms.TextBox
        Me.txtpcode = New System.Windows.Forms.TextBox
        Me.txtcity = New System.Windows.Forms.TextBox
        Me.txtstreet = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCname = New System.Windows.Forms.TextBox
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnclear = New System.Windows.Forms.Button
        Me.btndelete = New System.Windows.Forms.Button
        Me.btncancel = New System.Windows.Forms.Button
        Me.TxtNo = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtcountry)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.txtclimit)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.txtvat)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.txtcustomername)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.cmbID)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.txtfax)
        Me.Panel1.Controls.Add(Me.txttelephone)
        Me.Panel1.Controls.Add(Me.txtemail)
        Me.Panel1.Controls.Add(Me.txtpcode)
        Me.Panel1.Controls.Add(Me.txtcity)
        Me.Panel1.Controls.Add(Me.txtstreet)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtCname)
        Me.Panel1.Location = New System.Drawing.Point(7, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(382, 376)
        Me.Panel1.TabIndex = 0
        '
        'txtcountry
        '
        Me.txtcountry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcountry.Location = New System.Drawing.Point(12, 156)
        Me.txtcountry.MaxLength = 50
        Me.txtcountry.Name = "txtcountry"
        Me.txtcountry.Size = New System.Drawing.Size(170, 20)
        Me.txtcountry.TabIndex = 3
        Me.txtcountry.Text = ""
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.White
        Me.Label13.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label13.Location = New System.Drawing.Point(12, 140)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(168, 16)
        Me.Label13.TabIndex = 48
        Me.Label13.Text = "Country"
        '
        'txtclimit
        '
        Me.txtclimit.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtclimit.Location = New System.Drawing.Point(194, 272)
        Me.txtclimit.MaxLength = 12
        Me.txtclimit.Name = "txtclimit"
        Me.txtclimit.Size = New System.Drawing.Size(170, 20)
        Me.txtclimit.TabIndex = 10
        Me.txtclimit.Text = "0.00"
        Me.txtclimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.White
        Me.Label11.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label11.Location = New System.Drawing.Point(192, 256)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(170, 16)
        Me.Label11.TabIndex = 45
        Me.Label11.Text = "Credit Limit"
        '
        'txtvat
        '
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Location = New System.Drawing.Point(12, 272)
        Me.txtvat.MaxLength = 20
        Me.txtvat.Name = "txtvat"
        Me.txtvat.Size = New System.Drawing.Size(170, 20)
        Me.txtvat.TabIndex = 9
        Me.txtvat.Text = ""
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.White
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label10.Location = New System.Drawing.Point(10, 256)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(166, 16)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "V.A.T Reg No"
        '
        'txtcustomername
        '
        Me.txtcustomername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcustomername.Location = New System.Drawing.Point(143, 26)
        Me.txtcustomername.MaxLength = 255
        Me.txtcustomername.Name = "txtcustomername"
        Me.txtcustomername.Size = New System.Drawing.Size(216, 20)
        Me.txtcustomername.TabIndex = 1
        Me.txtcustomername.Text = ""
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label8.Location = New System.Drawing.Point(143, 10)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(214, 16)
        Me.Label8.TabIndex = 41
        Me.Label8.Text = "Customer Name  *"
        '
        'cmbID
        '
        Me.cmbID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbID.Location = New System.Drawing.Point(11, 24)
        Me.cmbID.Name = "cmbID"
        Me.cmbID.Size = New System.Drawing.Size(120, 21)
        Me.cmbID.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label7.Location = New System.Drawing.Point(10, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Customer ID"
        '
        'txtfax
        '
        Me.txtfax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfax.Location = New System.Drawing.Point(12, 234)
        Me.txtfax.MaxLength = 80
        Me.txtfax.Name = "txtfax"
        Me.txtfax.Size = New System.Drawing.Size(170, 20)
        Me.txtfax.TabIndex = 7
        Me.txtfax.Text = ""
        '
        'txttelephone
        '
        Me.txttelephone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttelephone.Location = New System.Drawing.Point(194, 196)
        Me.txttelephone.MaxLength = 80
        Me.txttelephone.Name = "txttelephone"
        Me.txttelephone.Size = New System.Drawing.Size(170, 20)
        Me.txttelephone.TabIndex = 6
        Me.txttelephone.Text = ""
        '
        'txtemail
        '
        Me.txtemail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtemail.Location = New System.Drawing.Point(13, 312)
        Me.txtemail.MaxLength = 50
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(354, 20)
        Me.txtemail.TabIndex = 11
        Me.txtemail.Text = ""
        '
        'txtpcode
        '
        Me.txtpcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpcode.Location = New System.Drawing.Point(12, 196)
        Me.txtpcode.MaxLength = 30
        Me.txtpcode.Name = "txtpcode"
        Me.txtpcode.Size = New System.Drawing.Size(170, 20)
        Me.txtpcode.TabIndex = 5
        Me.txtpcode.Text = ""
        '
        'txtcity
        '
        Me.txtcity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcity.Location = New System.Drawing.Point(193, 156)
        Me.txtcity.MaxLength = 50
        Me.txtcity.Name = "txtcity"
        Me.txtcity.Size = New System.Drawing.Size(171, 20)
        Me.txtcity.TabIndex = 4
        Me.txtcity.Text = ""
        '
        'txtstreet
        '
        Me.txtstreet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtstreet.Location = New System.Drawing.Point(11, 62)
        Me.txtstreet.MaxLength = 255
        Me.txtstreet.Multiline = True
        Me.txtstreet.Name = "txtstreet"
        Me.txtstreet.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtstreet.Size = New System.Drawing.Size(350, 76)
        Me.txtstreet.TabIndex = 2
        Me.txtstreet.Text = ""
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label9.Location = New System.Drawing.Point(11, 46)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(348, 16)
        Me.Label9.TabIndex = 32
        Me.Label9.Text = "Street"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label6.Location = New System.Drawing.Point(10, 218)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(168, 16)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Fax"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label5.Location = New System.Drawing.Point(192, 180)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(166, 16)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Telephone "
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 296)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(166, 16)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Email *"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 180)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(164, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Postal Code"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(191, 140)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 16)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "City/Sate"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(192, 218)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(170, 12)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Contact Name"
        '
        'txtCname
        '
        Me.txtCname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCname.Location = New System.Drawing.Point(194, 234)
        Me.txtCname.MaxLength = 255
        Me.txtCname.Name = "txtCname"
        Me.txtCname.Size = New System.Drawing.Size(170, 20)
        Me.txtCname.TabIndex = 8
        Me.txtCname.Text = ""
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnOk.Location = New System.Drawing.Point(38, 354)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 22)
        Me.btnOk.TabIndex = 12
        Me.btnOk.Text = "Save"
        '
        'btnclear
        '
        Me.btnclear.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnclear.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnclear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclear.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnclear.Location = New System.Drawing.Point(118, 354)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(75, 22)
        Me.btnclear.TabIndex = 12
        Me.btnclear.Text = "&Clear"
        '
        'btndelete
        '
        Me.btndelete.BackColor = System.Drawing.Color.LightSlateGray
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btndelete.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndelete.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btndelete.Location = New System.Drawing.Point(198, 354)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(75, 22)
        Me.btndelete.TabIndex = 13
        Me.btndelete.Text = "&Delete"
        '
        'btncancel
        '
        Me.btncancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.btncancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btncancel.Location = New System.Drawing.Point(278, 354)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.Size = New System.Drawing.Size(75, 22)
        Me.btncancel.TabIndex = 14
        Me.btncancel.Text = "&Cancel"
        '
        'TxtNo
        '
        Me.TxtNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtNo.Location = New System.Drawing.Point(240, 424)
        Me.TxtNo.MaxLength = 12
        Me.TxtNo.Name = "TxtNo"
        Me.TxtNo.Size = New System.Drawing.Size(72, 20)
        Me.TxtNo.TabIndex = 15
        Me.TxtNo.Text = "0"
        Me.TxtNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtNo.Visible = False
        '
        'CustomerAmend
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(394, 391)
        Me.Controls.Add(Me.TxtNo)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnclear)
        Me.Controls.Add(Me.btndelete)
        Me.Controls.Add(Me.btncancel)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustomerAmend"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Amend"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CustomerAmend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        _dr = ClsSales.LoadCustomerID
        While _dr.Read
            cmbID.Items.Add(_dr.Item("Customerid"))
        End While
        _dr.Close()

    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Private Sub cmbID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbID.SelectedValueChanged

        CustID = CStr(cmbID.SelectedItem)

        ClearTextBox()
        _dr = ClsSales.GetCustInfo(CustID)

        If _dr.Read Then
            If IsDBNull(_dr("Customername")) = False Then txtcustomername.Text = _dr("Customername")
            If IsDBNull(_dr("street")) = False Then txtstreet.Text = _dr("street")
            If IsDBNull(_dr("town")) = False Then txtcity.Text = _dr("town")
            If IsDBNull(_dr("country")) = False Then txtcountry.Text = _dr("country")
            If IsDBNull(_dr("Email")) = False Then txtemail.Text = _dr("Email")
            If IsDBNull(_dr("postalcode")) = False Then txtpcode.Text = _dr("postalcode")
            If IsDBNull(_dr("contactname")) = False Then txtCname.Text = _dr("contactname")
            If IsDBNull(_dr("telephone")) = False Then txttelephone.Text = _dr("telephone")
            If IsDBNull(_dr("fax")) = False Then txtfax.Text = _dr("fax")
            If IsDBNull(_dr("vat")) = False Then txtvat.Text = _dr("vat")
            txtclimit.Text = Format(_dr("creditlimit"), "0.00")
            TxtNo.Text = _dr("UniqueNo")
        End If
        _dr.Close()

    End Sub

    Sub ClearTextBox()

        Me.txtcustomername.Text = ""
        Me.txtstreet.Text = ""
        Me.txtcity.Text = ""
        Me.txtpcode.Text = ""
        Me.txttelephone.Text = ""
        Me.txtemail.Text = ""
        Me.txtfax.Text = ""
        Me.txtvat.Text = ""
        Me.txtcountry.Text = ""
        Me.txtCname.Text = ""
        Me.txtclimit.Text = "0.00"
        Me.TxtNo.Text = 0

    End Sub

    Private Sub btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click
        _dr = ClsSales.LoadCustomerID
        cmbID.Items.Clear()
        While _dr.Read
            cmbID.Items.Add(_dr.Item("Customerid"))
        End While
        _dr.Close()
        ClearTextBox()
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim GlbGlobal As New ClassGlobal()

        If Trim(cmbID.SelectedItem) = "" Then
            MsgBox("Select customer ID first.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        If Len(txtcustomername.Text) = 0 Then
            MsgBox("Enter customer name.      ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        CustID = cmbID.SelectedItem()

        If IsNumeric(txtclimit.Text) = False Then
            MsgBox("Enter numeric value in credit limit.      ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        If txtemail.Text.IndexOf("@"c) = -1 Or txtemail.Text.IndexOf("."c) = -1 Or txtemail.Text.IndexOf("@"c) <> txtemail.Text.LastIndexOf("@"c) Then
            MsgBox("Enter a valid Email address.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        Dim i As Integer

        For i = 0 To Me.Panel1.Controls.Count - 1
            If Convert.ToString(Me.Panel1.Controls.Item(i).GetType()) = "System.Windows.Forms.TextBox" Then
                If CType(Me.Panel1.Controls.Item(i), System.Windows.Forms.TextBox).Text.IndexOf("'"c) <> -1 Then
                    MsgBox("Character "" ' ""  not allowed.    ", MsgBoxStyle.Critical, "Infini Accounts Express")
                    Exit Sub
                End If

            End If
        Next

        Dim DTime As String, dLast As Date
        DTime = dLast.Now

        'Update Customer Record
        ClsSales.AmendCustomer(CustID, txtcustomername.Text, txtstreet.Text, txtcity.Text, txtcountry.Text, txtpcode.Text, txtCname.Text, txtemail.Text, txttelephone.Text, txtfax.Text, txtvat.Text, CDbl(txtclimit.Text))

        If (GlbGlobal.ExecuteCommand() = True) Then
            Dim MsgStr As String
            MsgStr = "Customer account has been updated.    " & vbNewLine & "Please Synchronize your database."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose(True)
        End If

    End Sub

    Private Sub txtclimit_TextLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtclimit.Leave ' .TextChanged

        If IsNumeric(txtclimit.Text) = False Then
            txtclimit.Text = "0.00"
        End If
        txtclimit.Text = Format(Val(txtclimit.Text), "0.00")

    End Sub
    Private Sub Email_Validation(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtemail.Leave
        If txtemail.Text <> "" Then
            If txtemail.Text.IndexOf("@"c) = -1 Or txtemail.Text.IndexOf("."c) = -1 Or txtemail.Text.IndexOf("@"c) <> txtemail.Text.LastIndexOf("@"c) Then
                MessageBox.Show("Enter a valid Email address.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtemail.Text = ""
            End If
        End If
    End Sub

    'Private Sub CustomerName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcustomername.Leave
    '    If txtcustomername.Text = "" Then
    '        MessageBox.Show("Customer name cannot Null.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '    End If
    'End Sub

    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Dim id As Integer
        Dim Flag As Integer
        id = cmbID.SelectedItem

        Flag = ClsSales.CustomerDeleted(id)
        If Flag = 1 Then
            MsgBox("Customer" & "'" & id & "'" & " has been deleted.", MsgBoxStyle.Information, "Infini Accounts Express")
            cmbID.Items.Remove(id)
        Else
            MsgBox("You cannot delete this Customer " & "'" & id & "'" & " because It has associated transactions.", MsgBoxStyle.Information, "Infini Accounts Express")
        End If

        'Clear TextBox's
        ClearTextBox()

    End Sub

End Class