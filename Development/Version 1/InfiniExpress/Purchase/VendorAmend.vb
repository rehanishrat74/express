Imports InfiniExpress.BLL
Public Class VendorAmend
    Inherits Windows.Forms.Form

#Region "Local Variables"
    Dim ClsPurchase As New ClassPurchases()
    Public _dr As IDataReader
    Public clsSales As New ClassSales()
    Public VendID As String
    Public VendorName As String
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
    Protected WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents Label11 As System.Windows.Forms.Label
    Protected WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents cmbID As System.Windows.Forms.ComboBox
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents txtfax As System.Windows.Forms.TextBox
    Protected WithEvents txttelephone As System.Windows.Forms.TextBox
    Protected WithEvents txtpcode As System.Windows.Forms.TextBox
    Protected WithEvents txtcity As System.Windows.Forms.TextBox
    Protected WithEvents txtstreet As System.Windows.Forms.TextBox
    Protected WithEvents Label9 As System.Windows.Forms.Label
    Protected WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents Label2 As System.Windows.Forms.Label
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Protected WithEvents btnclear As System.Windows.Forms.Button
    Protected WithEvents btndelete As System.Windows.Forms.Button
    Protected WithEvents btncancel As System.Windows.Forms.Button
    Protected WithEvents txtvendorname As System.Windows.Forms.TextBox
    Protected WithEvents txtclimit As System.Windows.Forms.TextBox
    Protected WithEvents txtvat As System.Windows.Forms.TextBox
    Protected WithEvents txtcname As System.Windows.Forms.TextBox
    Protected WithEvents txtcountry As System.Windows.Forms.TextBox
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Label12 As System.Windows.Forms.Label
    Protected WithEvents Label13 As System.Windows.Forms.Label
    Protected WithEvents txtemail As System.Windows.Forms.TextBox
    Protected WithEvents TxtNo As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VendorAmend))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtemail = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtcountry = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.btndelete = New System.Windows.Forms.Button()
        Me.btncancel = New System.Windows.Forms.Button()
        Me.txtclimit = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtvat = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtvendorname = New System.Windows.Forms.TextBox()
        Me.cmbID = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtfax = New System.Windows.Forms.TextBox()
        Me.txttelephone = New System.Windows.Forms.TextBox()
        Me.txtpcode = New System.Windows.Forms.TextBox()
        Me.txtcity = New System.Windows.Forms.TextBox()
        Me.txtstreet = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtcname = New System.Windows.Forms.TextBox()
        Me.TxtNo = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label13, Me.txtemail, Me.Label12, Me.txtcountry, Me.Label4, Me.btnOk, Me.btnclear, Me.btndelete, Me.btncancel, Me.txtclimit, Me.Label11, Me.txtvat, Me.Label10, Me.txtvendorname, Me.cmbID, Me.Label7, Me.txtfax, Me.txttelephone, Me.txtpcode, Me.txtcity, Me.txtstreet, Me.Label9, Me.Label6, Me.Label5, Me.Label3, Me.Label2, Me.Label1, Me.txtcname})
        Me.Panel1.Location = New System.Drawing.Point(7, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(382, 376)
        Me.Panel1.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.White
        Me.Label13.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label13.Location = New System.Drawing.Point(10, 302)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 12)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "Email *"
        '
        'txtemail
        '
        Me.txtemail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtemail.Location = New System.Drawing.Point(12, 318)
        Me.txtemail.MaxLength = 255
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(354, 20)
        Me.txtemail.TabIndex = 11
        Me.txtemail.Text = ""
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.White
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label12.Location = New System.Drawing.Point(146, 18)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(112, 12)
        Me.Label12.TabIndex = 53
        Me.Label12.Text = "Vendor Name *"
        '
        'txtcountry
        '
        Me.txtcountry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcountry.Location = New System.Drawing.Point(12, 164)
        Me.txtcountry.MaxLength = 50
        Me.txtcountry.Name = "txtcountry"
        Me.txtcountry.Size = New System.Drawing.Size(168, 20)
        Me.txtcountry.TabIndex = 3
        Me.txtcountry.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(10, 148)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(166, 16)
        Me.Label4.TabIndex = 51
        Me.Label4.Text = "Country"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnOk.Location = New System.Drawing.Point(38, 348)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 22)
        Me.btnOk.TabIndex = 50
        Me.btnOk.Text = "Save"
        '
        'btnclear
        '
        Me.btnclear.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnclear.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnclear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclear.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnclear.Location = New System.Drawing.Point(119, 348)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(75, 22)
        Me.btnclear.TabIndex = 49
        Me.btnclear.Text = "Clear"
        '
        'btndelete
        '
        Me.btndelete.BackColor = System.Drawing.Color.LightSlateGray
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btndelete.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndelete.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btndelete.Location = New System.Drawing.Point(200, 348)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(75, 22)
        Me.btndelete.TabIndex = 48
        Me.btndelete.Text = "Delete"
        '
        'btncancel
        '
        Me.btncancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.btncancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btncancel.Location = New System.Drawing.Point(282, 348)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.Size = New System.Drawing.Size(75, 22)
        Me.btncancel.TabIndex = 47
        Me.btncancel.Text = "Cancel"
        '
        'txtclimit
        '
        Me.txtclimit.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtclimit.Location = New System.Drawing.Point(194, 280)
        Me.txtclimit.MaxLength = 12
        Me.txtclimit.Name = "txtclimit"
        Me.txtclimit.Size = New System.Drawing.Size(172, 20)
        Me.txtclimit.TabIndex = 10
        Me.txtclimit.Text = "0.00"
        Me.txtclimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.White
        Me.Label11.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label11.Location = New System.Drawing.Point(192, 264)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(170, 16)
        Me.Label11.TabIndex = 45
        Me.Label11.Text = "Credit Limit"
        '
        'txtvat
        '
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Location = New System.Drawing.Point(13, 280)
        Me.txtvat.MaxLength = 20
        Me.txtvat.Name = "txtvat"
        Me.txtvat.Size = New System.Drawing.Size(168, 20)
        Me.txtvat.TabIndex = 9
        Me.txtvat.Text = ""
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.White
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label10.Location = New System.Drawing.Point(10, 264)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(166, 16)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "V.A.T Reg No"
        '
        'txtvendorname
        '
        Me.txtvendorname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvendorname.Location = New System.Drawing.Point(146, 34)
        Me.txtvendorname.MaxLength = 255
        Me.txtvendorname.Name = "txtvendorname"
        Me.txtvendorname.Size = New System.Drawing.Size(216, 20)
        Me.txtvendorname.TabIndex = 1
        Me.txtvendorname.Text = ""
        '
        'cmbID
        '
        Me.cmbID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbID.Location = New System.Drawing.Point(14, 34)
        Me.cmbID.Name = "cmbID"
        Me.cmbID.Size = New System.Drawing.Size(114, 21)
        Me.cmbID.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label7.Location = New System.Drawing.Point(14, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Vendor ID"
        '
        'txtfax
        '
        Me.txtfax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfax.Location = New System.Drawing.Point(194, 202)
        Me.txtfax.MaxLength = 80
        Me.txtfax.Name = "txtfax"
        Me.txtfax.Size = New System.Drawing.Size(172, 20)
        Me.txtfax.TabIndex = 6
        Me.txtfax.Text = ""
        '
        'txttelephone
        '
        Me.txttelephone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttelephone.Location = New System.Drawing.Point(12, 202)
        Me.txttelephone.MaxLength = 80
        Me.txttelephone.Name = "txttelephone"
        Me.txttelephone.Size = New System.Drawing.Size(168, 20)
        Me.txttelephone.TabIndex = 5
        Me.txttelephone.Text = ""
        '
        'txtpcode
        '
        Me.txtpcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpcode.Location = New System.Drawing.Point(12, 242)
        Me.txtpcode.MaxLength = 30
        Me.txtpcode.Name = "txtpcode"
        Me.txtpcode.Size = New System.Drawing.Size(168, 20)
        Me.txtpcode.TabIndex = 7
        Me.txtpcode.Text = ""
        '
        'txtcity
        '
        Me.txtcity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcity.Location = New System.Drawing.Point(194, 164)
        Me.txtcity.MaxLength = 50
        Me.txtcity.Name = "txtcity"
        Me.txtcity.Size = New System.Drawing.Size(172, 20)
        Me.txtcity.TabIndex = 4
        Me.txtcity.Text = ""
        '
        'txtstreet
        '
        Me.txtstreet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtstreet.Location = New System.Drawing.Point(14, 78)
        Me.txtstreet.MaxLength = 255
        Me.txtstreet.Multiline = True
        Me.txtstreet.Name = "txtstreet"
        Me.txtstreet.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtstreet.Size = New System.Drawing.Size(350, 66)
        Me.txtstreet.TabIndex = 2
        Me.txtstreet.Text = ""
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label9.Location = New System.Drawing.Point(14, 62)
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
        Me.Label6.Location = New System.Drawing.Point(194, 188)
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
        Me.Label5.Location = New System.Drawing.Point(10, 186)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(166, 16)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Telephone "
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 226)
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
        Me.Label2.Location = New System.Drawing.Point(192, 148)
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
        Me.Label1.Location = New System.Drawing.Point(192, 226)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(170, 12)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Contact Name"
        '
        'txtcname
        '
        Me.txtcname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcname.Location = New System.Drawing.Point(194, 242)
        Me.txtcname.MaxLength = 255
        Me.txtcname.Name = "txtcname"
        Me.txtcname.Size = New System.Drawing.Size(172, 20)
        Me.txtcname.TabIndex = 8
        Me.txtcname.Text = ""
        '
        'TxtNo
        '
        Me.TxtNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtNo.Location = New System.Drawing.Point(208, 400)
        Me.TxtNo.MaxLength = 12
        Me.TxtNo.Name = "TxtNo"
        Me.TxtNo.Size = New System.Drawing.Size(80, 20)
        Me.TxtNo.TabIndex = 7
        Me.TxtNo.Text = "0"
        Me.TxtNo.Visible = False
        '
        'VendorAmend
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(394, 391)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TxtNo, Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VendorAmend"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Vendor Amend"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub VendorAmend_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _dr = ClsPurchase.LoadVendorID
        cmbID.Items.Clear()
        While _dr.Read
            cmbID.Items.Add(_dr.Item("Vendorid"))
        End While
        _dr.Close()
    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Private Sub cmbID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbID.SelectedIndexChanged

        VendID = CStr(cmbID.SelectedItem)

        ClearTextBox()
        _dr = ClsPurchase.GetVendorDetails(VendID)
        If _dr.Read Then
            If IsDBNull(_dr("vendorname")) = False Then txtvendorname.Text = _dr("vendorname")
            If IsDBNull(_dr("street")) = False Then txtstreet.Text = _dr("street")
            If IsDBNull(_dr("town")) = False Then txtcity.Text = _dr("town")
            If IsDBNull(_dr("country")) = False Then txtcountry.Text = _dr("country")
            If IsDBNull(_dr("email")) = False Then txtemail.Text = _dr("email")
            If IsDBNull(_dr("postalcode")) = False Then txtpcode.Text = _dr("postalcode")
            If IsDBNull(_dr("contactname")) = False Then txtcname.Text = _dr("contactname")
            If IsDBNull(_dr("telephone")) = False Then txttelephone.Text = _dr("telephone")
            If IsDBNull(_dr("fax")) = False Then txtfax.Text = _dr("fax")
            If IsDBNull(_dr("vat")) = False Then txtvat.Text = _dr("vat")
            txtclimit.Text = Format(_dr("creditlimit"), "0.00")
            TxtNo.Text = _dr("UniqueNo")
        End If
        _dr.Close()

    End Sub

    Sub ClearTextBox()
        Me.txtvendorname.Text = ""
        Me.txtstreet.Text = ""
        Me.txtcity.Text = ""
        Me.txtpcode.Text = ""
        Me.txttelephone.Text = ""
        Me.txtemail.Text = ""
        Me.txtfax.Text = ""
        Me.txtvat.Text = ""
        Me.txtcountry.Text = ""
        Me.txtcname.Text = ""
        Me.txtclimit.Text = "0.00"
        Me.TxtNo.Text = 0

    End Sub

    Private Sub btnclear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click
        _dr = ClsPurchase.LoadVendorID
        cmbID.Items.Clear()
        While _dr.Read
            cmbID.Items.Add(_dr.Item("Vendorid"))
        End While
        _dr.Close()
        ClearTextBox()
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim GlbGlobal As New ClassGlobal()

        If cmbID.SelectedItem = "" Then
            MsgBox("Select vendor ID first.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If
        VendID = cmbID.SelectedItem()

        If IsNumeric(txtclimit.Text) = False Then
            MsgBox("Enter numeric value in credit limit field.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtclimit.Focus()
            Exit Sub
        End If

        If Len(txtclimit.Text) = 0 Then
            MsgBox("Enter credit limit value.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtclimit.Focus()
            Exit Sub
        End If

        If Len(txtemail.Text) = 0 Then
            MsgBox("Enter vendor email address.      ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtemail.Focus()
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

        'Update Vendor Record
        ClsPurchase.UpdateVendor(VendID, txtvendorname.Text, txtstreet.Text, txtcity.Text, txtcountry.Text, txtpcode.Text, txtcname.Text, txtemail.Text, txttelephone.Text, txtfax.Text, txtvat.Text, CDbl(txtclimit.Text), DTime, CInt(TxtNo.Text))

        If (GlbGlobal.ExecuteCommand() = True) Then
            Dim MsgStr As String
            MsgStr = "Vendor account has been updated.    " & vbNewLine & "Please Synchronize your database."
            MessageBox.Show(MsgStr, "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose(True)
        End If

    End Sub

    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Dim Flag As Integer
        VendID = cmbID.SelectedItem

        Flag = ClsPurchase.VendorDeleted(VendID)
        If Flag = 1 Then
            MsgBox("Vendor" & "'" & VendID & "'" & " has been deleted.", MsgBoxStyle.Information, "Infini Accounts Express")
            cmbID.Items.Remove(VendID)
        Else
            MsgBox("You cannot delete this vendor " & "'" & VendID & "'" & " because It has associated transactions.", MsgBoxStyle.Information, "Infini Accounts Express")
        End If
        'Clear TextBox's
        ClearTextBox()

    End Sub

    Private Sub Climit_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtclimit.Leave
        If IsNumeric(txtclimit.Text) = False Then
            txtclimit.Text = "0.00"
        End If
        txtclimit.Text = Format(Val(txtclimit.Text), "0.00")
    End Sub

    'Private Sub VendorName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendorname.Enter
    '    VendorName = txtvendorname.Text
    'End Sub

    'Private Sub VendorName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtvendorname.Leave
    '    If txtvendorname.Text = "" Then
    '        MessageBox.Show("Vendor name cannot be null.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        txtvendorname.Text = VendorName
    '    End If
    'End Sub

    Private Sub Email_Validation(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtemail.Leave
        If txtemail.Text <> "" Then
            If txtemail.Text.IndexOf("@"c) = -1 Or txtemail.Text.IndexOf("."c) = -1 Or txtemail.Text.IndexOf("@"c) <> txtemail.Text.LastIndexOf("@"c) Then
                MessageBox.Show("Enter a valid email address.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtemail.Text = ""
            End If
        End If
    End Sub

End Class
