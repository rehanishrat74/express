Imports InfiniExpress.BLL
Public Class CreateVendor
    Inherits System.Windows.Forms.Form
    Dim IEDService As New InfiniExpress.InfiniWebService.InfiniService
    Public ClsPurchase As New ClassPurchases()

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
    ' Protected WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Protected WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents txtcountry As System.Windows.Forms.TextBox
    Protected WithEvents Label13 As System.Windows.Forms.Label
    Protected WithEvents txtclimit As System.Windows.Forms.TextBox
    Protected WithEvents Label12 As System.Windows.Forms.Label
    Protected WithEvents txtvat As System.Windows.Forms.TextBox
    Protected WithEvents Label14 As System.Windows.Forms.Label
    Protected WithEvents Label15 As System.Windows.Forms.Label
    Protected WithEvents txtfax As System.Windows.Forms.TextBox
    Protected WithEvents txttelephone As System.Windows.Forms.TextBox
    Protected WithEvents txtemail As System.Windows.Forms.TextBox
    Protected WithEvents txtpcode As System.Windows.Forms.TextBox
    Protected WithEvents txtcity As System.Windows.Forms.TextBox
    Protected WithEvents txtstreet As System.Windows.Forms.TextBox
    Protected WithEvents Label17 As System.Windows.Forms.Label
    Protected WithEvents Label18 As System.Windows.Forms.Label
    Protected WithEvents Label23 As System.Windows.Forms.Label
    Protected WithEvents Label25 As System.Windows.Forms.Label
    Protected WithEvents Label26 As System.Windows.Forms.Label
    Protected WithEvents Label27 As System.Windows.Forms.Label
    Protected WithEvents Label28 As System.Windows.Forms.Label
    Protected WithEvents btncancel As System.Windows.Forms.Button
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Protected WithEvents txtVname As System.Windows.Forms.TextBox
    Protected WithEvents txtcontactname As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CreateVendor))
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btncancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.txtcountry = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtclimit = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtvat = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtVname = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtfax = New System.Windows.Forms.TextBox
        Me.txttelephone = New System.Windows.Forms.TextBox
        Me.txtemail = New System.Windows.Forms.TextBox
        Me.txtpcode = New System.Windows.Forms.TextBox
        Me.txtcity = New System.Windows.Forms.TextBox
        Me.txtstreet = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.txtcontactname = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.White
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Location = New System.Drawing.Point(132, 182)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(392, 38)
        Me.Panel5.TabIndex = 24
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btncancel)
        Me.Panel1.Controls.Add(Me.btnOk)
        Me.Panel1.Controls.Add(Me.txtcountry)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.txtclimit)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.txtvat)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.txtVname)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.txtfax)
        Me.Panel1.Controls.Add(Me.txttelephone)
        Me.Panel1.Controls.Add(Me.txtemail)
        Me.Panel1.Controls.Add(Me.txtpcode)
        Me.Panel1.Controls.Add(Me.txtcity)
        Me.Panel1.Controls.Add(Me.txtstreet)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.Label25)
        Me.Panel1.Controls.Add(Me.Label26)
        Me.Panel1.Controls.Add(Me.Label27)
        Me.Panel1.Controls.Add(Me.Label28)
        Me.Panel1.Controls.Add(Me.txtcontactname)
        Me.Panel1.Location = New System.Drawing.Point(2, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(382, 376)
        Me.Panel1.TabIndex = 26
        '
        'btncancel
        '
        Me.btncancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.btncancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btncancel.Location = New System.Drawing.Point(200, 344)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.TabIndex = 49
        Me.btncancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnOk.Location = New System.Drawing.Point(292, 344)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 50
        Me.btnOk.Text = "Save"
        '
        'txtcountry
        '
        Me.txtcountry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcountry.Location = New System.Drawing.Point(12, 156)
        Me.txtcountry.MaxLength = 49
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
        Me.txtclimit.MaxLength = 7
        Me.txtclimit.Name = "txtclimit"
        Me.txtclimit.Size = New System.Drawing.Size(170, 20)
        Me.txtclimit.TabIndex = 10
        Me.txtclimit.Text = "0.00"
        Me.txtclimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.White
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label12.Location = New System.Drawing.Point(192, 256)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(170, 16)
        Me.Label12.TabIndex = 45
        Me.Label12.Text = "Credit Limit"
        '
        'txtvat
        '
        Me.txtvat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtvat.Location = New System.Drawing.Point(12, 272)
        Me.txtvat.MaxLength = 19
        Me.txtvat.Name = "txtvat"
        Me.txtvat.Size = New System.Drawing.Size(170, 20)
        Me.txtvat.TabIndex = 9
        Me.txtvat.Text = ""
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.White
        Me.Label14.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label14.Location = New System.Drawing.Point(10, 256)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(166, 16)
        Me.Label14.TabIndex = 43
        Me.Label14.Text = "V.A.T Reg No"
        '
        'txtVname
        '
        Me.txtVname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVname.Location = New System.Drawing.Point(11, 24)
        Me.txtVname.MaxLength = 254
        Me.txtVname.Name = "txtVname"
        Me.txtVname.Size = New System.Drawing.Size(352, 20)
        Me.txtVname.TabIndex = 1
        Me.txtVname.Text = ""
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.White
        Me.Label15.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label15.Location = New System.Drawing.Point(11, 6)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(121, 14)
        Me.Label15.TabIndex = 41
        Me.Label15.Text = "Vendor Name  *"
        '
        'txtfax
        '
        Me.txtfax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtfax.Location = New System.Drawing.Point(12, 234)
        Me.txtfax.MaxLength = 79
        Me.txtfax.Name = "txtfax"
        Me.txtfax.Size = New System.Drawing.Size(170, 20)
        Me.txtfax.TabIndex = 7
        Me.txtfax.Text = ""
        '
        'txttelephone
        '
        Me.txttelephone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txttelephone.Location = New System.Drawing.Point(194, 196)
        Me.txttelephone.MaxLength = 79
        Me.txttelephone.Name = "txttelephone"
        Me.txttelephone.Size = New System.Drawing.Size(170, 20)
        Me.txttelephone.TabIndex = 6
        Me.txttelephone.Text = ""
        '
        'txtemail
        '
        Me.txtemail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtemail.Location = New System.Drawing.Point(13, 312)
        Me.txtemail.MaxLength = 49
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(354, 20)
        Me.txtemail.TabIndex = 11
        Me.txtemail.Text = ""
        '
        'txtpcode
        '
        Me.txtpcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtpcode.Location = New System.Drawing.Point(12, 196)
        Me.txtpcode.MaxLength = 29
        Me.txtpcode.Name = "txtpcode"
        Me.txtpcode.Size = New System.Drawing.Size(170, 20)
        Me.txtpcode.TabIndex = 5
        Me.txtpcode.Text = ""
        '
        'txtcity
        '
        Me.txtcity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcity.Location = New System.Drawing.Point(193, 156)
        Me.txtcity.MaxLength = 49
        Me.txtcity.Name = "txtcity"
        Me.txtcity.Size = New System.Drawing.Size(171, 20)
        Me.txtcity.TabIndex = 4
        Me.txtcity.Text = ""
        '
        'txtstreet
        '
        Me.txtstreet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtstreet.Location = New System.Drawing.Point(11, 62)
        Me.txtstreet.MaxLength = 254
        Me.txtstreet.Multiline = True
        Me.txtstreet.Name = "txtstreet"
        Me.txtstreet.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtstreet.Size = New System.Drawing.Size(350, 76)
        Me.txtstreet.TabIndex = 2
        Me.txtstreet.Text = ""
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.White
        Me.Label17.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label17.Location = New System.Drawing.Point(11, 46)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(348, 16)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Street"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.White
        Me.Label18.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label18.Location = New System.Drawing.Point(10, 218)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(168, 16)
        Me.Label18.TabIndex = 31
        Me.Label18.Text = "Fax"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.White
        Me.Label23.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label23.Location = New System.Drawing.Point(192, 180)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(166, 16)
        Me.Label23.TabIndex = 30
        Me.Label23.Text = "Telephone "
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.White
        Me.Label25.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label25.Location = New System.Drawing.Point(10, 296)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(166, 16)
        Me.Label25.TabIndex = 29
        Me.Label25.Text = "Email *"
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.White
        Me.Label26.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label26.Location = New System.Drawing.Point(11, 180)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(164, 16)
        Me.Label26.TabIndex = 28
        Me.Label26.Text = "Postal Code"
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.White
        Me.Label27.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label27.Location = New System.Drawing.Point(191, 140)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(168, 16)
        Me.Label27.TabIndex = 27
        Me.Label27.Text = "City/Sate"
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.White
        Me.Label28.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label28.Location = New System.Drawing.Point(192, 218)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(170, 12)
        Me.Label28.TabIndex = 19
        Me.Label28.Text = "Contact Name"
        '
        'txtcontactname
        '
        Me.txtcontactname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcontactname.Location = New System.Drawing.Point(194, 234)
        Me.txtcontactname.MaxLength = 254
        Me.txtcontactname.Name = "txtcontactname"
        Me.txtcontactname.Size = New System.Drawing.Size(170, 20)
        Me.txtcontactname.TabIndex = 8
        Me.txtcontactname.Text = ""
        '
        'CreateVendor
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(386, 383)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CreateVendor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Create Vendor"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Sub ClearTextbox()
        Me.txtcity.Text = ""
        Me.txtclimit.Text = ""
        Me.txtVname.Text = ""
        Me.txtcountry.Text = ""
        Me.txtcontactname.Text = ""
        Me.txtemail.Text = ""
        Me.txtfax.Text = ""
        Me.txtcity.Text = ""
        Me.txtpcode.Text = ""
        Me.txtstreet.Text = ""
        Me.txtvat.Text = ""
        Me.txttelephone.Text = ""
    End Sub

    Public Function CreateTableforOnlineVendor(ByVal CustomerId As String) As DataTable

        Dim dt As New DataTable("mytable")
        Dim dr As DataRow = dt.NewRow()
        'Dim Ldate As String = Date.Now

        dt.Columns.Add(New DataColumn("Vname", GetType(String)))
        dt.Columns.Add(New DataColumn("street", GetType(String)))
        dt.Columns.Add(New DataColumn("city", GetType(String)))
        dt.Columns.Add(New DataColumn("country", GetType(String)))
        dt.Columns.Add(New DataColumn("pcode", GetType(String)))
        dt.Columns.Add(New DataColumn("cname", GetType(String)))
        dt.Columns.Add(New DataColumn("email", GetType(String)))
        dt.Columns.Add(New DataColumn("telephone", GetType(String)))
        dt.Columns.Add(New DataColumn("fax", GetType(String)))
        dt.Columns.Add(New DataColumn("vat", GetType(String)))
        dt.Columns.Add(New DataColumn("climit", GetType(Double)))
        dt.Columns.Add(New DataColumn("CustomerID", GetType(String)))
        'dt.Columns.Add(New DataColumn("Ldate", GetType(String)))

        dr("Vname") = Trim(txtVname.Text)
        dr("street") = Trim(txtstreet.Text)
        dr("city") = Trim(txtcity.Text)
        dr("country") = Trim(txtcountry.Text)
        dr("pcode") = Trim(txtpcode.Text)
        dr("cname") = Trim(txtcontactname.Text)
        dr("email") = Trim(txtemail.Text)
        dr("telephone") = Trim(txttelephone.Text)
        dr("fax") = Trim(txtfax.Text)
        dr("vat") = Trim(txtvat.Text)
        dr("climit") = CType(Trim(txtclimit.Text), Double)
        dr("CustomerId") = CustomerId
        'dr("Ldate") = Ldate
        dt.Rows.Add(dr)

        Return dt

    End Function

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim Random_Number As String, MaxID As Int32
        Dim VendId As String
        Dim GlbGlobal As New ClassGlobal
        Dim VendorTable As DataTable
        Dim dsCustomerId As New DataSet
        Dim dtCustomerId As New DataTable
        Dim drCustomerId As DataRow
        Dim CustomerId As String
        Dim dtInsertVendor As DataTable
        Dim drInsertVendor As DataRow
        Dim sendDs As New DataSet
        Dim receiveDs As New DataSet

        'Validations    
        If Len(txtVname.Text) = 0 Then
            MsgBox("Enter Vendor name.      ", MsgBoxStyle.Information)
            Exit Sub
        End If

        If txtemail.Text.IndexOf("@"c) = -1 Or txtemail.Text.IndexOf("."c) = -1 Or txtemail.Text.IndexOf("@"c) <> txtemail.Text.LastIndexOf("@"c) Then
            MessageBox.Show("Enter a valid email address.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Len(txtemail.Text) = 0 Then
            MessageBox.Show("Enter customer email address.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Not Len(txtvat.Text) = 0 Then
            If IsNumeric(txtvat.Text) = False Then
                MessageBox.Show("Enter numeric value in the vat.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        End If

        If IsNumeric(txtclimit.Text) = False Then
            MessageBox.Show("Enter numeric value in the credit limit.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If txtcity.Text.IndexOf("'"c) <> -1 Or txtclimit.Text.IndexOf("'"c) <> -1 Or txtVname.Text.IndexOf("'"c) <> -1 Or txtcountry.Text.IndexOf("'"c) <> -1 Or txtemail.Text.IndexOf("'"c) <> -1 Or txtfax.Text.IndexOf("'"c) <> -1 Or txtpcode.Text.IndexOf("'"c) <> -1 Or txtstreet.Text.IndexOf("'"c) <> -1 Or txttelephone.Text.IndexOf("'"c) <> -1 Or txtvat.Text.IndexOf("'"c) <> -1 Then
            MsgBox("Invalid Character "" ' "" found.     ", MsgBoxStyle.Critical, "Infini Accounts Express")
            Exit Sub
        End If

        btnOk.Enabled = False

        Try
            'Get Customer ID through Sync Table
            dsCustomerId = ClsPurchase.GetSynchronizeTableData()
            dtCustomerId = dsCustomerId.Tables(0)
            For Each drCustomerId In dtCustomerId.Rows
                CustomerId = drCustomerId(2)
            Next
            'CreateTable for online vendor
            VendorTable = CreateTableforOnlineVendor(CustomerId)

            'Send DAtaset for online vendor
            sendDs.Merge(VendorTable)
            'Recevice DAtaset for online vendor
            receiveDs = IEDService.createNewVendor(sendDs)

            'If Not receiveDs.Tables(0).Rows.Count = 0 Then
            'dataset in to datatable
            dtInsertVendor = receiveDs.Tables(0)

            For Each drInsertVendor In dtInsertVendor.Rows
                'Insert customer in MS-Access 
                ClsPurchase.InsertOnlineVendor(drInsertVendor("VendorId"), drInsertVendor("VendorName"), drInsertVendor("Street"), drInsertVendor("Town"), drInsertVendor("Country"), drInsertVendor("PostalCode"), drInsertVendor("ContactName"), drInsertVendor("Email"), drInsertVendor("Telephone"), drInsertVendor("Fax"), drInsertVendor("Vat"), drInsertVendor("CreditLimit"))
            Next
            'Execute command 
            If (GlbGlobal.ExecuteCommand() = True) Then
                MessageBox.Show("Vendor has been created.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Exit Sub
            End If
            ClearTextbox()
            Me.Dispose(True)

            'End If

        Catch ex As Exception
            Throw New Exception("OnlineCreateVendor : " & ex.Message)
        End Try

        btnOk.Enabled = True

    End Sub

    Private Sub txtclimit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtclimit.Leave

        If IsNumeric(txtclimit.Text) = False Then
            txtclimit.Text = "0.00"
        End If
        txtclimit.Text = Format(Val(txtclimit.Text), "0.00")
    End Sub

End Class