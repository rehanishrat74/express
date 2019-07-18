Imports System
Imports System.Data
Imports System.IO
Imports InfiniExpress.BLL

Public Class CompanyPreferences

    Inherits System.Windows.Forms.Form
    Dim GlbAccount As New ClassAccount()
    Dim GlbGlobal As New ClassGlobal()

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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtTelephone As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnRemoveLogo As System.Windows.Forms.Button
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtCompanyLogo As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtFax As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailAddres As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtVatRegNo As System.Windows.Forms.TextBox
    Friend WithEvents txtCountry As System.Windows.Forms.TextBox
    Friend WithEvents txtCityTown As System.Windows.Forms.TextBox
    Friend WithEvents txtStreet As System.Windows.Forms.TextBox
    Friend WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CompanyPreferences))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtTelephone = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.btnRemoveLogo = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtCompanyLogo = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.txtEmailAddres = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtVatRegNo = New System.Windows.Forms.TextBox()
        Me.txtCountry = New System.Windows.Forms.TextBox()
        Me.txtCityTown = New System.Windows.Forms.TextBox()
        Me.txtStreet = New System.Windows.Forms.TextBox()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnClose, Me.txtTelephone, Me.Label12, Me.btnRemoveLogo, Me.btnBrowse, Me.txtPostalCode, Me.Label10, Me.txtCompanyLogo, Me.Label8, Me.txtFax, Me.txtEmailAddres, Me.Label4, Me.Label3, Me.btnSave, Me.txtVatRegNo, Me.txtCountry, Me.txtCityTown, Me.txtStreet, Me.txtCompanyName, Me.Label9, Me.Label7, Me.Label6, Me.Label5, Me.Label2})
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(448, 436)
        Me.Panel1.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(236, 404)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(60, 18)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "&Close"
        '
        'txtTelephone
        '
        Me.txtTelephone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTelephone.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTelephone.Location = New System.Drawing.Point(136, 225)
        Me.txtTelephone.MaxLength = 80
        Me.txtTelephone.Name = "txtTelephone"
        Me.txtTelephone.Size = New System.Drawing.Size(280, 21)
        Me.txtTelephone.TabIndex = 5
        Me.txtTelephone.Text = ""
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.White
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label12.Location = New System.Drawing.Point(16, 225)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 20)
        Me.Label12.TabIndex = 112
        Me.Label12.Text = "TelePhone"
        '
        'btnRemoveLogo
        '
        Me.btnRemoveLogo.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnRemoveLogo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRemoveLogo.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveLogo.ForeColor = System.Drawing.Color.White
        Me.btnRemoveLogo.Location = New System.Drawing.Point(346, 353)
        Me.btnRemoveLogo.Name = "btnRemoveLogo"
        Me.btnRemoveLogo.Size = New System.Drawing.Size(92, 18)
        Me.btnRemoveLogo.TabIndex = 11
        Me.btnRemoveLogo.Text = "&Remove Logo"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnBrowse.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.ForeColor = System.Drawing.Color.White
        Me.btnBrowse.Location = New System.Drawing.Point(284, 353)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(60, 18)
        Me.btnBrowse.TabIndex = 10
        Me.btnBrowse.Text = "&Browse"
        '
        'txtPostalCode
        '
        Me.txtPostalCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPostalCode.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalCode.Location = New System.Drawing.Point(136, 161)
        Me.txtPostalCode.MaxLength = 50
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New System.Drawing.Size(280, 21)
        Me.txtPostalCode.TabIndex = 3
        Me.txtPostalCode.Text = ""
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.White
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 161)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(100, 20)
        Me.Label10.TabIndex = 111
        Me.Label10.Text = "Postal Code"
        '
        'txtCompanyLogo
        '
        Me.txtCompanyLogo.BackColor = System.Drawing.Color.White
        Me.txtCompanyLogo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCompanyLogo.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyLogo.Location = New System.Drawing.Point(136, 353)
        Me.txtCompanyLogo.MaxLength = 255
        Me.txtCompanyLogo.Name = "txtCompanyLogo"
        Me.txtCompanyLogo.ReadOnly = True
        Me.txtCompanyLogo.Size = New System.Drawing.Size(144, 21)
        Me.txtCompanyLogo.TabIndex = 9
        Me.txtCompanyLogo.Text = ""
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label8.Location = New System.Drawing.Point(16, 353)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 20)
        Me.Label8.TabIndex = 109
        Me.Label8.Text = "Company Logo"
        '
        'txtFax
        '
        Me.txtFax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFax.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFax.Location = New System.Drawing.Point(136, 257)
        Me.txtFax.MaxLength = 80
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(280, 21)
        Me.txtFax.TabIndex = 6
        Me.txtFax.Text = ""
        '
        'txtEmailAddres
        '
        Me.txtEmailAddres.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmailAddres.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmailAddres.Location = New System.Drawing.Point(136, 289)
        Me.txtEmailAddres.MaxLength = 50
        Me.txtEmailAddres.Name = "txtEmailAddres"
        Me.txtEmailAddres.Size = New System.Drawing.Size(280, 21)
        Me.txtEmailAddres.TabIndex = 7
        Me.txtEmailAddres.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 289)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 20)
        Me.Label4.TabIndex = 106
        Me.Label4.Text = "Email Address *"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 257)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 20)
        Me.Label3.TabIndex = 104
        Me.Label3.Text = "Fax"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(164, 404)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(60, 18)
        Me.btnSave.TabIndex = 12
        Me.btnSave.Text = "&Save"
        '
        'txtVatRegNo
        '
        Me.txtVatRegNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVatRegNo.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVatRegNo.Location = New System.Drawing.Point(136, 321)
        Me.txtVatRegNo.MaxLength = 20
        Me.txtVatRegNo.Name = "txtVatRegNo"
        Me.txtVatRegNo.Size = New System.Drawing.Size(280, 21)
        Me.txtVatRegNo.TabIndex = 8
        Me.txtVatRegNo.Text = ""
        '
        'txtCountry
        '
        Me.txtCountry.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCountry.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCountry.Location = New System.Drawing.Point(136, 193)
        Me.txtCountry.MaxLength = 50
        Me.txtCountry.Name = "txtCountry"
        Me.txtCountry.Size = New System.Drawing.Size(280, 21)
        Me.txtCountry.TabIndex = 4
        Me.txtCountry.Text = ""
        '
        'txtCityTown
        '
        Me.txtCityTown.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCityTown.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCityTown.Location = New System.Drawing.Point(136, 129)
        Me.txtCityTown.MaxLength = 50
        Me.txtCityTown.Name = "txtCityTown"
        Me.txtCityTown.Size = New System.Drawing.Size(280, 21)
        Me.txtCityTown.TabIndex = 2
        Me.txtCityTown.Text = ""
        '
        'txtStreet
        '
        Me.txtStreet.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtStreet.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStreet.Location = New System.Drawing.Point(136, 49)
        Me.txtStreet.MaxLength = 255
        Me.txtStreet.Multiline = True
        Me.txtStreet.Name = "txtStreet"
        Me.txtStreet.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStreet.Size = New System.Drawing.Size(280, 72)
        Me.txtStreet.TabIndex = 1
        Me.txtStreet.Text = ""
        '
        'txtCompanyName
        '
        Me.txtCompanyName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCompanyName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyName.Location = New System.Drawing.Point(136, 17)
        Me.txtCompanyName.MaxLength = 255
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(280, 21)
        Me.txtCompanyName.TabIndex = 0
        Me.txtCompanyName.Text = ""
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 321)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(100, 20)
        Me.Label9.TabIndex = 93
        Me.Label9.Text = "VAT Reg No."
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 193)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 20)
        Me.Label7.TabIndex = 92
        Me.Label7.Text = "Country"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 129)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 20)
        Me.Label6.TabIndex = 91
        Me.Label6.Text = "City/Town"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 49)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 20)
        Me.Label5.TabIndex = 90
        Me.Label5.Text = "Street"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 20)
        Me.Label2.TabIndex = 89
        Me.Label2.Text = "Company Name *"
        '
        'CompanyPreferences
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(456, 445)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CompanyPreferences"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Company Preferences"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CompanyPreferences_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim LogoName As String
        Dim Rd As IDataReader = GlbAccount.CompanyInfo
        While Rd.Read()
            If IsDBNull(Rd("Name")) = False Then txtCompanyName.Text = Rd("Name")
            If IsDBNull(Rd("Street")) = False Then txtStreet.Text = Rd("Street")
            If IsDBNull(Rd("City")) = False Then txtCityTown.Text = Rd("City")
            If IsDBNull(Rd("PostalCode")) = False Then txtPostalCode.Text = Rd("PostalCode")
            If IsDBNull(Rd("Country")) = False Then txtCountry.Text = Rd("Country")
            If IsDBNull(Rd("Telephone")) = False Then txtTelephone.Text = Rd("Telephone")
            If IsDBNull(Rd("Fax")) = False Then txtFax.Text = Rd("Fax")
            If IsDBNull(Rd("Email")) = False Then txtEmailAddres.Text = Rd("Email")
            If IsDBNull(Rd("VatNo")) = False Then txtVatRegNo.Text = Rd("VATNo")
            If IsDBNull(Rd("LogoName")) = False Then LogoName = Rd("LogoName")
        End While
        Rd.Close()
        txtCompanyLogo.Text = LogoName
        If LogoName = "" Then
            btnRemoveLogo.Enabled = False
        End If

    End Sub

    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim Logo As Byte(), LogoName As String
        Dim Checked As Boolean

        If txtCompanyName.Text.Length = 0 Then
            MessageBox.Show("Company name cannot be null.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If txtEmailAddres.Text.Length = 0 Or txtEmailAddres.Text.IndexOf("@") = -1 Or txtEmailAddres.Text.IndexOf(".") = -1 Then
            MessageBox.Show("Invalid email address.       ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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

        LogoName = Trim(txtCompanyLogo.Text)

        If LogoName <> "" Then
            Try
                Dim objFileInfo As System.IO.FileInfo = New System.IO.FileInfo(LogoName)
                Dim objFS As New FileStream(LogoName, FileMode.Open, FileAccess.Read)
                Dim objBR As New BinaryReader(objFS)
                Logo = objBR.ReadBytes(objFileInfo.Length)
                objBR.Close()
                objFS.Close()
                objFileInfo = Nothing
                objBR = Nothing
                objFS = Nothing
                Checked = True
            Catch objfileinfo As Exception
                Checked = False
                LogoName = ""
                txtCompanyLogo.Text = ""
                objFileInfo = Nothing
            End Try
        Else
            Checked = False
        End If

        GlbAccount.UpdateCompanyInfo(txtCompanyName.Text, txtStreet.Text, _
               txtCityTown.Text, txtPostalCode.Text, txtCountry.Text, txtTelephone.Text, _
               txtFax.Text, txtEmailAddres.Text, txtVatRegNo.Text, Logo, LogoName, Checked)

        If GlbGlobal.ExecuteCommand() = True Then
            MessageBox.Show("Company detail has been updated.       ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Exit Sub
        End If
        Me.Dispose(True)

    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnBrowse_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim fdlg As New OpenFileDialog()
        fdlg.Title = "Infini Accounts Express Company Logo"
        fdlg.InitialDirectory = "C:\"
        fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fdlg.FilterIndex = 2
        fdlg.RestoreDirectory = True
        If fdlg.ShowDialog() = DialogResult.OK Then
            txtCompanyLogo.Text = fdlg.FileName.ToString()
            btnRemoveLogo.Enabled = True
        End If
        Invalidate()
    End Sub

    Private Sub btnRemoveLogo_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveLogo.Click
        GlbAccount.RemoveCompanyLogo()
        If GlbGlobal.ExecuteCommand() = True Then
            MessageBox.Show("Company logo has been removed.      ", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        btnRemoveLogo.Enabled = False
    End Sub

End Class
