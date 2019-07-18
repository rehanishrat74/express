Imports InfiniExpress.BLL
Public Class CreditCardInfo
    Inherits System.Windows.Forms.Form

    Public ClsSales As New ClassSales

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
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Protected WithEvents Label15 As System.Windows.Forms.Label
    Protected WithEvents Label3 As System.Windows.Forms.Label
    Protected WithEvents Label4 As System.Windows.Forms.Label
    Protected WithEvents Label5 As System.Windows.Forms.Label
    Protected WithEvents Label6 As System.Windows.Forms.Label
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents Label8 As System.Windows.Forms.Label
    Protected WithEvents Label9 As System.Windows.Forms.Label
    Protected WithEvents Label10 As System.Windows.Forms.Label
    Protected WithEvents Label11 As System.Windows.Forms.Label
    Protected WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtshipComp As System.Windows.Forms.TextBox
    Friend WithEvents txtTrackNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtDaddress As System.Windows.Forms.TextBox
    Friend WithEvents txtcardholdername As System.Windows.Forms.TextBox
    Friend WithEvents txtcardnumber As System.Windows.Forms.TextBox
    Friend WithEvents txtedate As System.Windows.Forms.TextBox
    Friend WithEvents txtScode As System.Windows.Forms.TextBox
    Friend WithEvents txtcardaddress As System.Windows.Forms.TextBox
    Protected WithEvents Button4 As System.Windows.Forms.Button
    Protected WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cmdType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents CmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents CmbYear As System.Windows.Forms.ComboBox
    Protected WithEvents Label1 As System.Windows.Forms.Label
    Protected WithEvents Label18 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmbYear = New System.Windows.Forms.ComboBox
        Me.CmbMonth = New System.Windows.Forms.ComboBox
        Me.cmdType = New System.Windows.Forms.ComboBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtcardaddress = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtScode = New System.Windows.Forms.TextBox
        Me.txtedate = New System.Windows.Forms.TextBox
        Me.txtcardnumber = New System.Windows.Forms.TextBox
        Me.txtcardholdername = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDaddress = New System.Windows.Forms.TextBox
        Me.txtTrackNumber = New System.Windows.Forms.TextBox
        Me.txtshipComp = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.CmbYear)
        Me.Panel1.Controls.Add(Me.CmbMonth)
        Me.Panel1.Controls.Add(Me.cmdType)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.btnOk)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.txtcardaddress)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.txtScode)
        Me.Panel1.Controls.Add(Me.txtedate)
        Me.Panel1.Controls.Add(Me.txtcardnumber)
        Me.Panel1.Controls.Add(Me.txtcardholdername)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.txtDaddress)
        Me.Panel1.Controls.Add(Me.txtTrackNumber)
        Me.Panel1.Controls.Add(Me.txtshipComp)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(6, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(420, 340)
        Me.Panel1.TabIndex = 2
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.White
        Me.Label18.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(257, 222)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(85, 14)
        Me.Label18.TabIndex = 65
        Me.Label18.Text = "( MM/YYYYY )"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(175, 223)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(12, 14)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "/"
        '
        'CmbYear
        '
        Me.CmbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbYear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbYear.Items.AddRange(New Object() {"2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015"})
        Me.CmbYear.Location = New System.Drawing.Point(189, 219)
        Me.CmbYear.Name = "CmbYear"
        Me.CmbYear.Size = New System.Drawing.Size(63, 21)
        Me.CmbYear.TabIndex = 7
        '
        'CmbMonth
        '
        Me.CmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbMonth.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbMonth.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.CmbMonth.Location = New System.Drawing.Point(132, 219)
        Me.CmbMonth.Name = "CmbMonth"
        Me.CmbMonth.Size = New System.Drawing.Size(39, 21)
        Me.CmbMonth.TabIndex = 6
        '
        'cmdType
        '
        Me.cmdType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmdType.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdType.Items.AddRange(New Object() {"Visa Card", "Master Card / Euro Card", "Dinner Club / Carte Balance", "Discovery", "enRoute", "JBC", "Amex Card", "Debit Card"})
        Me.cmdType.Location = New System.Drawing.Point(132, 147)
        Me.cmdType.Name = "cmdType"
        Me.cmdType.Size = New System.Drawing.Size(162, 21)
        Me.cmdType.TabIndex = 3
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.White
        Me.Button4.Location = New System.Drawing.Point(221, 312)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(80, 22)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "&Cancel"
        Me.Button4.Visible = False
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(132, 312)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(80, 22)
        Me.btnOk.TabIndex = 9
        Me.btnOk.Text = "&OK"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.White
        Me.Label12.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(12, 12)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 14)
        Me.Label12.TabIndex = 61
        Me.Label12.Text = "Shopping Details"
        '
        'txtcardaddress
        '
        Me.txtcardaddress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcardaddress.Location = New System.Drawing.Point(132, 267)
        Me.txtcardaddress.MaxLength = 255
        Me.txtcardaddress.Multiline = True
        Me.txtcardaddress.Name = "txtcardaddress"
        Me.txtcardaddress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtcardaddress.Size = New System.Drawing.Size(276, 33)
        Me.txtcardaddress.TabIndex = 8
        Me.txtcardaddress.Text = ""
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.White
        Me.Label11.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(12, 273)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(99, 14)
        Me.Label11.TabIndex = 59
        Me.Label11.Text = "Card Address *"
        '
        'txtScode
        '
        Me.txtScode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtScode.Location = New System.Drawing.Point(132, 243)
        Me.txtScode.MaxLength = 50
        Me.txtScode.Name = "txtScode"
        Me.txtScode.Size = New System.Drawing.Size(199, 21)
        Me.txtScode.TabIndex = 7
        Me.txtScode.Text = ""
        '
        'txtedate
        '
        Me.txtedate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtedate.Location = New System.Drawing.Point(327, 309)
        Me.txtedate.MaxLength = 20
        Me.txtedate.Name = "txtedate"
        Me.txtedate.Size = New System.Drawing.Size(85, 21)
        Me.txtedate.TabIndex = 6
        Me.txtedate.Text = ""
        Me.txtedate.Visible = False
        '
        'txtcardnumber
        '
        Me.txtcardnumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcardnumber.Location = New System.Drawing.Point(132, 195)
        Me.txtcardnumber.MaxLength = 16
        Me.txtcardnumber.Name = "txtcardnumber"
        Me.txtcardnumber.Size = New System.Drawing.Size(200, 21)
        Me.txtcardnumber.TabIndex = 5
        Me.txtcardnumber.Text = ""
        '
        'txtcardholdername
        '
        Me.txtcardholdername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcardholdername.Location = New System.Drawing.Point(132, 171)
        Me.txtcardholdername.MaxLength = 200
        Me.txtcardholdername.Name = "txtcardholdername"
        Me.txtcardholdername.Size = New System.Drawing.Size(200, 21)
        Me.txtcardholdername.TabIndex = 4
        Me.txtcardholdername.Text = ""
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.White
        Me.Label10.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(12, 249)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 14)
        Me.Label10.TabIndex = 53
        Me.Label10.Text = "Security Code"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(12, 225)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 14)
        Me.Label9.TabIndex = 52
        Me.Label9.Text = "Expiry Date *"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(12, 201)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 14)
        Me.Label8.TabIndex = 51
        Me.Label8.Text = "Card Number *"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(12, 177)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 14)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Card Holder Name *"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(12, 150)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 14)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Card Type *"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.White
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(12, 129)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(160, 14)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Credit Card Information"
        '
        'txtDaddress
        '
        Me.txtDaddress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDaddress.Location = New System.Drawing.Point(132, 80)
        Me.txtDaddress.MaxLength = 250
        Me.txtDaddress.Multiline = True
        Me.txtDaddress.Name = "txtDaddress"
        Me.txtDaddress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDaddress.Size = New System.Drawing.Size(276, 33)
        Me.txtDaddress.TabIndex = 2
        Me.txtDaddress.Text = ""
        '
        'txtTrackNumber
        '
        Me.txtTrackNumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTrackNumber.Location = New System.Drawing.Point(132, 56)
        Me.txtTrackNumber.MaxLength = 49
        Me.txtTrackNumber.Name = "txtTrackNumber"
        Me.txtTrackNumber.Size = New System.Drawing.Size(200, 21)
        Me.txtTrackNumber.TabIndex = 1
        Me.txtTrackNumber.Text = ""
        '
        'txtshipComp
        '
        Me.txtshipComp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtshipComp.Location = New System.Drawing.Point(132, 32)
        Me.txtshipComp.MaxLength = 200
        Me.txtshipComp.Name = "txtshipComp"
        Me.txtshipComp.Size = New System.Drawing.Size(200, 21)
        Me.txtshipComp.TabIndex = 0
        Me.txtshipComp.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.White
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 14)
        Me.Label4.TabIndex = 44
        Me.Label4.Text = "Delivery Address"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 14)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Tracking Number"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.White
        Me.Label15.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(12, 32)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(120, 14)
        Me.Label15.TabIndex = 42
        Me.Label15.Text = "Shipping Company"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(273, 372)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(93, 21)
        Me.Label16.TabIndex = 69
        Me.Label16.Text = "Label16"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(174, 372)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(93, 21)
        Me.Label14.TabIndex = 68
        Me.Label14.Text = "Label14"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(75, 372)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(93, 21)
        Me.Label13.TabIndex = 67
        Me.Label13.Text = "Label13"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(369, 372)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 21)
        Me.Label2.TabIndex = 70
        Me.Label2.Text = "Label2"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(0, 366)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(93, 21)
        Me.Label17.TabIndex = 71
        Me.Label17.Text = "Label17"
        '
        'CreditCardInfo
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(429, 353)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(250, 12)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CreditCardInfo"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Credit Card Information"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose()
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim tempInvDetail As String = "Invoice From Infini Accounts Express"
        Dim tempInvTp As String = "Invoice", temptp As String
        Dim tempAc As String = "@ACC" & Label17.Text
        Dim SysDT As Date
        Dim DT As String = SysDT.Now
        Dim dStr As String

        dStr = "01/" & CmbMonth.Text & "/" & CmbYear.Text

        If Len(txtcardholdername.Text) = 0 Then
            MsgBox("Please enter card holder name.  ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtcardholdername.Focus()
            Exit Sub
        End If

        If Len(txtcardnumber.Text) = 0 Then
            MsgBox("Please enter card number.  ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtcardnumber.Focus()
            Exit Sub
        End If

        If IsNumeric(txtcardnumber.Text) = False Then
            MessageBox.Show("Enter numeric value in the credit card number.     ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtcardnumber.Focus()
            Exit Sub
        End If

        If Len(txtcardaddress.Text) = 0 Then
            MsgBox("Please enter card addresss.  ", MsgBoxStyle.Information, "Infini Accounts Express")
            txtcardaddress.Focus()
            Exit Sub
        End If

        If Label2.Text = "Credit Card" Then
            temptp = "CC"
        Else
            temptp = "CB"
        End If

        txtedate.Text = dStr

        'Insert Credit Card Information 
        Try
            ClsSales.InsertCardInfo(Label13.Text, Label16.Text, Label14.Text, txtshipComp.Text, txtTrackNumber.Text, txtDaddress.Text, cmdType.SelectedItem, txtcardholdername.Text, txtcardnumber.Text, txtedate.Text, txtScode.Text, txtcardaddress.Text, DT, tempAc, tempInvDetail, tempInvTp, temptp, Label17.Text)
            If ClassGlobal.ExecuteCommand = True Then
                MsgBox("Credit card information has been saved.", MsgBoxStyle.Information)
                Me.Dispose()
            End If
        Catch ex As Exception
            Throw New Exception("Credit Card Info : " & ex.Message)
        End Try

    End Sub

    Private Sub CreditCardInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim tStr As Integer
        Dim ccDate As Date = Now.Date

        tStr = Month(ccDate)
        cmdType.SelectedIndex = 0
        CmbMonth.Text = CStr(tStr)
        tStr = Year(ccDate)
        CmbYear.Text = CStr(tStr)

    End Sub
End Class
