Imports System
Imports System.Data
Imports InfiniExpress.BLL

Public Class FrmLogin

    Inherits System.Windows.Forms.Form
    Dim GlbSynchronize As New ClassSynchronize
    Dim IEDService As New InfiniExpress.InfiniWebService.InfiniService
    Dim UserValid As String

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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPassword As System.Windows.Forms.TextBox
    Friend WithEvents TxtLoginId As System.Windows.Forms.TextBox
    Protected WithEvents CmdAuthenticate As System.Windows.Forms.Button
    Protected WithEvents CmdCancel As System.Windows.Forms.Button
    Friend WithEvents LblErr As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmLogin))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.LblErr = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPassword = New System.Windows.Forms.TextBox
        Me.TxtLoginId = New System.Windows.Forms.TextBox
        Me.CmdAuthenticate = New System.Windows.Forms.Button
        Me.CmdCancel = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.LblErr)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(314, 176)
        Me.Panel1.TabIndex = 20
        '
        'LblErr
        '
        Me.LblErr.ForeColor = System.Drawing.Color.Maroon
        Me.LblErr.Location = New System.Drawing.Point(13, 48)
        Me.LblErr.Name = "LblErr"
        Me.LblErr.Size = New System.Drawing.Size(288, 16)
        Me.LblErr.TabIndex = 4
        Me.LblErr.Text = "Label Error Messages"
        Me.LblErr.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(287, 30)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Please Provide the customer id and password to synchronizes data from the account" & _
        " centre."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TxtPassword)
        Me.GroupBox1.Controls.Add(Me.TxtLoginId)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(9, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(294, 91)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Password :"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Customer ID :"
        '
        'TxtPassword
        '
        Me.TxtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtPassword.Location = New System.Drawing.Point(109, 54)
        Me.TxtPassword.MaxLength = 20
        Me.TxtPassword.Name = "TxtPassword"
        Me.TxtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TxtPassword.Size = New System.Drawing.Size(155, 21)
        Me.TxtPassword.TabIndex = 1
        Me.TxtPassword.Text = ""
        '
        'TxtLoginId
        '
        Me.TxtLoginId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtLoginId.Location = New System.Drawing.Point(109, 22)
        Me.TxtLoginId.MaxLength = 10
        Me.TxtLoginId.Name = "TxtLoginId"
        Me.TxtLoginId.Size = New System.Drawing.Size(107, 21)
        Me.TxtLoginId.TabIndex = 0
        Me.TxtLoginId.Text = ""
        '
        'CmdAuthenticate
        '
        Me.CmdAuthenticate.BackColor = System.Drawing.Color.LightSlateGray
        Me.CmdAuthenticate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CmdAuthenticate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdAuthenticate.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CmdAuthenticate.Location = New System.Drawing.Point(136, 190)
        Me.CmdAuthenticate.Name = "CmdAuthenticate"
        Me.CmdAuthenticate.Size = New System.Drawing.Size(90, 23)
        Me.CmdAuthenticate.TabIndex = 2
        Me.CmdAuthenticate.Text = "&Login"
        '
        'CmdCancel
        '
        Me.CmdCancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.CmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CmdCancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdCancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CmdCancel.Location = New System.Drawing.Point(232, 190)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.Size = New System.Drawing.Size(90, 23)
        Me.CmdCancel.TabIndex = 3
        Me.CmdCancel.Text = "&Cancel"
        '
        'FrmLogin
        '
        Me.AcceptButton = Me.CmdAuthenticate
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(326, 219)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CmdAuthenticate)
        Me.Controls.Add(Me.CmdCancel)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Login to the Online Accounts Centre"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCancel.Click
        Me.Dispose(True)
        'Application.Exit()
    End Sub

    Private Sub CmdAuthenticate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAuthenticate.Click

        Dim UpdateFlag As Boolean
        LblErr.Text = ""

        If Trim(TxtLoginId.Text) <> "" And Trim(TxtPassword.Text) <> "" Then

            Try

                If IsNumeric(TxtLoginId.Text) = False Then
                    MsgBox("Please enter numeric value (0-9) only.      ", MsgBoxStyle.Critical, "Infini Accounts Express")
                    TxtLoginId.Focus()
                    Exit Sub
                End If
                If TxtPassword.Text.IndexOf("'"c) <> -1 Then
                    MsgBox("Invalid Character "" ' "" found.     ", MsgBoxStyle.Critical, "Infini Accounts Express")
                    TxtPassword.Focus()
                    Exit Sub
                End If

                UserValid = IEDService.CustomerIdValidation(CInt(TxtLoginId.Text), CStr(TxtPassword.Text))

                If UserValid = "False" Then
                    MsgBox("Please Provide correct customer ID and password.      ", MsgBoxStyle.Information, "Infini Accounts Express")
                    TxtLoginId.Focus()
                    Exit Sub
                End If

            Catch e1 As Exception
                UserValid = ""
                MsgBox("Login Failed...  " + e1.Message + "", MsgBoxStyle.Critical, "Infini Accounts Express")
                LblErr.Text = "Login Failed......"
            End Try

        Else
            MsgBox("Customer ID and password are required fields.      ", vbInformation, "Infini Accounts Express")
            TxtLoginId.Focus()
            LblErr.Text = "Required customer id and password."
            Exit Sub
        End If

        If UserValid <> "" Then

            Dim IsCheckService As Boolean
            IsCheckService = IEDService.GetServiceInfo(CInt(UserValid))

            If IsCheckService = True Then IsService = True

            UpdateFlag = GlbSynchronize.UpdateUserSynchronizeInfo(CInt(UserValid))

            If UpdateFlag = True Then
                Me.Visible = False
                Application.DoEvents()
                Dim EMDI As New MDI
                EMDI.Show()
            End If

        End If

    End Sub

    Private Sub FrmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Check Running Instance Of This Application
        CheckForExistingInstance()
        LblErr.Text = ""
        TxtLoginId.Focus()

    End Sub

    Public Sub CheckForExistingInstance()
        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MsgBox("Another Instance of this application is already running, Multiple Instances Forbidden.     ", MsgBoxStyle.Information, "Infini Accounts Express")
            Application.Exit()
        End If
    End Sub

    Private Sub TxtLoginId_TextLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLoginId.Leave

        If Trim(TxtLoginId.Text) = "" Then Exit Sub
        If IsNumeric(TxtLoginId.Text) = False Then
            MsgBox("Enter only numeric value.    ", MsgBoxStyle.Critical, "Infini Accounts Express")
            TxtLoginId.Text = ""
        End If

    End Sub

End Class
