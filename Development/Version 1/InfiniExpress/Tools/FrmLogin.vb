Imports System
Imports System.Data
Imports InfiniExpress.BLL
Imports System.Threading

Public Class FrmLogin

    Inherits System.Windows.Forms.Form
    Dim PBar As New ProgressBar()
    Dim GlbSynchronize As New ClassSynchronize()
    Dim IEDService As New InfiniExpress.localhost.InfiniService()
    Dim UserValid As String

    Private Class Worker
        Public Event ProgressUpdate(ByVal Percent As Integer)

        Protected Sub OnProgressUpdate(ByVal _Percent As Integer)
            RaiseEvent ProgressUpdate(_Percent)
        End Sub

        Public Sub DoWork()
            Dim i As Integer
            For i = 0 To 99
                Me.OnProgressUpdate(i)
                Thread.Sleep(50)
            Next
        End Sub

        'Public Sub DoWork2()
        '    Dim i As Integer
        '    For i = 90 To 99
        '        Me.OnProgressUpdate(i)
        '        Thread.Sleep(100)
        '    Next
        'End Sub

        Public Sub Start()
            Dim t As New Thread(AddressOf Me.DoWork)
            t.Start()
        End Sub

    End Class

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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPassword As System.Windows.Forms.TextBox
    Friend WithEvents TxtLoginId As System.Windows.Forms.TextBox
    Protected WithEvents CmdAuthenticate As System.Windows.Forms.Button
    Protected WithEvents CmdCancel As System.Windows.Forms.Button
    Friend WithEvents SBar As System.Windows.Forms.StatusBar
    Friend WithEvents LblErr As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmLogin))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.LblErr = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtPassword = New System.Windows.Forms.TextBox()
        Me.TxtLoginId = New System.Windows.Forms.TextBox()
        Me.CmdAuthenticate = New System.Windows.Forms.Button()
        Me.CmdCancel = New System.Windows.Forms.Button()
        Me.SBar = New System.Windows.Forms.StatusBar()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.LblErr, Me.Label3, Me.GroupBox1})
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
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Label1, Me.TxtPassword, Me.TxtLoginId})
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
        Me.Label2.Location = New System.Drawing.Point(13, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Password :"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 26)
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
        Me.TxtPassword.Size = New System.Drawing.Size(168, 21)
        Me.TxtPassword.TabIndex = 1
        Me.TxtPassword.Text = "asifali"
        '
        'TxtLoginId
        '
        Me.TxtLoginId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtLoginId.Location = New System.Drawing.Point(109, 22)
        Me.TxtLoginId.MaxLength = 10
        Me.TxtLoginId.Name = "TxtLoginId"
        Me.TxtLoginId.Size = New System.Drawing.Size(168, 21)
        Me.TxtLoginId.TabIndex = 0
        Me.TxtLoginId.Text = "64072"
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
        'SBar
        '
        Me.SBar.Location = New System.Drawing.Point(0, 223)
        Me.SBar.Name = "SBar"
        Me.SBar.ShowPanels = True
        Me.SBar.Size = New System.Drawing.Size(330, 16)
        Me.SBar.TabIndex = 21
        Me.SBar.Text = "SBar"
        '
        'FrmLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(330, 239)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.SBar, Me.Panel1, Me.CmdAuthenticate, Me.CmdCancel})
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = " Login to the Online Accounts Centre"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCancel.Click
        Me.Dispose(True)
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

            UpdateFlag = GlbSynchronize.UpdateUserSynchronizeInfo(CInt(UserValid))
            If UpdateFlag = True Then
                SynchronizeData()
            End If

        End If

    End Sub

    Private Sub SynchronizeData()

        Dim dtds As DataSet
        Dim wds As DataSet
        Dim SynFlag As Boolean
        Dim worker As New Worker()

        Try

            AddHandler worker.ProgressUpdate, AddressOf Me.HandleProgress
            worker.Start()

            ''Get un Synchronize Data from DeskTop Application
            dtds = GlbSynchronize.GetUnSynchronizeDeskTopData()
            dtds.WriteXml("D:\Express.xml")
            ''GlbSynchronize.DataManipulation(dtds)

            Dim myTable As DataTable
            Dim myRow As DataRow
            'Dim myColumn As DataColumn
            Dim th As String, TblFlag As Boolean

            For Each myTable In dtds.Tables

                For Each myRow In myTable.Rows

                    th = myTable.TableName
                    If th <> "TblSynchronize" Then
                        TblFlag = True
                        Exit For
                    End If

                Next myRow

            Next myTable

            If TblFlag = True Then
                ''Send Data desktop to Web for Synchronize
                SynFlag = IEDService.SynchronizeFromWeb(dtds)  'For Manipulation

                If SynFlag = True Then
                    dtds.Dispose() 'Reinitialaize
                    'Update Syn Flag = Y for DestTop
                    SynFlag = GlbSynchronize.UpdateSynFlagDeskTopData
                End If

            End If

            'If SynFlag = True Then
            'dtds.Dispose() 'Reinitialaize
            'SynFlag = GlbSynchronize.UpdateSynFlagDeskTopData 'Update Syn Flag = Y for DestTop

            'Get DateTime For Web Synchronization
            wds = GlbSynchronize.GetSynchronizeTableData
            ''Get un Synchronize Data from Web Service
            dtds = IEDService.SynchronizeFromDeskTop(wds)

            SynFlag = GlbSynchronize.DeskTopSynchronizeData(dtds) 'Manipulation Data in DeskTop Application

            If SynFlag = True Then
                GlbSynchronize.UpdateAfterValidation() 'Update Synchronize DateTime
                'Update Syn Flag = Y for Web
                SynFlag = IEDService.UpdateSynchronizeFlag(wds)
            End If

            'worker.DoWork2() 'Progress Bar End
            dtds = Nothing
            wds = Nothing
            MsgBox("Synchronization has been completed.", MsgBoxStyle.Information, "Infini Accounts Express")
            Me.Dispose(True)
            'IEDService = New InfiniExpress.localhost.InfiniService()
            'Else
            'IEDService = New InfiniExpress.localhost.InfiniService()
            'LblErr.Text = "Data transfer failed...."
            'MsgBox("Data transfer failed.     ", MsgBoxStyle.Critical, "Infini Accounts Express")
            'Me.Dispose(True)
            'End If

        Catch e1 As Exception
            dtds = Nothing
            wds = Nothing
            IEDService = New InfiniExpress.localhost.InfiniService()
            LblErr.Text = "Data transfer failed...."
            MsgBox("Data transfer failed...  " + e1.Message + "", MsgBoxStyle.Critical, "Infini Accounts Express")
            Me.Dispose(True)
        End Try

    End Sub

    Private Sub HandleProgress(ByVal percent As Integer)
        Me.PBar.Value = percent
    End Sub

    Private Sub FrmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        SBar.Controls.Add(PBar)
        LblErr.Text = ""

    End Sub

    Private Sub TxtLoginId_TextLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLoginId.Leave

        If Trim(TxtLoginId.Text) = "" Then Exit Sub
        If IsNumeric(TxtLoginId.Text) = False Then
            MsgBox("Enter only numeric value.    ", MsgBoxStyle.Critical, "Infini Accounts Express")
            TxtLoginId.Text = ""
        End If

    End Sub
End Class
