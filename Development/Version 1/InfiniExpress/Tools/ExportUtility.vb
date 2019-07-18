Imports System
Imports System.Data.OleDb
Imports InfiniExpress.BLL
Imports vb = Microsoft.VisualBasic.Compatibility.VB6

Public Class ImpExpUtility

    Inherits System.Windows.Forms.Form
    Dim ClsGlobal As New ClassGlobal()
    Dim ClsUtility As New ClassUtility()
    Public Shared dbPath As String

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DirList As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveList As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents FileList As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Protected WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ImpExpUtility))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DirList = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.DriveList = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.FileList = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(322, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Export Data From Infini Express To Infini Accounts Pro"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel1.Location = New System.Drawing.Point(442, 66)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(44, 16)
        Me.LinkLabel1.TabIndex = 1
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Select infini accounts pro database"
        '
        'lblPath
        '
        Me.lblPath.BackColor = System.Drawing.Color.White
        Me.lblPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPath.Location = New System.Drawing.Point(532, 288)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(290, 28)
        Me.lblPath.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Silver
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(440, 288)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "DataBase Path"
        '
        'DirList
        '
        Me.DirList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirList.IntegralHeight = False
        Me.DirList.Location = New System.Drawing.Point(14, 76)
        Me.DirList.Name = "DirList"
        Me.DirList.Size = New System.Drawing.Size(284, 116)
        Me.DirList.TabIndex = 5
        '
        'DriveList
        '
        Me.DriveList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DriveList.Location = New System.Drawing.Point(14, 44)
        Me.DriveList.Name = "DriveList"
        Me.DriveList.Size = New System.Drawing.Size(134, 22)
        Me.DriveList.TabIndex = 6
        '
        'FileList
        '
        Me.FileList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileList.Location = New System.Drawing.Point(14, 200)
        Me.FileList.Name = "FileList"
        Me.FileList.Pattern = "*.MDB"
        Me.FileList.ReadOnly = False
        Me.FileList.Size = New System.Drawing.Size(284, 69)
        Me.FileList.TabIndex = 7
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(10, 40)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(144, 30)
        Me.Panel1.TabIndex = 8
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(10, 72)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(294, 210)
        Me.Panel2.TabIndex = 9
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(420, 292)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(102, 20)
        Me.TextBox1.TabIndex = 10
        Me.TextBox1.Text = ""
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Button1.Location = New System.Drawing.Point(10, 294)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(98, 24)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "&Export Data"
        '
        'ImpExpUtility
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(338, 327)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.TextBox1, Me.FileList, Me.DriveList, Me.DirList, Me.Label2, Me.lblPath, Me.LinkLabel1, Me.Label1, Me.Panel1, Me.Panel2})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ImpExpUtility"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Export Utility"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub DriveList_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DriveList.SelectedValueChanged
        Try
            DirList.Path = DriveList.Drive
        Catch EE As Exception
            MsgBox("Selected drive is empty.    ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End Try
    End Sub

    Private Sub DirList_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DirList.SelectedValueChanged
        FileList.Path = DirList.Path
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim IntVar As Integer

        Dim Str As String = UCase(FileList.FileName)
        If Str <> "ACCEXPPRO.MDB" Then
            MsgBox("Wrong database has been selected. Please select database name " & " AccExoPro.mdb      ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        ElseIf Str = "" Then
            MsgBox("Select database.    ", MsgBoxStyle.Information, "Infini Accounts Express")
            Exit Sub
        End If

        IntVar = MsgBox("This Process will replace your accounts information from infiniaccounts express to infiniaccounts pro." & vbCrLf & "                                                                 Are you sure", MsgBoxStyle.OKCancel)

        If IntVar = 1 Then

            dbPath = DirList.Path & "\" & FileList.FileName

            'DbPath For Accounts_Pro 
            ClsUtility.Dbpath_Pro(dbPath)

            'Delete All "Accounts_Pro" Table
            ClsUtility.DeleteAllTable_Pro()

            'Selecting & Inserting All Customer 
            ClsUtility.AllCustomer()

            'Selecting & Inserting All Vendor
            ClsUtility.AllVendor()

            'Selecting & Inserting Company Information
            ClsUtility.Company()

            'Selecting & Inserting Bank Information
            ClsUtility.Bank()

            'Selecting & Inserting Financial Year Data
            ClsUtility.FinancialYear()

            'Selecting & Updating TaxIDs 
            ClsUtility.UpdateTaxIDs()

            'Selecting & Inserting Transaction Data
            ClsUtility.Transaction()

            'Selecting & Inserting Ledger Data
            ClsUtility.Ledger()

            'Selecting & Inserting Outstanding Data
            ClsUtility.Outstanding()

            'Inserting N_Record Data
            ClsUtility.N_Record()

        MsgBox("All data has been export sucessfully.       ", MsgBoxStyle.Information, "Infini Accounts Express")
        End If
    End Sub

    Private Sub DirList_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DirList.MouseDown
        FileList.Path = DirList.Path
    End Sub


    Private Sub ImpExpUtility_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
