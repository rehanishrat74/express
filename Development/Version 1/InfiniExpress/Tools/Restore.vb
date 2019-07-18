Imports System
Imports System.Data
Imports InfiniExpress.BLL
Imports IOsystem = System.IO

Public Class Restore

    Inherits System.Windows.Forms.Form
    Dim str As String
    Dim clsutility As New ClassUtility()

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DirList As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveList As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Restore))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DirList = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.DriveList = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.DirList, Me.DriveList})
        Me.Panel1.Location = New System.Drawing.Point(8, 165)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(360, 190)
        Me.Panel1.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Black
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(6, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(190, 16)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Location Of Your Backup Files"
        '
        'DirList
        '
        Me.DirList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirList.IntegralHeight = False
        Me.DirList.Location = New System.Drawing.Point(6, 52)
        Me.DirList.Name = "DirList"
        Me.DirList.Size = New System.Drawing.Size(348, 130)
        Me.DirList.TabIndex = 7
        '
        'DriveList
        '
        Me.DriveList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DriveList.Location = New System.Drawing.Point(6, 28)
        Me.DriveList.Name = "DriveList"
        Me.DriveList.Size = New System.Drawing.Size(192, 22)
        Me.DriveList.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(8, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(354, 44)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "This is a dangerous operation. The system may fail to work in case of data loss. " & _
        "You are requried to make a backup first. "
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.Label5.Font = New System.Drawing.Font("Verdana", 9.0!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "CAUTION !"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(8, 84)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(354, 76)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "It is advise to take the backup of your current data. The restore process will wi" & _
        "peout the data from your system and will replace with the new one. So if you are" & _
        " not sure then take backup of your existing data files first."
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(293, 365)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(74, 18)
        Me.Button1.TabIndex = 63
        Me.Button1.Text = "&Cancel"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(213, 365)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(74, 18)
        Me.btnOk.TabIndex = 62
        Me.btnOk.Text = "&OK"
        '
        'Restore
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(376, 389)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.btnOk, Me.Label6, Me.Label5, Me.Label1, Me.Panel1})
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Restore"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Restore"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub DriveList_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DriveList.SelectedValueChanged
        Try
            DirList.Path = DriveList.Drive
            str = DirList.Path
        Catch EE As Exception
            MsgBox("Selected drive is empty")
            Exit Sub
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        Dim ds As New DataSet()
        Dim myRow As DataRow
        Dim myCol As DataColumn

        'Check Existing Files
        If IOsystem.File.Exists(DirList.Path & "\ExpressBackUp.XML") = False Then
            MsgBox("Select backup file folder. Where your backup files is.")
            Exit Sub
        End If

        'Delete All From AccountsExpress
        clsutility.DeleteAllTable_Exp()

        'Insert Express Customer
        clsutility.EAllCustomer(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express Vendor
        clsutility.EAllVendor(DirList.Path & "\ExpressBackUp.XML")

        'Update Express Company
        clsutility.ECompany(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express Bank
        clsutility.EBank(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express FinancialYear
        clsutility.EFinancialYear(DirList.Path & "\ExpressBackUp.XML")

        'Update Express TaxIDs
        clsutility.EUpdateTaxIDs(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express Outstanding
        clsutility.EOutstanding(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express Transaction
        clsutility.ETransaction(DirList.Path & "\ExpressBackUp.XML")

        'Insert Express Ledger
        clsutility.ELedger(DirList.Path & "\ExpressBackUp.XMl")

        'Insert Express Creditcard Info
        'clsutility.ECREDITCARDINFO(DirList.Path & "\ExpressBackUp.XMl")


        MsgBox("All data has been restore secessfully.      ", MsgBoxStyle.Information, "Infini Accounts Express")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

End Class
