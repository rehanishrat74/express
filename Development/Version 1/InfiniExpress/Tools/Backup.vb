Imports System
Imports System.Data
Imports InfiniExpress.BLL
Imports IOsystem = System.IO

Public Class Backup

    Inherits System.Windows.Forms.Form
    Private clsUtility As New ClassUtility()

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
    Friend WithEvents DriveList As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents DirList As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtfname As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Backup))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DriveList = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox()
        Me.DirList = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtfname = New System.Windows.Forms.TextBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(7, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(352, 68)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Allow you to take backup your database in the form of XML file, Now what you have" & _
        " to do is to choose your destination path Or by default it is your application p" & _
        "ath where ur dababase is."
        '
        'DriveList
        '
        Me.DriveList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DriveList.Location = New System.Drawing.Point(6, 28)
        Me.DriveList.Name = "DriveList"
        Me.DriveList.Size = New System.Drawing.Size(154, 22)
        Me.DriveList.TabIndex = 8
        '
        'DirList
        '
        Me.DirList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirList.IntegralHeight = False
        Me.DirList.Location = New System.Drawing.Point(6, 52)
        Me.DirList.Name = "DirList"
        Me.DirList.Size = New System.Drawing.Size(284, 104)
        Me.DirList.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Black
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(6, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 16)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Location For Backup"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.Label3, Me.txtfname, Me.Label2, Me.DirList, Me.DriveList})
        Me.Panel1.Location = New System.Drawing.Point(9, 76)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(360, 214)
        Me.Panel1.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Red
        Me.Label4.Location = New System.Drawing.Point(168, 182)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(172, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Folder name max 15chr long"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Black
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(6, 162)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(160, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Destination Folder Name"
        '
        'txtfname
        '
        Me.txtfname.Location = New System.Drawing.Point(6, 180)
        Me.txtfname.MaxLength = 15
        Me.txtfname.Name = "txtfname"
        Me.txtfname.Size = New System.Drawing.Size(160, 20)
        Me.txtfname.TabIndex = 10
        Me.txtfname.Text = ""
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.ForeColor = System.Drawing.Color.White
        Me.btnOk.Location = New System.Drawing.Point(214, 300)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(74, 18)
        Me.btnOk.TabIndex = 60
        Me.btnOk.Text = "&OK"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(294, 300)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(74, 18)
        Me.Button1.TabIndex = 61
        Me.Button1.Text = "&Cancel"
        '
        'Backup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(376, 325)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.btnOk, Me.Label1, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Backup"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Backup"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub DriveList_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DriveList.SelectedValueChanged
        Try
            DirList.Path = DriveList.Drive
        Catch EE As Exception
            MsgBox("Selected drive is empty")
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Dim response As Integer
        Dim strPath As String
        Dim strFileName As String
        Dim Rvalue As Integer

        If Len(Trim(txtfname.Text)) = 0 Then
            strPath = DirList.Path & "\"
            If IOsystem.File.Exists(strPath & "\ExpressBackUp.XML") = True Then
                response = MsgBox("Backup file already exist in this folder,would you like to overwrite them.", MsgBoxStyle.YesNo)

                If response = vbYes Then
                    IOsystem.File.Delete(strPath & "\ExpressBackUp.XML")
                    Rvalue = clsUtility.WriteXML4BackUp(strPath)
                Else
                    Exit Sub
                End If

            Else
                Rvalue = clsUtility.WriteXML4BackUp(strPath)
            End If
        Else
            strPath = DirList.Path & "\" & Trim(txtfname.Text) & "\"
            IOsystem.Directory.CreateDirectory(strPath)
            Rvalue = clsUtility.WriteXML4BackUp(strPath)
        End If

        txtfname.Text = ""
        If Rvalue = 0 Then
            MsgBox("Backup has been made in the form of XML file.         ", MsgBoxStyle.Information, "Infini Accounts Express")
        End If

    End Sub

End Class
