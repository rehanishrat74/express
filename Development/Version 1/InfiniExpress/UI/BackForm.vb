Imports InfiniExpress.BLL

Public Class BackForm

    Inherits System.Windows.Forms.Form
    Private ClsReports As New ClassReports()
    Dim Crow As Integer, Ccol As Integer

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
    Friend WithEvents Sales As System.Windows.Forms.Panel
    Friend WithEvents PictureBox12 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PSales As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Purchase As System.Windows.Forms.Panel
    Friend WithEvents PictureBox13 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox14 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox16 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox17 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox18 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox19 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox20 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox21 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox22 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox23 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox24 As System.Windows.Forms.PictureBox
    Friend WithEvents Account As System.Windows.Forms.Panel
    Friend WithEvents PictureBox25 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox27 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox30 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox31 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox33 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox35 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox36 As System.Windows.Forms.PictureBox
    Friend WithEvents Reports As System.Windows.Forms.Panel
    Friend WithEvents PictureBox26 As System.Windows.Forms.PictureBox
    Friend WithEvents arrow1 As System.Windows.Forms.PictureBox
    Friend WithEvents arrow2 As System.Windows.Forms.PictureBox
    Friend WithEvents arrow3 As System.Windows.Forms.PictureBox
    Friend WithEvents arrow4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox28 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox29 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox32 As System.Windows.Forms.PictureBox
    Friend WithEvents SalesReport As System.Windows.Forms.Panel
    Friend WithEvents PictureBox34 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox37 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox38 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox39 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox40 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox41 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox42 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox43 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox44 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox45 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox48 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox49 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox50 As System.Windows.Forms.PictureBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGrid3 As System.Windows.Forms.DataGrid
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents PictureBox46 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents Tool As System.Windows.Forms.Panel
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox51 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox56 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox58 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox59 As System.Windows.Forms.PictureBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents arrow5 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BackForm))
        Me.Sales = New System.Windows.Forms.Panel
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox46 = New System.Windows.Forms.PictureBox
        Me.PSales = New System.Windows.Forms.PictureBox
        Me.PictureBox12 = New System.Windows.Forms.PictureBox
        Me.PictureBox11 = New System.Windows.Forms.PictureBox
        Me.PictureBox10 = New System.Windows.Forms.PictureBox
        Me.PictureBox9 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Purchase = New System.Windows.Forms.Panel
        Me.PictureBox13 = New System.Windows.Forms.PictureBox
        Me.PictureBox14 = New System.Windows.Forms.PictureBox
        Me.PictureBox15 = New System.Windows.Forms.PictureBox
        Me.PictureBox16 = New System.Windows.Forms.PictureBox
        Me.PictureBox17 = New System.Windows.Forms.PictureBox
        Me.PictureBox18 = New System.Windows.Forms.PictureBox
        Me.PictureBox19 = New System.Windows.Forms.PictureBox
        Me.PictureBox20 = New System.Windows.Forms.PictureBox
        Me.PictureBox21 = New System.Windows.Forms.PictureBox
        Me.PictureBox22 = New System.Windows.Forms.PictureBox
        Me.PictureBox23 = New System.Windows.Forms.PictureBox
        Me.PictureBox24 = New System.Windows.Forms.PictureBox
        Me.Account = New System.Windows.Forms.Panel
        Me.PictureBox25 = New System.Windows.Forms.PictureBox
        Me.PictureBox27 = New System.Windows.Forms.PictureBox
        Me.PictureBox30 = New System.Windows.Forms.PictureBox
        Me.PictureBox31 = New System.Windows.Forms.PictureBox
        Me.PictureBox33 = New System.Windows.Forms.PictureBox
        Me.PictureBox35 = New System.Windows.Forms.PictureBox
        Me.PictureBox36 = New System.Windows.Forms.PictureBox
        Me.Reports = New System.Windows.Forms.Panel
        Me.PictureBox32 = New System.Windows.Forms.PictureBox
        Me.PictureBox29 = New System.Windows.Forms.PictureBox
        Me.PictureBox28 = New System.Windows.Forms.PictureBox
        Me.PictureBox26 = New System.Windows.Forms.PictureBox
        Me.arrow1 = New System.Windows.Forms.PictureBox
        Me.arrow2 = New System.Windows.Forms.PictureBox
        Me.arrow3 = New System.Windows.Forms.PictureBox
        Me.arrow4 = New System.Windows.Forms.PictureBox
        Me.SalesReport = New System.Windows.Forms.Panel
        Me.Button6 = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.PictureBox40 = New System.Windows.Forms.PictureBox
        Me.PictureBox34 = New System.Windows.Forms.PictureBox
        Me.PictureBox37 = New System.Windows.Forms.PictureBox
        Me.PictureBox38 = New System.Windows.Forms.PictureBox
        Me.PictureBox39 = New System.Windows.Forms.PictureBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Button7 = New System.Windows.Forms.Button
        Me.DataGrid3 = New System.Windows.Forms.DataGrid
        Me.PictureBox41 = New System.Windows.Forms.PictureBox
        Me.PictureBox42 = New System.Windows.Forms.PictureBox
        Me.PictureBox43 = New System.Windows.Forms.PictureBox
        Me.PictureBox44 = New System.Windows.Forms.PictureBox
        Me.PictureBox45 = New System.Windows.Forms.PictureBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Button5 = New System.Windows.Forms.Button
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.PictureBox48 = New System.Windows.Forms.PictureBox
        Me.PictureBox49 = New System.Windows.Forms.PictureBox
        Me.PictureBox50 = New System.Windows.Forms.PictureBox
        Me.Tool = New System.Windows.Forms.Panel
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.PictureBox51 = New System.Windows.Forms.PictureBox
        Me.PictureBox56 = New System.Windows.Forms.PictureBox
        Me.PictureBox58 = New System.Windows.Forms.PictureBox
        Me.PictureBox59 = New System.Windows.Forms.PictureBox
        Me.arrow5 = New System.Windows.Forms.PictureBox
        Me.Sales.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Purchase.SuspendLayout()
        Me.Account.SuspendLayout()
        Me.Reports.SuspendLayout()
        Me.SalesReport.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Tool.SuspendLayout()
        Me.SuspendLayout()
        '
        'Sales
        '
        Me.Sales.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Sales.Controls.Add(Me.PictureBox6)
        Me.Sales.Controls.Add(Me.PictureBox1)
        Me.Sales.Controls.Add(Me.PictureBox46)
        Me.Sales.Controls.Add(Me.PSales)
        Me.Sales.Controls.Add(Me.PictureBox12)
        Me.Sales.Controls.Add(Me.PictureBox11)
        Me.Sales.Controls.Add(Me.PictureBox10)
        Me.Sales.Controls.Add(Me.PictureBox9)
        Me.Sales.Controls.Add(Me.PictureBox5)
        Me.Sales.Controls.Add(Me.PictureBox4)
        Me.Sales.Controls.Add(Me.PictureBox3)
        Me.Sales.Controls.Add(Me.PictureBox2)
        Me.Sales.Location = New System.Drawing.Point(6, 4)
        Me.Sales.Name = "Sales"
        Me.Sales.Size = New System.Drawing.Size(544, 382)
        Me.Sales.TabIndex = 39
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(123, 190)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(297, 64)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox6.TabIndex = 58
        Me.PictureBox6.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(6, 48)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 86)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'PictureBox46
        '
        Me.PictureBox46.Image = CType(resources.GetObject("PictureBox46.Image"), System.Drawing.Image)
        Me.PictureBox46.Location = New System.Drawing.Point(74, 79)
        Me.PictureBox46.Name = "PictureBox46"
        Me.PictureBox46.Size = New System.Drawing.Size(148, 33)
        Me.PictureBox46.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox46.TabIndex = 52
        Me.PictureBox46.TabStop = False
        '
        'PSales
        '
        Me.PSales.Image = CType(resources.GetObject("PSales.Image"), System.Drawing.Image)
        Me.PSales.Location = New System.Drawing.Point(-2, 2)
        Me.PSales.Name = "PSales"
        Me.PSales.Size = New System.Drawing.Size(546, 30)
        Me.PSales.TabIndex = 51
        Me.PSales.TabStop = False
        '
        'PictureBox12
        '
        Me.PictureBox12.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox12.Image = CType(resources.GetObject("PictureBox12.Image"), System.Drawing.Image)
        Me.PictureBox12.Location = New System.Drawing.Point(435, 280)
        Me.PictureBox12.Name = "PictureBox12"
        Me.PictureBox12.Size = New System.Drawing.Size(104, 86)
        Me.PictureBox12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox12.TabIndex = 50
        Me.PictureBox12.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(228, 268)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(104, 86)
        Me.PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox11.TabIndex = 49
        Me.PictureBox11.TabStop = False
        '
        'PictureBox10
        '
        Me.PictureBox10.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(8, 280)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(108, 86)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox10.TabIndex = 48
        Me.PictureBox10.TabStop = False
        '
        'PictureBox9
        '
        Me.PictureBox9.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(423, 162)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(114, 86)
        Me.PictureBox9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox9.TabIndex = 47
        Me.PictureBox9.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(2, 160)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(104, 86)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox5.TabIndex = 43
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(297, 80)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(132, 32)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox4.TabIndex = 42
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(188, 46)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(122, 90)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox3.TabIndex = 41
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(433, 44)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 92)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 40
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Button8)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Location = New System.Drawing.Point(6, 412)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(544, 38)
        Me.Panel1.TabIndex = 40
        '
        'Button8
        '
        Me.Button8.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button8.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button8.Image = CType(resources.GetObject("Button8.Image"), System.Drawing.Image)
        Me.Button8.Location = New System.Drawing.Point(328, 6)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(100, 24)
        Me.Button8.TabIndex = 4
        '
        'Button4
        '
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Location = New System.Drawing.Point(436, 6)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 24)
        Me.Button4.TabIndex = 3
        '
        'Button3
        '
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Location = New System.Drawing.Point(220, 6)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 24)
        Me.Button3.TabIndex = 2
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(112, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 24)
        Me.Button2.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(6, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 24)
        Me.Button1.TabIndex = 0
        '
        'Purchase
        '
        Me.Purchase.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Purchase.Controls.Add(Me.PictureBox13)
        Me.Purchase.Controls.Add(Me.PictureBox14)
        Me.Purchase.Controls.Add(Me.PictureBox15)
        Me.Purchase.Controls.Add(Me.PictureBox16)
        Me.Purchase.Controls.Add(Me.PictureBox17)
        Me.Purchase.Controls.Add(Me.PictureBox18)
        Me.Purchase.Controls.Add(Me.PictureBox19)
        Me.Purchase.Controls.Add(Me.PictureBox20)
        Me.Purchase.Controls.Add(Me.PictureBox21)
        Me.Purchase.Controls.Add(Me.PictureBox22)
        Me.Purchase.Controls.Add(Me.PictureBox23)
        Me.Purchase.Controls.Add(Me.PictureBox24)
        Me.Purchase.Location = New System.Drawing.Point(6, 4)
        Me.Purchase.Name = "Purchase"
        Me.Purchase.Size = New System.Drawing.Size(544, 380)
        Me.Purchase.TabIndex = 54
        '
        'PictureBox13
        '
        Me.PictureBox13.Image = CType(resources.GetObject("PictureBox13.Image"), System.Drawing.Image)
        Me.PictureBox13.Location = New System.Drawing.Point(-2, 2)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(548, 32)
        Me.PictureBox13.TabIndex = 64
        Me.PictureBox13.TabStop = False
        '
        'PictureBox14
        '
        Me.PictureBox14.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox14.Image = CType(resources.GetObject("PictureBox14.Image"), System.Drawing.Image)
        Me.PictureBox14.Location = New System.Drawing.Point(10, 162)
        Me.PictureBox14.Name = "PictureBox14"
        Me.PictureBox14.Size = New System.Drawing.Size(84, 100)
        Me.PictureBox14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox14.TabIndex = 63
        Me.PictureBox14.TabStop = False
        '
        'PictureBox15
        '
        Me.PictureBox15.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox15.Image = CType(resources.GetObject("PictureBox15.Image"), System.Drawing.Image)
        Me.PictureBox15.Location = New System.Drawing.Point(421, 292)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(118, 86)
        Me.PictureBox15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox15.TabIndex = 61
        Me.PictureBox15.TabStop = False
        '
        'PictureBox16
        '
        Me.PictureBox16.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox16.Image = CType(resources.GetObject("PictureBox16.Image"), System.Drawing.Image)
        Me.PictureBox16.Location = New System.Drawing.Point(224, 278)
        Me.PictureBox16.Name = "PictureBox16"
        Me.PictureBox16.Size = New System.Drawing.Size(104, 90)
        Me.PictureBox16.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox16.TabIndex = 60
        Me.PictureBox16.TabStop = False
        '
        'PictureBox17
        '
        Me.PictureBox17.Image = CType(resources.GetObject("PictureBox17.Image"), System.Drawing.Image)
        Me.PictureBox17.Location = New System.Drawing.Point(92, 66)
        Me.PictureBox17.Name = "PictureBox17"
        Me.PictureBox17.Size = New System.Drawing.Size(138, 58)
        Me.PictureBox17.TabIndex = 62
        Me.PictureBox17.TabStop = False
        '
        'PictureBox18
        '
        Me.PictureBox18.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox18.Image = CType(resources.GetObject("PictureBox18.Image"), System.Drawing.Image)
        Me.PictureBox18.Location = New System.Drawing.Point(2, 52)
        Me.PictureBox18.Name = "PictureBox18"
        Me.PictureBox18.Size = New System.Drawing.Size(110, 102)
        Me.PictureBox18.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox18.TabIndex = 53
        Me.PictureBox18.TabStop = False
        '
        'PictureBox19
        '
        Me.PictureBox19.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox19.Image = CType(resources.GetObject("PictureBox19.Image"), System.Drawing.Image)
        Me.PictureBox19.Location = New System.Drawing.Point(8, 280)
        Me.PictureBox19.Name = "PictureBox19"
        Me.PictureBox19.Size = New System.Drawing.Size(108, 88)
        Me.PictureBox19.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox19.TabIndex = 59
        Me.PictureBox19.TabStop = False
        '
        'PictureBox20
        '
        Me.PictureBox20.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox20.Image = CType(resources.GetObject("PictureBox20.Image"), System.Drawing.Image)
        Me.PictureBox20.Location = New System.Drawing.Point(456, 166)
        Me.PictureBox20.Name = "PictureBox20"
        Me.PictureBox20.Size = New System.Drawing.Size(74, 100)
        Me.PictureBox20.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox20.TabIndex = 58
        Me.PictureBox20.TabStop = False
        '
        'PictureBox21
        '
        Me.PictureBox21.Image = CType(resources.GetObject("PictureBox21.Image"), System.Drawing.Image)
        Me.PictureBox21.Location = New System.Drawing.Point(106, 208)
        Me.PictureBox21.Name = "PictureBox21"
        Me.PictureBox21.Size = New System.Drawing.Size(336, 56)
        Me.PictureBox21.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox21.TabIndex = 57
        Me.PictureBox21.TabStop = False
        '
        'PictureBox22
        '
        Me.PictureBox22.Image = CType(resources.GetObject("PictureBox22.Image"), System.Drawing.Image)
        Me.PictureBox22.Location = New System.Drawing.Point(314, 68)
        Me.PictureBox22.Name = "PictureBox22"
        Me.PictureBox22.Size = New System.Drawing.Size(140, 58)
        Me.PictureBox22.TabIndex = 56
        Me.PictureBox22.TabStop = False
        '
        'PictureBox23
        '
        Me.PictureBox23.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox23.Image = CType(resources.GetObject("PictureBox23.Image"), System.Drawing.Image)
        Me.PictureBox23.Location = New System.Drawing.Point(228, 54)
        Me.PictureBox23.Name = "PictureBox23"
        Me.PictureBox23.Size = New System.Drawing.Size(102, 96)
        Me.PictureBox23.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox23.TabIndex = 55
        Me.PictureBox23.TabStop = False
        '
        'PictureBox24
        '
        Me.PictureBox24.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox24.Image = CType(resources.GetObject("PictureBox24.Image"), System.Drawing.Image)
        Me.PictureBox24.Location = New System.Drawing.Point(442, 52)
        Me.PictureBox24.Name = "PictureBox24"
        Me.PictureBox24.Size = New System.Drawing.Size(104, 102)
        Me.PictureBox24.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox24.TabIndex = 54
        Me.PictureBox24.TabStop = False
        '
        'Account
        '
        Me.Account.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Account.Controls.Add(Me.PictureBox25)
        Me.Account.Controls.Add(Me.PictureBox27)
        Me.Account.Controls.Add(Me.PictureBox30)
        Me.Account.Controls.Add(Me.PictureBox31)
        Me.Account.Controls.Add(Me.PictureBox33)
        Me.Account.Controls.Add(Me.PictureBox35)
        Me.Account.Controls.Add(Me.PictureBox36)
        Me.Account.Location = New System.Drawing.Point(6, 4)
        Me.Account.Name = "Account"
        Me.Account.Size = New System.Drawing.Size(544, 380)
        Me.Account.TabIndex = 55
        '
        'PictureBox25
        '
        Me.PictureBox25.Image = CType(resources.GetObject("PictureBox25.Image"), System.Drawing.Image)
        Me.PictureBox25.Location = New System.Drawing.Point(-2, 2)
        Me.PictureBox25.Name = "PictureBox25"
        Me.PictureBox25.Size = New System.Drawing.Size(548, 30)
        Me.PictureBox25.TabIndex = 64
        Me.PictureBox25.TabStop = False
        '
        'PictureBox27
        '
        Me.PictureBox27.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox27.Image = CType(resources.GetObject("PictureBox27.Image"), System.Drawing.Image)
        Me.PictureBox27.Location = New System.Drawing.Point(422, 286)
        Me.PictureBox27.Name = "PictureBox27"
        Me.PictureBox27.Size = New System.Drawing.Size(118, 86)
        Me.PictureBox27.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox27.TabIndex = 61
        Me.PictureBox27.TabStop = False
        '
        'PictureBox30
        '
        Me.PictureBox30.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox30.Image = CType(resources.GetObject("PictureBox30.Image"), System.Drawing.Image)
        Me.PictureBox30.Location = New System.Drawing.Point(6, 52)
        Me.PictureBox30.Name = "PictureBox30"
        Me.PictureBox30.Size = New System.Drawing.Size(138, 102)
        Me.PictureBox30.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox30.TabIndex = 53
        Me.PictureBox30.TabStop = False
        '
        'PictureBox31
        '
        Me.PictureBox31.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox31.Image = CType(resources.GetObject("PictureBox31.Image"), System.Drawing.Image)
        Me.PictureBox31.Location = New System.Drawing.Point(8, 286)
        Me.PictureBox31.Name = "PictureBox31"
        Me.PictureBox31.Size = New System.Drawing.Size(108, 88)
        Me.PictureBox31.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox31.TabIndex = 59
        Me.PictureBox31.TabStop = False
        '
        'PictureBox33
        '
        Me.PictureBox33.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox33.Image = CType(resources.GetObject("PictureBox33.Image"), System.Drawing.Image)
        Me.PictureBox33.Location = New System.Drawing.Point(204, 264)
        Me.PictureBox33.Name = "PictureBox33"
        Me.PictureBox33.Size = New System.Drawing.Size(154, 104)
        Me.PictureBox33.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox33.TabIndex = 57
        Me.PictureBox33.TabStop = False
        '
        'PictureBox35
        '
        Me.PictureBox35.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox35.Image = CType(resources.GetObject("PictureBox35.Image"), System.Drawing.Image)
        Me.PictureBox35.Location = New System.Drawing.Point(230, 56)
        Me.PictureBox35.Name = "PictureBox35"
        Me.PictureBox35.Size = New System.Drawing.Size(102, 96)
        Me.PictureBox35.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox35.TabIndex = 55
        Me.PictureBox35.TabStop = False
        '
        'PictureBox36
        '
        Me.PictureBox36.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox36.Image = CType(resources.GetObject("PictureBox36.Image"), System.Drawing.Image)
        Me.PictureBox36.Location = New System.Drawing.Point(413, 56)
        Me.PictureBox36.Name = "PictureBox36"
        Me.PictureBox36.Size = New System.Drawing.Size(128, 102)
        Me.PictureBox36.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox36.TabIndex = 54
        Me.PictureBox36.TabStop = False
        '
        'Reports
        '
        Me.Reports.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Reports.Controls.Add(Me.PictureBox32)
        Me.Reports.Controls.Add(Me.PictureBox29)
        Me.Reports.Controls.Add(Me.PictureBox28)
        Me.Reports.Controls.Add(Me.PictureBox26)
        Me.Reports.Location = New System.Drawing.Point(6, 3)
        Me.Reports.Name = "Reports"
        Me.Reports.Size = New System.Drawing.Size(544, 383)
        Me.Reports.TabIndex = 56
        '
        'PictureBox32
        '
        Me.PictureBox32.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox32.Image = CType(resources.GetObject("PictureBox32.Image"), System.Drawing.Image)
        Me.PictureBox32.Location = New System.Drawing.Point(218, 234)
        Me.PictureBox32.Name = "PictureBox32"
        Me.PictureBox32.Size = New System.Drawing.Size(112, 92)
        Me.PictureBox32.TabIndex = 54
        Me.PictureBox32.TabStop = False
        '
        'PictureBox29
        '
        Me.PictureBox29.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox29.Image = CType(resources.GetObject("PictureBox29.Image"), System.Drawing.Image)
        Me.PictureBox29.Location = New System.Drawing.Point(350, 104)
        Me.PictureBox29.Name = "PictureBox29"
        Me.PictureBox29.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox29.TabIndex = 53
        Me.PictureBox29.TabStop = False
        '
        'PictureBox28
        '
        Me.PictureBox28.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox28.Image = CType(resources.GetObject("PictureBox28.Image"), System.Drawing.Image)
        Me.PictureBox28.Location = New System.Drawing.Point(84, 106)
        Me.PictureBox28.Name = "PictureBox28"
        Me.PictureBox28.Size = New System.Drawing.Size(98, 110)
        Me.PictureBox28.TabIndex = 52
        Me.PictureBox28.TabStop = False
        '
        'PictureBox26
        '
        Me.PictureBox26.Image = CType(resources.GetObject("PictureBox26.Image"), System.Drawing.Image)
        Me.PictureBox26.Location = New System.Drawing.Point(-2, 0)
        Me.PictureBox26.Name = "PictureBox26"
        Me.PictureBox26.Size = New System.Drawing.Size(546, 40)
        Me.PictureBox26.TabIndex = 51
        Me.PictureBox26.TabStop = False
        '
        'arrow1
        '
        Me.arrow1.Image = CType(resources.GetObject("arrow1.Image"), System.Drawing.Image)
        Me.arrow1.Location = New System.Drawing.Point(46, 386)
        Me.arrow1.Name = "arrow1"
        Me.arrow1.Size = New System.Drawing.Size(39, 25)
        Me.arrow1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.arrow1.TabIndex = 60
        Me.arrow1.TabStop = False
        '
        'arrow2
        '
        Me.arrow2.Image = CType(resources.GetObject("arrow2.Image"), System.Drawing.Image)
        Me.arrow2.Location = New System.Drawing.Point(152, 386)
        Me.arrow2.Name = "arrow2"
        Me.arrow2.Size = New System.Drawing.Size(39, 25)
        Me.arrow2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.arrow2.TabIndex = 61
        Me.arrow2.TabStop = False
        '
        'arrow3
        '
        Me.arrow3.Image = CType(resources.GetObject("arrow3.Image"), System.Drawing.Image)
        Me.arrow3.Location = New System.Drawing.Point(260, 386)
        Me.arrow3.Name = "arrow3"
        Me.arrow3.Size = New System.Drawing.Size(39, 25)
        Me.arrow3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.arrow3.TabIndex = 62
        Me.arrow3.TabStop = False
        '
        'arrow4
        '
        Me.arrow4.Image = CType(resources.GetObject("arrow4.Image"), System.Drawing.Image)
        Me.arrow4.Location = New System.Drawing.Point(366, 386)
        Me.arrow4.Name = "arrow4"
        Me.arrow4.Size = New System.Drawing.Size(39, 25)
        Me.arrow4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.arrow4.TabIndex = 63
        Me.arrow4.TabStop = False
        '
        'SalesReport
        '
        Me.SalesReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SalesReport.Controls.Add(Me.Button6)
        Me.SalesReport.Controls.Add(Me.DataGrid1)
        Me.SalesReport.Controls.Add(Me.PictureBox40)
        Me.SalesReport.Controls.Add(Me.PictureBox34)
        Me.SalesReport.Controls.Add(Me.PictureBox37)
        Me.SalesReport.Controls.Add(Me.PictureBox38)
        Me.SalesReport.Controls.Add(Me.PictureBox39)
        Me.SalesReport.Location = New System.Drawing.Point(6, 6)
        Me.SalesReport.Name = "SalesReport"
        Me.SalesReport.Size = New System.Drawing.Size(544, 380)
        Me.SalesReport.TabIndex = 64
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.White
        Me.Button6.Location = New System.Drawing.Point(98, 346)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(74, 18)
        Me.Button6.TabIndex = 58
        Me.Button6.Text = "Get Report"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.CaptionText = "Sales Report List"
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 40)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(260, 294)
        Me.DataGrid1.TabIndex = 56
        '
        'PictureBox40
        '
        Me.PictureBox40.Image = CType(resources.GetObject("PictureBox40.Image"), System.Drawing.Image)
        Me.PictureBox40.Location = New System.Drawing.Point(322, 396)
        Me.PictureBox40.Name = "PictureBox40"
        Me.PictureBox40.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox40.TabIndex = 55
        Me.PictureBox40.TabStop = False
        '
        'PictureBox34
        '
        Me.PictureBox34.Image = CType(resources.GetObject("PictureBox34.Image"), System.Drawing.Image)
        Me.PictureBox34.Location = New System.Drawing.Point(110, 396)
        Me.PictureBox34.Name = "PictureBox34"
        Me.PictureBox34.Size = New System.Drawing.Size(112, 118)
        Me.PictureBox34.TabIndex = 54
        Me.PictureBox34.TabStop = False
        '
        'PictureBox37
        '
        Me.PictureBox37.Image = CType(resources.GetObject("PictureBox37.Image"), System.Drawing.Image)
        Me.PictureBox37.Location = New System.Drawing.Point(216, 400)
        Me.PictureBox37.Name = "PictureBox37"
        Me.PictureBox37.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox37.TabIndex = 53
        Me.PictureBox37.TabStop = False
        '
        'PictureBox38
        '
        Me.PictureBox38.Image = CType(resources.GetObject("PictureBox38.Image"), System.Drawing.Image)
        Me.PictureBox38.Location = New System.Drawing.Point(4, 396)
        Me.PictureBox38.Name = "PictureBox38"
        Me.PictureBox38.Size = New System.Drawing.Size(98, 110)
        Me.PictureBox38.TabIndex = 52
        Me.PictureBox38.TabStop = False
        '
        'PictureBox39
        '
        Me.PictureBox39.Image = CType(resources.GetObject("PictureBox39.Image"), System.Drawing.Image)
        Me.PictureBox39.Location = New System.Drawing.Point(-2, 0)
        Me.PictureBox39.Name = "PictureBox39"
        Me.PictureBox39.Size = New System.Drawing.Size(546, 40)
        Me.PictureBox39.TabIndex = 51
        Me.PictureBox39.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Button7)
        Me.Panel2.Controls.Add(Me.DataGrid3)
        Me.Panel2.Controls.Add(Me.PictureBox41)
        Me.Panel2.Controls.Add(Me.PictureBox42)
        Me.Panel2.Controls.Add(Me.PictureBox43)
        Me.Panel2.Controls.Add(Me.PictureBox44)
        Me.Panel2.Controls.Add(Me.PictureBox45)
        Me.Panel2.Location = New System.Drawing.Point(6, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(544, 384)
        Me.Panel2.TabIndex = 65
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.ForeColor = System.Drawing.Color.White
        Me.Button7.Location = New System.Drawing.Point(100, 344)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(74, 18)
        Me.Button7.TabIndex = 59
        Me.Button7.Text = "Get Report"
        '
        'DataGrid3
        '
        Me.DataGrid3.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid3.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid3.CaptionText = "Purchase Report List"
        Me.DataGrid3.CaptionVisible = False
        Me.DataGrid3.DataMember = ""
        Me.DataGrid3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid3.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid3.Location = New System.Drawing.Point(12, 40)
        Me.DataGrid3.Name = "DataGrid3"
        Me.DataGrid3.Size = New System.Drawing.Size(260, 294)
        Me.DataGrid3.TabIndex = 57
        '
        'PictureBox41
        '
        Me.PictureBox41.Image = CType(resources.GetObject("PictureBox41.Image"), System.Drawing.Image)
        Me.PictureBox41.Location = New System.Drawing.Point(248, 400)
        Me.PictureBox41.Name = "PictureBox41"
        Me.PictureBox41.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox41.TabIndex = 55
        Me.PictureBox41.TabStop = False
        '
        'PictureBox42
        '
        Me.PictureBox42.Image = CType(resources.GetObject("PictureBox42.Image"), System.Drawing.Image)
        Me.PictureBox42.Location = New System.Drawing.Point(174, 400)
        Me.PictureBox42.Name = "PictureBox42"
        Me.PictureBox42.Size = New System.Drawing.Size(112, 106)
        Me.PictureBox42.TabIndex = 54
        Me.PictureBox42.TabStop = False
        '
        'PictureBox43
        '
        Me.PictureBox43.Image = CType(resources.GetObject("PictureBox43.Image"), System.Drawing.Image)
        Me.PictureBox43.Location = New System.Drawing.Point(350, 400)
        Me.PictureBox43.Name = "PictureBox43"
        Me.PictureBox43.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox43.TabIndex = 53
        Me.PictureBox43.TabStop = False
        '
        'PictureBox44
        '
        Me.PictureBox44.Image = CType(resources.GetObject("PictureBox44.Image"), System.Drawing.Image)
        Me.PictureBox44.Location = New System.Drawing.Point(84, 402)
        Me.PictureBox44.Name = "PictureBox44"
        Me.PictureBox44.Size = New System.Drawing.Size(98, 110)
        Me.PictureBox44.TabIndex = 52
        Me.PictureBox44.TabStop = False
        '
        'PictureBox45
        '
        Me.PictureBox45.Image = CType(resources.GetObject("PictureBox45.Image"), System.Drawing.Image)
        Me.PictureBox45.Location = New System.Drawing.Point(-2, 0)
        Me.PictureBox45.Name = "PictureBox45"
        Me.PictureBox45.Size = New System.Drawing.Size(546, 32)
        Me.PictureBox45.TabIndex = 51
        Me.PictureBox45.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Button5)
        Me.Panel3.Controls.Add(Me.DataGrid2)
        Me.Panel3.Controls.Add(Me.PictureBox48)
        Me.Panel3.Controls.Add(Me.PictureBox49)
        Me.Panel3.Controls.Add(Me.PictureBox50)
        Me.Panel3.Location = New System.Drawing.Point(6, 4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(544, 382)
        Me.Panel3.TabIndex = 66
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(98, 344)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(74, 18)
        Me.Button5.TabIndex = 60
        Me.Button5.Text = "Get Report"
        '
        'DataGrid2
        '
        Me.DataGrid2.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid2.CaptionFont = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid2.CaptionText = "Account Report List"
        Me.DataGrid2.CaptionVisible = False
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(8, 38)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(260, 294)
        Me.DataGrid2.TabIndex = 59
        '
        'PictureBox48
        '
        Me.PictureBox48.Image = CType(resources.GetObject("PictureBox48.Image"), System.Drawing.Image)
        Me.PictureBox48.Location = New System.Drawing.Point(350, 416)
        Me.PictureBox48.Name = "PictureBox48"
        Me.PictureBox48.Size = New System.Drawing.Size(104, 108)
        Me.PictureBox48.TabIndex = 53
        Me.PictureBox48.TabStop = False
        '
        'PictureBox49
        '
        Me.PictureBox49.Image = CType(resources.GetObject("PictureBox49.Image"), System.Drawing.Image)
        Me.PictureBox49.Location = New System.Drawing.Point(84, 418)
        Me.PictureBox49.Name = "PictureBox49"
        Me.PictureBox49.Size = New System.Drawing.Size(98, 110)
        Me.PictureBox49.TabIndex = 52
        Me.PictureBox49.TabStop = False
        '
        'PictureBox50
        '
        Me.PictureBox50.Image = CType(resources.GetObject("PictureBox50.Image"), System.Drawing.Image)
        Me.PictureBox50.Location = New System.Drawing.Point(-2, 0)
        Me.PictureBox50.Name = "PictureBox50"
        Me.PictureBox50.Size = New System.Drawing.Size(546, 32)
        Me.PictureBox50.TabIndex = 51
        Me.PictureBox50.TabStop = False
        '
        'Tool
        '
        Me.Tool.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Tool.Controls.Add(Me.PictureBox8)
        Me.Tool.Controls.Add(Me.PictureBox51)
        Me.Tool.Controls.Add(Me.PictureBox56)
        Me.Tool.Controls.Add(Me.PictureBox58)
        Me.Tool.Controls.Add(Me.PictureBox59)
        Me.Tool.Location = New System.Drawing.Point(8, 0)
        Me.Tool.Name = "Tool"
        Me.Tool.Size = New System.Drawing.Size(544, 384)
        Me.Tool.TabIndex = 67
        Me.Tool.Visible = False
        '
        'PictureBox8
        '
        Me.PictureBox8.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(32, 222)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(104, 116)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox8.TabIndex = 39
        Me.PictureBox8.TabStop = False
        '
        'PictureBox51
        '
        Me.PictureBox51.Image = CType(resources.GetObject("PictureBox51.Image"), System.Drawing.Image)
        Me.PictureBox51.Location = New System.Drawing.Point(-2, 2)
        Me.PictureBox51.Name = "PictureBox51"
        Me.PictureBox51.Size = New System.Drawing.Size(546, 30)
        Me.PictureBox51.TabIndex = 51
        Me.PictureBox51.TabStop = False
        '
        'PictureBox56
        '
        Me.PictureBox56.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox56.Image = CType(resources.GetObject("PictureBox56.Image"), System.Drawing.Image)
        Me.PictureBox56.Location = New System.Drawing.Point(402, 226)
        Me.PictureBox56.Name = "PictureBox56"
        Me.PictureBox56.Size = New System.Drawing.Size(104, 114)
        Me.PictureBox56.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox56.TabIndex = 43
        Me.PictureBox56.TabStop = False
        '
        'PictureBox58
        '
        Me.PictureBox58.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox58.Image = CType(resources.GetObject("PictureBox58.Image"), System.Drawing.Image)
        Me.PictureBox58.Location = New System.Drawing.Point(28, 62)
        Me.PictureBox58.Name = "PictureBox58"
        Me.PictureBox58.Size = New System.Drawing.Size(122, 118)
        Me.PictureBox58.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox58.TabIndex = 41
        Me.PictureBox58.TabStop = False
        '
        'PictureBox59
        '
        Me.PictureBox59.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox59.Image = CType(resources.GetObject("PictureBox59.Image"), System.Drawing.Image)
        Me.PictureBox59.Location = New System.Drawing.Point(396, 72)
        Me.PictureBox59.Name = "PictureBox59"
        Me.PictureBox59.Size = New System.Drawing.Size(124, 100)
        Me.PictureBox59.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox59.TabIndex = 40
        Me.PictureBox59.TabStop = False
        '
        'arrow5
        '
        Me.arrow5.Image = CType(resources.GetObject("arrow5.Image"), System.Drawing.Image)
        Me.arrow5.Location = New System.Drawing.Point(476, 386)
        Me.arrow5.Name = "arrow5"
        Me.arrow5.Size = New System.Drawing.Size(39, 25)
        Me.arrow5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.arrow5.TabIndex = 68
        Me.arrow5.TabStop = False
        Me.arrow5.Visible = False
        '
        'BackForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(560, 488)
        Me.Controls.Add(Me.arrow5)
        Me.Controls.Add(Me.arrow4)
        Me.Controls.Add(Me.arrow3)
        Me.Controls.Add(Me.arrow2)
        Me.Controls.Add(Me.arrow1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.SalesReport)
        Me.Controls.Add(Me.Reports)
        Me.Controls.Add(Me.Account)
        Me.Controls.Add(Me.Purchase)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Tool)
        Me.Controls.Add(Me.Sales)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BackForm"
        Me.ShowInTaskbar = False
        Me.Text = "BackForm"
        Me.Sales.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Purchase.ResumeLayout(False)
        Me.Account.ResumeLayout(False)
        Me.Reports.ResumeLayout(False)
        Me.SalesReport.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Tool.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PictureBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CreateCustomer()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
        Me.SendToBack()
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerSummary()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New CustomerAmend()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New SalesInvoice()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New SalesCreditInvoice()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New CustomerActivity()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New CustomerReceipts()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CC As New CustomerRefund()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CreateCustomer()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerSummary()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerAmend()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New SalesInvoice()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New SalesCreditInvoice()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerActivity()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerReceipts()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CC As New CustomerRefund()
        CC.MdiParent = Me.MdiParent
        CC.Location = New System.Drawing.Point(0, 10)
        CC.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        arrow5.Visible = False
        Tool.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        SalesReport.Visible = False
        arrow1.Visible = False
        arrow2.Visible = True
        arrow3.Visible = False
        arrow4.Visible = False
        Reports.Visible = False
        Account.Visible = False
        Sales.Visible = False
        Purchase.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        arrow1.Visible = True
        arrow2.Visible = False
        arrow3.Visible = False
        arrow4.Visible = False
        arrow5.Visible = False
        Tool.Visible = False
        Reports.Visible = False
        Account.Visible = False
        Purchase.Visible = False
        Sales.Visible = True
        Panel2.Visible = False
        Panel3.Visible = False
        SalesReport.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        arrow5.Visible = False
        Tool.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        SalesReport.Visible = False
        arrow1.Visible = False
        arrow2.Visible = False
        arrow3.Visible = True
        arrow4.Visible = False
        Reports.Visible = False
        Purchase.Visible = False
        Sales.Visible = False
        Account.Visible = True
    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Cvvendor As New CreateVendor()
        Cvvendor.MdiParent = Me.MdiParent
        Cvvendor.Location = New System.Drawing.Point(0, 10)
        Cvvendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New VendorSummary()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New VendorAmend()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New PurchaseInvoice()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New PurchaseCreditInvoice()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New VendorActivity()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New VendorPayment()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim vendor As New VendorRefund()
        vendor.MdiParent = Me.MdiParent
        vendor.Location = New System.Drawing.Point(0, 10)
        vendor.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox30.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New CompanyPreferences()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox35.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New BankAmend()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox36.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New LedgerSummary()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox33.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New VAT()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox31.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New TaxCode()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New TaxCode()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox27.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Account As New FinancialYear()
        Account.MdiParent = Me.MdiParent
        Account.Location = New System.Drawing.Point(0, 10)
        Account.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        arrow5.Visible = True
        Tool.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        SalesReport.Visible = False
        arrow1.Visible = False
        arrow2.Visible = False
        arrow3.Visible = False
        arrow4.Visible = False
        Account.Visible = False
        Purchase.Visible = False
        Sales.Visible = False
        Reports.Visible = True
    End Sub

    Private Sub PictureBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseEnter
        PictureBox1.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox3.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox2.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox5_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox5.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox9_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox9.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox10_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox10.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox11_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox11.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox12_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseEnter
        PictureBox12.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox18_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox18.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox23_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox23.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox14_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox14.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox24_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox24.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox20_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox20.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox19_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox19.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox16_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox16.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox15_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseEnter
        PictureBox15.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox30_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox30.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox31_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox31.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox35_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox35.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox36_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox36.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox33_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox33.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox27_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseEnter
        PictureBox27.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox28_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.MouseEnter
        PictureBox28.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox29_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.MouseEnter
        PictureBox29.Cursor = Cursors.Hand
    End Sub
    Private Sub PictureBox32_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.MouseEnter
        PictureBox32.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox38.Click
        ShowReport.ReportName = "CustomerList"
        Dim SR As New ShowReport()
        ' SR.MdiParent = Me.MdiParent
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub PictureBox37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox37.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim CidforActivity As New GetCustomerID()
        CidforActivity.MdiParent = Me.MdiParent
        CidforActivity.Show()
        CidforActivity.Activate()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox34.Click
        ShowReport.ReportName = "SalesInvoice"
        Dim CidforActivity As New GetCustomerID()
        CidforActivity.MdiParent = Me.MdiParent
        CidforActivity.Show()
        CidforActivity.Activate()
    End Sub

    Private Sub PictureBox40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox40.Click
        ShowReport.ReportName = "SalesReceipts"
        Dim CidforActivity As New GetCustomerID()
        CidforActivity.MdiParent = Me.MdiParent
        CidforActivity.Show()
        CidforActivity.Activate()
    End Sub

    Private Sub PictureBox44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox44.Click
        ShowReport.ReportName = "VendorList"
        Dim SR As New ShowReport()
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub PictureBox43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox43.Click
        ShowReport.ReportName = "VendorActivity"
        Dim VendorId As New GetVendorID()
        VendorId.MdiParent = Me.MdiParent
        VendorId.Show()
        VendorId.Activate()

    End Sub

    Private Sub PictureBox42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox42.Click
        ShowReport.ReportName = "PurchaseInvoices"
        Dim VendorId As New GetVendorID()
        VendorId.MdiParent = Me.MdiParent
        VendorId.Show()
        VendorId.Activate()
    End Sub

    Private Sub PictureBox41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox41.Click
        ShowReport.ReportName = "PurchasePayments"
        Dim VendorId As New GetVendorID()
        VendorId.MdiParent = Me.MdiParent
        VendorId.Show()
        VendorId.Activate()
    End Sub

    Private Sub PictureBox49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox49.Click
        ShowReport.ReportName = "LedgerBalance"
        Dim SR As New ShowReport()
        SR.Show()
        SR.Activate()
    End Sub

    Private Sub PictureBox48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox48.Click
        ShowReport.ReportName = "LedgerActivity"
        Dim LedgerId As New GetLedgerID()
        LedgerId.MdiParent = Me.MdiParent
        LedgerId.Show()
        LedgerId.Activate()
    End Sub

    Private Sub PictureBox28_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.Click

        Reports.Visible = False
        SalesReport.Visible = True
        Dim Cart As New DataTable()
        Dim CartDS As New DataSet()
        Cart.TableName = "Cart"
        Cart.Columns.Add(New DataColumn("Report Name", GetType(String)))
        Cart.Columns.Add(New DataColumn("Description", GetType(String)))

        Cart.Rows.Add(New Object() {"Customer List"})
        Cart.Rows.Add(New Object() {"Customer Activity"})
        Cart.Rows.Add(New Object() {"Sales Invoice"})
        Cart.Rows.Add(New Object() {"Sales Receipts"})
        Cart.Rows.Add(New Object() {"Customer Creditcard Activity"})

        CartDS.Tables.Add(Cart)

        Dim ts As New DataGridTableStyle(), cs As DataGridBrowser.DataGridNoActiveCellColumn
        DataGrid1.TableStyles.Clear()

        ' Add the custom column style.
        cs = New DataGridBrowser.DataGridNoActiveCellColumn()
        cs.Width = 240
        cs.MappingName = "Report Name"
        cs.HeaderText = "Report Name"
        ts.GridColumnStyles.Add(cs)

        ts.MappingName = "Cart"
        ts.RowHeadersVisible = False
        DataGrid1.TableStyles.Add(ts)
        DataGrid1.DataSource = CartDS
        DataGrid1.DataMember = "Cart"

        'To delete the extra row
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

    End Sub

    Private Sub PictureBox32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox32.Click

        Reports.Visible = False
        Panel3.Visible = True
        Dim Cart As New DataTable()
        Dim CartDS As New DataSet()
        Cart.TableName = "Cart"
        Cart.Columns.Add(New DataColumn("Report Name", GetType(String)))
        Cart.Columns.Add(New DataColumn("Description", GetType(String)))

        Cart.Rows.Add(New Object() {"Ledger Activity"})
        Cart.Rows.Add(New Object() {"Ledger Balance"})

        CartDS.Tables.Add(Cart)

        Dim ts As New DataGridTableStyle(), cs As DataGridBrowser.DataGridNoActiveCellColumn
        DataGrid2.TableStyles.Clear()

        ' Add the custom column style.
        cs = New DataGridBrowser.DataGridNoActiveCellColumn()
        cs.Width = 240
        cs.MappingName = "Report Name"
        cs.HeaderText = "Report Name"
        ts.GridColumnStyles.Add(cs)

        ts.MappingName = "Cart"
        ts.RowHeadersVisible = False
        DataGrid2.TableStyles.Add(ts)
        DataGrid2.DataSource = CartDS
        DataGrid2.DataMember = "Cart"

        'To delete the extra row
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid2.DataSource, DataGrid2.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

    End Sub

    Private Sub PictureBox29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox29.Click

        Reports.Visible = False
        Panel2.Visible = True

        Dim Cart As New DataTable()
        Dim CartDS As New DataSet()
        Cart.TableName = "Cart"
        Cart.Columns.Add(New DataColumn("Report Name", GetType(String)))
        Cart.Columns.Add(New DataColumn("Description", GetType(String)))

        Cart.Rows.Add(New Object() {"Vendor List"})
        Cart.Rows.Add(New Object() {"Vendor Activity"})
        Cart.Rows.Add(New Object() {"Purchase Invoice"})
        Cart.Rows.Add(New Object() {"Purchase Payment"})

        CartDS.Tables.Add(Cart)

        Dim ts As New DataGridTableStyle(), cs As DataGridBrowser.DataGridNoActiveCellColumn
        DataGrid3.TableStyles.Clear()

        ' Add the custom column style.
        cs = New DataGridBrowser.DataGridNoActiveCellColumn()
        cs.Width = 240
        cs.MappingName = "Report Name"
        cs.HeaderText = "Report Name"
        ts.GridColumnStyles.Add(cs)

        ts.MappingName = "Cart"
        ts.RowHeadersVisible = False
        DataGrid3.TableStyles.Add(ts)
        DataGrid3.DataSource = CartDS
        DataGrid3.DataMember = "Cart"

        'To delete the extra row
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid3.DataSource, DataGrid3.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

    End Sub

    Private Sub PictureBox44_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox44.MouseEnter
        PictureBox44.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox43_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox43.MouseEnter
        PictureBox42.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox42_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox42.MouseEnter
        PictureBox42.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox41_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox41.MouseEnter
        PictureBox41.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox38_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox38.MouseEnter
        PictureBox38.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox37_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox37.MouseEnter
        PictureBox37.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox34_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox34.MouseEnter
        PictureBox34.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox40_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox40.MouseEnter
        PictureBox40.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox49_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox49.MouseEnter
        PictureBox49.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox48_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox48.MouseEnter
        PictureBox48.Cursor = Cursors.Hand
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim str As String = DataGrid1.Item(Crow, DataGrid1.CurrentCell.ColumnNumber)

        Select Case str
            Case "Customer List"
                ShowReport.ReportName = "CustomerList"
                Dim SR As New ShowReport()
                ' SR.MdiParent = Me.MdiParent
                SR.Show()
                SR.Activate()
            Case "Sales Invoice"
                ShowReport.ReportName = "SalesInvoice"
                Dim CidforActivity As New GetCustomerID()
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
            Case "Sales Receipts"
                ShowReport.ReportName = "SalesReceipts"
                Dim CidforActivity As New GetCustomerID()
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
            Case "Customer Activity"
                Dim CidforActivity As New GetCustomerID()
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
        End Select

    End Sub

    Private Sub DataGrid1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.Click
        Crow = DataGrid1.CurrentRowIndex
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim str As String = DataGrid1.Item(Crow, DataGrid1.CurrentCell.ColumnNumber)

        Select Case str
            Case "Customer List"
                ShowReport.ReportName = "CustomerList"
                Dim SR As New ShowReport()
                ' SR.MdiParent = Me.MdiParent
                SR.Show()
                SR.Activate()
            Case "Sales Invoice"
                ShowReport.ReportName = "SalesInvoice"
                Dim CidforActivity As New GetCustomerID()
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
            Case "Sales Receipts"
                ShowReport.ReportName = "SalesReceipts"
                Dim CidforActivity As New GetCustomerID()
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
            Case "Customer Activity"
                ShowReport.ReportName = "CustomerActivity"
                Dim CidforActivity As New GetCustomerID
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()
            Case "Customer Creditcard Activity"
                ShowReport.ReportName = "Customer Creditcard Activity"
                Dim CidforActivity As New CustomerCreditInfo
                CidforActivity.MdiParent = Me.MdiParent
                CidforActivity.Show()
                CidforActivity.Activate()

        End Select
    End Sub

    Private Sub DataGrid2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid2.Click
        Crow = DataGrid2.CurrentRowIndex
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim str As String = DataGrid2.Item(Crow, DataGrid2.CurrentCell.ColumnNumber)

        Select Case str
            Case "Ledger Balance"
                ShowReport.ReportName = "LedgerBalance"
                Dim SR As New ShowReport()
                SR.Show()
                SR.Activate()
            Case "Ledger Activity"
                ShowReport.ReportName = "LedgerActivity"
                Dim LedgerId As New GetLedgerID()
                LedgerId.MdiParent = Me.MdiParent
                LedgerId.Show()
                LedgerId.Activate()
        End Select
    End Sub

    Private Sub DataGrid3_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles DataGrid3.Navigate
        Crow = DataGrid3.CurrentRowIndex
    End Sub

    Private Sub DataGrid3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid3.Click
        Crow = DataGrid3.CurrentRowIndex
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim str As String = DataGrid3.Item(Crow, DataGrid3.CurrentCell.ColumnNumber)

        Select Case str
            Case "Vendor List"
                ShowReport.ReportName = "VendorList"
                Dim SR As New ShowReport()
                SR.Show()
                SR.Activate()
            Case "Purchase Invoice"
                ShowReport.ReportName = "PurchaseInvoices"
                Dim VendorId As New GetVendorID()
                VendorId.MdiParent = Me.MdiParent
                VendorId.Show()
                VendorId.Activate()
            Case "Purchase Payment"
                ShowReport.ReportName = "PurchasePayments"
                Dim VendorId As New GetVendorID()
                VendorId.MdiParent = Me.MdiParent
                VendorId.Show()
                VendorId.Activate()
            Case "Vendor Activity"
                ShowReport.ReportName = "VendorActivity"
                Dim VendorId As New GetVendorID()
                VendorId.MdiParent = Me.MdiParent
                VendorId.Show()
                VendorId.Activate()
        End Select
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        arrow5.Visible = False
        Tool.Visible = True
        arrow1.Visible = False
        arrow2.Visible = False
        arrow3.Visible = False
        arrow4.Visible = True
        Reports.Visible = False
        Account.Visible = False
        Purchase.Visible = False
        Sales.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        SalesReport.Visible = False
    End Sub

    Private Sub PictureBox58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox58.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Bform As New Backup()
        Bform.MdiParent = Me.MdiParent
        Bform.Show()
        Bform.Activate()
        MDI.numOfForms = 1
    End Sub

    Private Sub PictureBox59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox59.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Bform As New Restore()
        Bform.MdiParent = Me.MdiParent
        Bform.Show()
        Bform.Activate()
        MDI.numOfForms = 1

    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim Bform As New ImpExpUtility()
        Bform.MdiParent = Me.MdiParent
        Bform.Show()
        Bform.Activate()
        MDI.numOfForms = 1

    End Sub

    Private Sub PictureBox56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox56.Click
        If MDI.numOfForms = 1 Then
            MessageBox.Show("You must close the open window first.    ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim child As New FrmBar
        'child.MdiParent = Me
        child.Show()
        MDI.numOfForms = 1
    End Sub

    Private Sub Purchase_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Purchase.Paint
        Me.SendToBack()
    End Sub

    Private Sub Sales_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Sales.Paint
        Me.SendToBack()
    End Sub

    Private Sub Tool_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Tool.Paint
        Me.SendToBack()
    End Sub

    Private Sub Reports_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Reports.Paint
        Me.SendToBack()
    End Sub

    Private Sub Account_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Account.Paint
        Me.SendToBack()
    End Sub
End Class
