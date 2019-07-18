Imports InfiniExpress.BLL

Public Class CustomerSummary

    Inherits System.Windows.Forms.Form
    Public ClsSales As New ClassSales()
    Dim DS As DataSet

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
    Protected WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CustomerSummary))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.DataGrid1, Me.Button1})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(530, 242)
        Me.Panel1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(6, 214)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 18)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "&Cancel"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackColor = System.Drawing.Color.White
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(6, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsVisible = False
        Me.DataGrid1.RowHeadersVisible = False
        Me.DataGrid1.Size = New System.Drawing.Size(516, 198)
        Me.DataGrid1.TabIndex = 8
        '
        'CustomerSummary
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(538, 249)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustomerSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Summary"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CustomerSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populatingGrid()
    End Sub

    Sub populatingGrid()

        DS = ClsSales.LoadCustSummary
        DataGrid1.TableStyles.Clear()
        'Formating The DataGrid Bt Adding a DataGrid Table Style
        Dim DgridTstyle As New DataGridTableStyle()
        Dim CustomerID As New DataGridBrowser.DataGridNoActiveCellColumn() 'parentid
        Dim CustomerName As New DataGridBrowser.DataGridNoActiveCellColumn()    'InvTp
        Dim Email As New DataGridBrowser.DataGridNoActiveCellColumn() 'InvDetails
        Dim SubBal As New DataGridBrowser.DataGridNoActiveCellColumn()  'ODebit

        'Setting Default Properties for datagrid like backcolor,forecolor,fonts
        With DgridTstyle
            .HeaderFont = New System.Drawing.Font _
            ("Verdana", 8.25F, System.Drawing.FontStyle.Regular, _
             System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
            .AlternatingBackColor = Color.Lavender
        End With

        With CustomerID
            .HeaderText = "Customer ID"
            .MappingName = "customerid"
            .Width = 80
            .ReadOnly = True
        End With

        With CustomerName
            .HeaderText = "Customer Name"
            .MappingName = "CustomerName"
            .Width = 190
            .NullText = ""
        End With

        With Email
            .HeaderText = "Email"
            .MappingName = "Email"
            .Width = 148
            .NullText = ""
        End With

        With SubBal
            .HeaderText = "Balance" & " " & Chr(9)
            .MappingName = "Balance"
            .Width = 77
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .NullText = "0.00"
            .Format = "#0.00"
        End With

        'Add The GridColumnStyle to Datagrid Column Style
        With DgridTstyle.GridColumnStyles
            .Add(CustomerID)
            .Add(CustomerName)
            .Add(Email)
            .Add(SubBal)
        End With

        DataGrid1.DataSource = DS.Tables(0)
        Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
        CType(Cm.List, DataView).AllowNew = False
        DgridTstyle.RowHeadersVisible = False
        DgridTstyle.MappingName = DS.Tables(0).TableName
        DataGrid1.TableStyles.Add(DgridTstyle)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose(True)
    End Sub

End Class
