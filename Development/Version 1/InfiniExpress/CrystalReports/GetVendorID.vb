Imports InfiniExpress.BLL
Public Class GetVendorID
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(GetVendorID))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnClose, Me.DataGrid1, Me.Button1})
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(336, 362)
        Me.Panel1.TabIndex = 0
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionText = "Select Vendor"
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 10)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(320, 320)
        Me.DataGrid1.TabIndex = 8
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(8, 336)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(74, 18)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Get Report"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(96, 336)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(74, 18)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        '
        'GetVendorID
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(344, 371)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GetVendorID"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Select Vendor"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim myds As New DataSet()
    Dim GlbReports As New ClassReports()

    Private Sub GetVendorID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim DgridTstyle As New DataGridTableStyle()
        myds = GlbReports.GetVendID()
        Dim VendorID As New DataGridBrowser.DataGridNoActiveCellColumn() 'parentid
        Dim VendorName As New DataGridBrowser.DataGridNoActiveCellColumn()
        With DgridTstyle
            '.HeaderBackColor = Color.Navy
            '.HeaderForeColor = Color.White
            .HeaderFont = New System.Drawing.Font _
            ("Verdana", 8.25F, System.Drawing.FontStyle.Regular, _
             System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
            .SelectionBackColor = Color.Navy
            .AlternatingBackColor = Color.Lavender
        End With

        With VendorID
            .HeaderText = "Vendor ID"
            .MappingName = "VendorID"
            .Width = DataGrid1.Width / 3 + 3
            .ReadOnly = True
        End With

        With VendorName
            .HeaderText = "Vendor Name"
            .MappingName = "VendorName"
            .Width = 190
            .NullText = ""
        End With

        With DgridTstyle.GridColumnStyles
            .Add(VendorID)
            .Add(VendorName)
        End With
        If myds.Tables(0).Rows.Count > 0 Then
            DataGrid1.DataSource = myds.Tables(0)

            'Stoping Extra Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False

            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = myds.Tables(0).TableName '("TBLCUSTOMER")
            DataGrid1.TableStyles.Add(DgridTstyle)

        End If

    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ShowReport.VendId = "" Then
            MessageBox.Show("You must select a vendor first.", "Infini Accounts", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim objSR As New ShowReport()
        Me.Dispose(True)
        objSR.Show()
        ShowReport.VendId = ""
        'MDI.numOfForms = 1
    End Sub
    Private Sub datagrid_Doubleclick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.DoubleClick
        Try
            ShowReport.VendId = Convert.ToString(DataGrid1.Item(DataGrid1.CurrentCell.RowNumber, 0))
            ShowReport.VendName = Convert.ToString(DataGrid1.Item(DataGrid1.CurrentCell.RowNumber, 1))
            Dim objSR As New ShowReport()
            Me.Dispose(True)
            objSR.Show()
            ShowReport.VendId = ""
        Catch ex As Exception
            Exit Sub
        End Try
        'MDI.numOfForms = 1
    End Sub

    Private Sub datagrid_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.Click
        Try
            ShowReport.VendId = Convert.ToString(DataGrid1.Item(DataGrid1.CurrentCell.RowNumber, 0))
            ShowReport.VendName = Convert.ToString(DataGrid1.Item(DataGrid1.CurrentCell.RowNumber, 1))
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Dispose(True)
    End Sub
End Class
