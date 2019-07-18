'Imports System
'Imports System.Data
Imports InfiniExpress.BLL

Public Class CustomerActivity

    Inherits System.Windows.Forms.Form
    Public ClsSales As New ClassSales()
    Public ClsGlobal As New ClassGlobal()
    Public Ds As New DataSet()
    Public Dr As IDataReader
    Dim CustID As Integer

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
    Protected WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Protected WithEvents btncancel As System.Windows.Forms.Button
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTableStyle2 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Protected WithEvents Label7 As System.Windows.Forms.Label
    Protected WithEvents cmbcid As System.Windows.Forms.ComboBox
    Protected WithEvents Label8 As System.Windows.Forms.Label
    Protected WithEvents txtcname As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CustomerActivity))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.txtcname = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cmbcid = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btncancel = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTableStyle2 = New System.Windows.Forms.DataGridTableStyle
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtcname)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.cmbcid)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.btncancel)
        Me.Panel1.Controls.Add(Me.DataGrid1)
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(554, 258)
        Me.Panel1.TabIndex = 0
        '
        'txtcname
        '
        Me.txtcname.BackColor = System.Drawing.Color.White
        Me.txtcname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtcname.Enabled = False
        Me.txtcname.Location = New System.Drawing.Point(336, 14)
        Me.txtcname.Name = "txtcname"
        Me.txtcname.Size = New System.Drawing.Size(200, 21)
        Me.txtcname.TabIndex = 23
        Me.txtcname.Text = ""
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label7.Location = New System.Drawing.Point(226, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 18)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Customer Name"
        '
        'cmbcid
        '
        Me.cmbcid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbcid.ItemHeight = 13
        Me.cmbcid.Location = New System.Drawing.Point(93, 14)
        Me.cmbcid.Name = "cmbcid"
        Me.cmbcid.Size = New System.Drawing.Size(123, 21)
        Me.cmbcid.TabIndex = 21
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.White
        Me.Label8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label8.Location = New System.Drawing.Point(10, 17)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 18)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Customer ID"
        '
        'btncancel
        '
        Me.btncancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.btncancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.ForeColor = System.Drawing.Color.White
        Me.btncancel.Location = New System.Drawing.Point(8, 232)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.Size = New System.Drawing.Size(60, 18)
        Me.btncancel.TabIndex = 7
        Me.btncancel.Text = "Cancel"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 38)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(536, 190)
        Me.DataGrid1.TabIndex = 1
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1, Me.DataGridTableStyle2})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = ""
        Me.DataGridTableStyle1.ReadOnly = True
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Sheraz"
        Me.DataGridTextBoxColumn1.MappingName = ""
        Me.DataGridTextBoxColumn1.Width = 75
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "Ahmed"
        Me.DataGridTextBoxColumn2.MappingName = ""
        Me.DataGridTextBoxColumn2.Width = 75
        '
        'DataGridTableStyle2
        '
        Me.DataGridTableStyle2.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle2.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn3})
        Me.DataGridTableStyle2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle2.MappingName = ""
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "uyiui"
        Me.DataGridTextBoxColumn3.MappingName = ""
        Me.DataGridTextBoxColumn3.Width = 75
        '
        'CustomerActivity
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(562, 267)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustomerActivity"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Activity"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Sub populatingGrid(ByVal tempCid As Integer)

        Ds = ClsSales.GetCustomerActSummary(tempCid)

        DataGrid1.TableStyles.Clear()

        'Formating The DataGrid Bt Adding a DataGrid Table Style
        Dim DgridTstyle As New DataGridTableStyle
        DgridTstyle.GridColumnStyles.Clear()

        Dim TranNo As New DataGridBrowser.DataGridNoActiveCellColumn   'parentid
        Dim Type As New DataGridBrowser.DataGridNoActiveCellColumn      'InvTp
        Dim Details As New DataGridBrowser.DataGridNoActiveCellColumn   'InvDetails
        Dim InvDate As New DataGridBrowser.DataGridNoActiveCellColumn   'InvDate  
        Dim PayStatus As New DataGridBrowser.DataGridNoActiveCellColumn   'TransID
        Dim Amount As New DataGridBrowser.DataGridNoActiveCellColumn    'OSAmt
        Dim Debit As New DataGridBrowser.DataGridNoActiveCellColumn   'LDEbit
        Dim Credit As New DataGridBrowser.DataGridNoActiveCellColumn     'LCredit

        'Setting Default Properties for datagrid like backcolor,forecolor,fonts
        With DgridTstyle
            .HeaderFont = New System.Drawing.Font _
            ("Verdana", 8.25F, System.Drawing.FontStyle.Regular, _
             System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
            .AlternatingBackColor = Color.Lavender
            .PreferredRowHeight = 3
        End With

        With TranNo
            .HeaderText = "TNo"
            .MappingName = "parentid"
            .Width = 30
            .ReadOnly = True
        End With

        With Type
            .HeaderText = "Type"
            .MappingName = "invtp"
            .Width = 35
            .ReadOnly = True
        End With

        With Details
            .HeaderText = "Service Details"
            .MappingName = "invDetails"
            .Width = 170
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With InvDate
            .HeaderText = "Date"
            .MappingName = "invdate"
            .Width = 75
            .ReadOnly = True
            .TextBox.ReadOnly = True
        End With

        With PayStatus
            .HeaderText = ""
            .Width = 10
            .MappingName = "Paystatus"
        End With

        With Amount
            .HeaderText = "OSAmount" & " " & Chr(9)
            .MappingName = "osamt"
            .Width = 65
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .Format = "0.00"
        End With

        With Debit
            .HeaderText = "Debit" & " " & Chr(9)
            .Width = 65
            .MappingName = "Ldebit"
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .Format = "0.00"
        End With

        With Credit
            .HeaderText = "Credit" & " " & Chr(9)
            .Width = 65
            .MappingName = "LCredit"
            .Alignment = HorizontalAlignment.Right
            .TextBox.TextAlign = HorizontalAlignment.Right
            .Format = "0.00"
        End With

        'Add The GridColumnStyle to Datagrid Column Style
        With DgridTstyle.GridColumnStyles
            .Add(TranNo)
            .Add(Type)
            .Add(Details)
            .Add(InvDate)
            .Add(PayStatus)
            .Add(Amount)
            .Add(Debit)
            .Add(Credit)
        End With

        'Bind The Datagrid to Dataset
        If Ds.Tables(0).Rows.Count > 0 Then
            DataGrid1.DataSource = Ds.Tables(0)
            'Stoping Extra Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False
            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = Ds.Tables(0).TableName '("TblTransaction")
            DataGrid1.TableStyles.Add(DgridTstyle)
        Else
            'Binding DataGrid With Empty DataSource
            DataGrid1.DataSource = Ds.Tables(0)
            'Stoping Extra Row In DataGrid At The End Of The Row
            Dim Cm As CurrencyManager = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
            CType(Cm.List, DataView).AllowNew = False
            DgridTstyle.RowHeadersVisible = False
            DgridTstyle.MappingName = Ds.Tables(0).TableName
            DataGrid1.TableStyles.Add(DgridTstyle)
        End If
    End Sub

    Private Sub CustomerActivity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dr = ClsSales.LoadCustomerID
        While Dr.Read
            cmbcid.Items.Add(Dr.Item("Customerid"))
        End While
        populatingGrid(CustID)
    End Sub

    Private Sub cmbcid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbcid.SelectedIndexChanged

        CustID = (cmbcid.SelectedItem)
        Dr = ClsSales.GetCustInfo(CustID)
        Dr.Read()
        If IsDBNull(Dr("Customername")) = False Then
            txtcname.Text = Dr("Customername")
        End If
        Dr.Close()
        populatingGrid(CustID)

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
