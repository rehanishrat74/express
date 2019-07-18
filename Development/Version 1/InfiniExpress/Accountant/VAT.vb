Imports System
Imports InfiniExpress.BLL

Public Class VAT

    Inherits System.Windows.Forms.Form

#Region "Variable Declaration"
    Private GlbAccount As New ClassAccount()
    Private GlbGlobal As New ClassGlobal()
    Public VatCriteria As String, StrQ As String, tp As String
    Public BetTrans As Integer
    Private TranChk As Boolean
    Public BefTrans As Integer
    Public BolChk As Boolean, FILTER1 As String, FILTER2 As String, Query As String
    Public Result1 As Double, Result2 As Double, Total As Double
    Public Amount As Double
    Private Chk As Boolean
    Dim I As Integer
#End Region

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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents txtChk As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents chkbox As System.Windows.Forms.CheckBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnReconcile As System.Windows.Forms.Button
    Friend WithEvents btnCalculate As System.Windows.Forms.Button
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VAT))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.txtChk = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnReconcile = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.chkbox = New System.Windows.Forms.CheckBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.DataGrid2 = New System.Windows.Forms.DataGrid()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(104, 264)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = ""
        Me.TextBox1.Visible = False
        Me.TextBox1.WordWrap = False
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(400, 264)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.TabIndex = 9
        Me.TextBox2.Text = ""
        Me.TextBox2.Visible = False
        '
        'txtChk
        '
        Me.txtChk.Location = New System.Drawing.Point(240, 264)
        Me.txtChk.Name = "txtChk"
        Me.txtChk.TabIndex = 10
        Me.txtChk.Text = "FALSE"
        Me.txtChk.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnReconcile, Me.btnClose, Me.chkbox, Me.btnClear, Me.btnCalculate, Me.DataGrid2, Me.DataGrid1})
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(550, 168)
        Me.Panel1.TabIndex = 11
        '
        'btnReconcile
        '
        Me.btnReconcile.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnReconcile.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnReconcile.ForeColor = System.Drawing.Color.White
        Me.btnReconcile.Location = New System.Drawing.Point(330, 136)
        Me.btnReconcile.Name = "btnReconcile"
        Me.btnReconcile.Size = New System.Drawing.Size(68, 18)
        Me.btnReconcile.TabIndex = 3
        Me.btnReconcile.Text = "&Reconcile"
        Me.btnReconcile.Visible = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(482, 136)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(60, 18)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "&Close"
        '
        'chkbox
        '
        Me.chkbox.BackColor = System.Drawing.Color.White
        Me.chkbox.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.chkbox.Location = New System.Drawing.Point(78, 136)
        Me.chkbox.Name = "chkbox"
        Me.chkbox.Size = New System.Drawing.Size(136, 24)
        Me.chkbox.TabIndex = 2
        Me.chkbox.Text = "Include Reconciled"
        Me.chkbox.Visible = False
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnClear.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClear.ForeColor = System.Drawing.Color.White
        Me.btnClear.Location = New System.Drawing.Point(410, 136)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(60, 18)
        Me.btnClear.TabIndex = 4
        Me.btnClear.Text = "Clear"
        '
        'btnCalculate
        '
        Me.btnCalculate.BackColor = System.Drawing.Color.LightSlateGray
        Me.btnCalculate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCalculate.ForeColor = System.Drawing.Color.White
        Me.btnCalculate.Location = New System.Drawing.Point(6, 136)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(64, 18)
        Me.btnCalculate.TabIndex = 1
        Me.btnCalculate.Text = "Calculate"
        '
        'DataGrid2
        '
        Me.DataGrid2.AllowNavigation = False
        Me.DataGrid2.AllowSorting = False
        Me.DataGrid2.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.DataGrid2.CaptionText = "Value Added Tax Return"
        Me.DataGrid2.CaptionVisible = False
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(6, 8)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.ReadOnly = True
        Me.DataGrid2.Size = New System.Drawing.Size(536, 120)
        Me.DataGrid2.TabIndex = 0
        '
        'DataGrid1
        '
        Me.DataGrid1.AllowNavigation = False
        Me.DataGrid1.AllowSorting = False
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.White
        Me.DataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.DataGrid1.CaptionText = "Value Added Tax Return"
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(6, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.Size = New System.Drawing.Size(536, 120)
        Me.DataGrid1.TabIndex = 13
        Me.DataGrid1.Visible = False
        '
        'VAT
        '
        Me.AutoScale = False
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(558, 177)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1, Me.txtChk, Me.TextBox2, Me.TextBox1})
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(0, 10)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VAT"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Value Added Tax Return"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub VAT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BindGrid()
    End Sub

    Sub BindGrid()

        Dim Cart As New DataTable()
        Dim CartDS As New DataSet()
        Cart.TableName = "Cart"
        Cart.Columns.Add(New DataColumn("Value Added Tax Return", GetType(String)))
        Cart.Columns.Add(New DataColumn("Amount", GetType(String)))

        Cart.Rows.Add(New Object() {"VAT payable on taxable supplies (output tax)", "0.00"})
        Cart.Rows.Add(New Object() {"VAT claimable on purchases (input tax)", "0.00"})
        Cart.Rows.Add(New Object() {"VAT payable to or reclaimable from HM Customs & Excise", "0.00"})
        Cart.Rows.Add(New Object() {"Total value of sales excluding VAT", "0.00"})
        Cart.Rows.Add(New Object() {"Total value of purchases excluding VAT", "0.00"})
        CartDS.Tables.Add(Cart)

        Dim ts As New DataGridTableStyle(), cs As DataGridTextBoxColumn

        ' Add the custom column style.
        cs = New DataGridTextBoxColumn()
        cs.Width = (2 / 3) * DataGrid1.Width
        cs.MappingName = "Value Added Tax Return"        ' Map to decimal column.
        cs.HeaderText = "Value Added Tax Return"
        ts.GridColumnStyles.Add(cs)

        ' Add the standard column style.
        cs = New DataGridTextBoxColumn()
        cs.Width = (1 / 3) * DataGrid1.Width - 7
        cs.MappingName = "Amount"          ' Map to integer column.
        cs.HeaderText = "Amount " & Chr(9)
        cs.Format = "0.00"
        cs.Alignment = HorizontalAlignment.Right
        cs.TextBox.TextAlign = HorizontalAlignment.Right
        ts.GridColumnStyles.Add(cs)

        ts.MappingName = "Cart"     ' Map table style to TestTable.
        ts.RowHeadersVisible = False
        DataGrid1.TableStyles.Add(ts)
        DataGrid1.DataSource = CartDS
        DataGrid1.DataMember = "Cart"

        'To delete the extra row
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid1.DataSource, DataGrid1.DataMember), CurrencyManager)
        CType(cm.List, DataView).AllowNew = False

        DataGrid1.Enabled = False
        DataGrid1.Visible = True
        DataGrid2.Visible = False

    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnReconcile_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReconcile.Click


        If (MessageBox.Show("Do you wish to update the transaction status as reconciled for VAT?", "VAT Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK) Then

            If (txtChk.Text = "FALSE") Or (txtChk.Text = "TRUE" And BetTrans >= 0) Then
                'If txtChk.Text = "FALSE" Then
                '    Query = " SELECT TBLTRANSACTION.Parentid,TBLTRANSACTION.Invtp,TBLTRANSACTION.v From TBLTRANSACTION " & _
                '            " WHERE (TBLTRANSACTION.V = 'N')"
                '    '" WHERE (TBLTRANSACTION.V <> '-')"
                'ElseIf txtChk.Text = "TRUE" And BetTrans > 0 Then
                '    Query = " SELECT TBLTRANSACTION.Parentid,TBLTRANSACTION.Invtp,TBLTRANSACTION.v From TBLTRANSACTION " & _
                '            " WHERE (TBLTRANSACTION.V = 'N')"
                '    '" WHERE (TBLTRANSACTION.V <> '-')"
                'ElseIf txtChk.Text = "TRUE" And BetTrans = 0 Then
                '    Query = " SELECT TBLTRANSACTION.Parentid,TBLTRANSACTION.Invtp,TBLTRANSACTION.v From TBLTRANSACTION " & _
                '            " WHERE (TBLTRANSACTION.V = 'N')"
                '    '" WHERE (TBLTRANSACTION.V <> '-')"
                GlbAccount.ReconcileVAT()
                GlbGlobal.ExecuteCommand()
            End If

            btnReconcile.Visible = False

            'GlbAccount.ReconcileTrans(Query)

            DataGrid2.Visible = False
            DataGrid1.Visible = True
            DataGrid1.Enabled = False
            chkbox.Checked = False

        End If

    End Sub

    Private Sub btnCalculate_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        Dim J As Integer

        If TextBox1.Text <= TextBox2.Text Then

            btnReconcile.Visible = True
            chkbox.Visible = True
            Dim objDataView As DataView
            Dim objDT As New DataTable("CalculatedVAT")
            objDataView = CalculateVAT(TextBox1.Text, TextBox2.Text, chkbox.Checked)

            Dim ts As New DataGridTableStyle(), cs As DataGridTextBoxColumn

            ' Add the custom column style.
            cs = New DataGridTextBoxColumn()
            cs.Width = (2 / 3) * DataGrid1.Width
            cs.MappingName = "VAT"        ' Map to vat column.
            cs.HeaderText = "Value Added Tax Return"
            ts.GridColumnStyles.Add(cs)

            cs = New DataGridTextBoxColumn()
            cs.Width = (1 / 3) * DataGrid1.Width - 7
            cs.MappingName = "Amount"          ' Map to amount column.
            cs.HeaderText = "Amount"
            cs.Format = "0.00"
            cs.Alignment = HorizontalAlignment.Right
            cs.TextBox.TextAlign = HorizontalAlignment.Right
            ts.GridColumnStyles.Add(cs)
            ts.MappingName = objDataView.Table.TableName     ' Map table style to TestTable.
            ts.RowHeadersVisible = False
            DataGrid2.TableStyles.Add(ts)
            DataGrid2.DataSource = objDataView

            Dim cm As CurrencyManager
            cm = CType(Me.BindingContext(DataGrid2.DataSource, DataGrid2.DataMember), CurrencyManager)
            CType(cm.List, DataView).AllowNew = False

            DataGrid1.Visible = False
            DataGrid2.Visible = True
            DataGrid2.Enabled = False

        Else
            MessageBox.Show("Please enter an appropriate date range to calculate the VAT payable.   ", "Infini Accounts Express", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DataGrid1.Visible = True
            DataGrid2.Visible = False
            TextBox1.Text = Format(Now.Date, "dd/MM/yyyy")
            TextBox2.Text = Format(Now.Date, "dd/MM/yyyy")
            BindGrid()
        End If

    End Sub

    Private Sub btnClear_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        DataGrid1.Visible = True
        DataGrid2.Visible = False
        btnReconcile.Visible = False
        chkbox.Checked = False
        TextBox1.Text = Format(Now.Date, "dd/MM/yyyy")
        TextBox2.Text = Format(Now.Date, "dd/MM/yyyy")
        DataGrid1.TableStyles.Clear()
        BindGrid()
    End Sub

#Region "CALCULATE VAT"

    Public Function CalculateVAT(ByVal Date1 As String, ByVal Date2 As String, ByVal ChkBox As Boolean) As DataView

        Dim tempD1 As String = Date1, tempD2 As String = Date2, MsgStr As String
        Dim tempBol As Boolean = ChkBox
        Dim Res1 As Double, Res2 As Double
        Dim I As Integer, j As Integer
        Dim Str1 As String, Str2 As String
        Dim Dt As New DataTable()

        Dim DC1 As New DataColumn("VAT", GetType(String))
        Dim DC2 As New DataColumn("Group", GetType(String))
        Dim DC3 As New DataColumn("Amount", GetType(Double))

        Dt.Columns.Add(DC1)
        Dt.Columns.Add(DC2)
        Dt.Columns.Add(DC3)

        Dim Dr As DataRow

        If ChkBox = True Then
            VatCriteria = "<>'-'"
        Else
            VatCriteria = "='N'"
        End If

        BetTrans = GlbAccount.VATBetweenTranscation(Date1, Date2, VatCriteria)
        BefTrans = GlbAccount.VATBeforeTranscation(Date1, VatCriteria)

        If BetTrans > 0 And BefTrans > 0 Or BetTrans = 0 And BefTrans > 0 Then
            txtChk.Text = "TRUE"
            For I = 1 To 5
                DetailQry(I)

                Res1 = GlbAccount.VATQUERY1(FILTER1)
                Res2 = GlbAccount.VATQUERY2(FILTER2)
                Total = ((Res1) - (Res2))

                Dr = Dt.NewRow()
                Dr("VAT") = StrQ
                Dr("Group") = I
                Dr("Amount") = Format(Total, "0.00")
                Dt.Rows.Add(Dr)
            Next

            Dim Dv As DataView = New DataView(Dt)
            Return Dv
        Else
            Dim Rs1 As Double, Rs2 As Double
            For I = 1 To 5

                DetailQry(I)
                Rs1 = GlbAccount.VATQUERY1(FILTER1)
                Rs2 = GlbAccount.VATQUERY2(FILTER2)
                Dr = Dt.NewRow()
                Dr("VAT") = StrQ
                Dr("Group") = I
                Dr("Amount") = "0.00"
                Dt.Rows.Add(Dr)
            Next

            Dim Dv As DataView = New DataView(Dt)
            Return Dv

        End If

    End Function

    Private Function DetailQry(ByVal TempI As Integer)

        Select Case TempI

            Case 1
                FILTER1 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='SI') Or (TBLTRANSACTION.Invtp='BR') Or (TBLTRANSACTION.Invtp='VR') Or (TBLTRANSACTION.Invtp='CR')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "

                FILTER2 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE TBLTRANSACTION.Invtp='SC' AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "
                StrQ = "VAT due in this period on sales"
            Case 2
                FILTER1 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTransaction " & _
                          " WHERE ((TBLTRANSACTION.Invtp='PI') Or (TBLTRANSACTION.Invtp='BP') Or (TBLTRANSACTION.Invtp='CP') Or (TBLTRANSACTION.Invtp='VP')) AND TBLTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "

                FILTER2 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTransaction " & _
                          " WHERE TBLTRANSACTION.Invtp='PC' AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "
                StrQ = "VAT reclaimed in this period on purchases"
            Case 3
                FILTER1 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='SI') Or (TBLTRANSACTION.Invtp='PC') Or (TBLTRANSACTION.Invtp='BR') Or (TBLTRANSACTION.Invtp='VR') Or (TBLTRANSACTION.Invtp='CR')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "

                FILTER2 = " SELECT Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='SC') Or (TBLTRANSACTION.Invtp='PI') Or (TBLTRANSACTION.Invtp='BP') Or (TBLTRANSACTION.Invtp='CP') Or (TBLTRANSACTION.Invtp='VP')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "
                StrQ = "VAT payable to or reclaimable from HM Customs & Excise"
            Case 4
                FILTER1 = " SELECT Sum(TBlTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='SI') Or (TBLTRANSACTION.Invtp='BR') Or (TBLTRANSACTION.Invtp='VR') Or (TBLTRANSACTION.Invtp='CR') Or (TBLTRANSACTION.Invtp='JC')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "

                FILTER2 = " SELECT Sum(TBlTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='SC') Or (TBLTRANSACTION.Invtp='SD')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "
                StrQ = "Total value of sales excluding VAT"

            Case 5
                FILTER1 = " SELECT Sum(TBLTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='PI') Or (TBLTRANSACTION.Invtp='BP') Or (TBLTRANSACTION.Invtp='CP') Or (TBLTRANSACTION.Invtp='VP') Or (TBLTRANSACTION.Invtp='JD')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "

                FILTER2 = " SELECT Sum(TBLTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                          " WHERE ((TBLTRANSACTION.Invtp='PC') Or (TBLTRANSACTION.Invtp='PD')) AND TBlTRANSACTION.v" & VatCriteria & "" & _
                          " GROUP BY TBLTRANSACTION.Invtp "
                StrQ = "Total value of purchases excluding VAT"

        End Select

    End Function

#End Region

#Region "CALCULATE RECONCILE"

    Private Function Check_Tp(ByVal tp As String) As Integer

        Dim CurrCol As Integer

        If tp = "SI" Or tp = "PI" Then
            CurrCol = 1
        ElseIf tp = "SC" Or tp = "PC" Then
            CurrCol = 2
        ElseIf tp = "BR" Or tp = "VR" Or tp = "CR" Or tp = "BP" Or tp = "VP" Or tp = "CP" Then
            CurrCol = 3
        ElseIf tp = "JC" Or tp = "JD" Then
            CurrCol = 4
        End If

        Return CurrCol

    End Function

    Private Function Qry(ByVal tempVar As Integer) As String

        Dim Dquery As String

        Select Case tempVar
            'when Date Datatype is implemented we have to change the query
        Case 1
                Dquery = "SELECT TBLTRANSACTION.TaxID, Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                         "WHERE TBLTRANSACTION.Invtp ='SI' Or TBLTRANSACTION.Invtp ='BR' Or TBLTRANSACTION.Invtp ='VR' Or TBLTRANSACTION.Invtp ='CR' Or TBLTRANSACTION.Invtp ='SC' AND TBlTRANSACTION.v='" & VatCriteria & "'" & _
                         "GROUP BY TBLTRANSACTION.TaxID, TBLTRANSACTION.Invtp " & _
                         "ORDER BY TBLTRANSACTION.TaxID"
            Case 2
                Dquery = "SELECT TBLTRANSACTION.TaxID, Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                         "WHERE TBLTRANSACTION.Invtp='PI' Or TBLTRANSACTION.Invtp='BP' Or TBLTRANSACTION.Invtp='CP' Or TBLTRANSACTION.Invtp='VP' Or TBLTRANSACTION.Invtp='PC' AND TBlTRANSACTION.v='" & VatCriteria & "'" & _
                         "GROUP BY TBLTRANSACTION.TaxID, TBLTRANSACTION.Invtp " & _
                         "ORDER BY TBLTRANSACTION.TaxID"
            Case 3
                Dquery = "SELECT TBLTRANSACTION.TaxID, Sum(TBLTRANSACTION.Invvat) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                         "WHERE TBLTRANSACTION.Invtp='SI' Or TBLTRANSACTION.Invtp='PC' Or TBLTRANSACTION.Invtp='BR' Or TBLTRANSACTION.Invtp='VR' Or TBLTRANSACTION.Invtp='CR' Or TBLTRANSACTION.Invtp='SC' Or TBLTRANSACTION.Invtp='PI' Or TBLTRANSACTION.Invtp='BP' Or TBLTRANSACTION.Invtp='CP' Or TBLTRANSACTION.Invtp='VP' AND TBlTRANSACTION.v='" & VatCriteria & "'" & _
                         "GROUP BY TBLTRANSACTION.TaxID, TBLTRANSACTION.Invtp " & _
                         "ORDER BY TBLTRANSACTION.TaxID"
            Case 4
                Dquery = "SELECT TBLTRANSACTION.TaxID, Sum(TBLTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                         "WHERE TBLTRANSACTION.Invtp='SI' Or TBLTRANSACTION.Invtp='BR' Or TBLTRANSACTION.Invtp='VR' Or TBLTRANSACTION.Invtp='CR' Or TBLTRANSACTION.Invtp='JC' Or TBLTRANSACTION.Invtp='SC' Or TBLTRANSACTION.Invtp='SD' AND TBlTRANSACTION.v='" & VatCriteria & "'" & _
                         "GROUP BY TBLTRANSACTION.TaxID, TBLTRANSACTION.Invtp " & _
                         "ORDER BY TBLTRANSACTION.TaxID"
            Case 5
                Dquery = "SELECT TBLTRANSACTION.TaxID, Sum(TBLTRANSACTION.Invnet) AS SumOfvat, TBLTRANSACTION.Invtp From TBLTRANSACTION " & _
                         "WHERE TBLTRANSACTION.Invtp='PI' Or TBLTRANSACTION.Invtp='BP' Or TBLTRANSACTION.Invtp='CP' Or TBLTRANSACTION.Invtp='VP' Or TBLTRANSACTION.Invtp='JD' Or TBLTRANSACTION.Invtp='PC' Or TBLTRANSACTION.Invtp='PD' AND TBlTRANSACTION.v='" & VatCriteria & "'" & _
                         "GROUP BY TBLTRANSACTION.TaxID, TBLTRANSACTION.Invtp " & _
                         "ORDER BY TBLTRANSACTION.TaxID"
        End Select

        Return Dquery

    End Function

#End Region

End Class
