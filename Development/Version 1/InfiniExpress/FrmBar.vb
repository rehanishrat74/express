Imports System
Imports System.Data
Imports InfiniExpress.BLL

Public Class FrmBar

    Inherits System.Windows.Forms.Form
    Public TblFlag As Boolean

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
    Protected WithEvents CmdSynch As System.Windows.Forms.Button
    Protected WithEvents CmdCancel As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmBar))
        Me.CmdSynch = New System.Windows.Forms.Button
        Me.CmdCancel = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CmdSynch
        '
        Me.CmdSynch.BackColor = System.Drawing.Color.LightSlateGray
        Me.CmdSynch.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CmdSynch.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdSynch.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CmdSynch.Location = New System.Drawing.Point(88, 72)
        Me.CmdSynch.Name = "CmdSynch"
        Me.CmdSynch.Size = New System.Drawing.Size(96, 23)
        Me.CmdSynch.TabIndex = 0
        Me.CmdSynch.Text = "&Synchronize"
        '
        'CmdCancel
        '
        Me.CmdCancel.BackColor = System.Drawing.Color.LightSlateGray
        Me.CmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CmdCancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdCancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CmdCancel.Location = New System.Drawing.Point(192, 72)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.Size = New System.Drawing.Size(96, 23)
        Me.CmdCancel.TabIndex = 1
        Me.CmdCancel.Text = "&Cancel"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(280, 56)
        Me.Panel1.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(261, 40)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Synchronizes data from the account centre. http:\\www.accountscentre.com"
        '
        'FrmBar
        '
        Me.AcceptButton = Me.CmdSynch
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(102, Byte), CType(153, Byte), CType(204, Byte))
        Me.ClientSize = New System.Drawing.Size(292, 103)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CmdCancel)
        Me.Controls.Add(Me.CmdSynch)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmBar"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Synchronization"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub SynchronizeData()

        Dim dtds As DataSet, tempDs As DataSet
        Dim wds As DataSet
        Dim dtable As New DataTable
        Dim SynFlag As Boolean
        Dim IEDService As New InfiniExpress.InfiniWebService.InfiniService
        Dim GlbSynchronize As New ClassSynchronize
        Dim Rcount As Integer
        Dim r As Integer, i As Integer

        Try
            ' Synchronize Data From InfiniAccouns Express(DESKTOP) To
            ' InfiniAccounts Express(WEB BASE)

            'STEP 1
            'Get un Synchronize Miscellaneous Data from DeskTop Application
            dtds = GlbSynchronize.GetMiscellaneousDataFromDeskTop

            TblFlag = ChkExistingTable(dtds)

            If TblFlag = True Then
                ''Send Data desktop to Web for Synchronize
                SynFlag = IEDService.SynchronizeFromWeb(dtds)

                If SynFlag = True Then
                    'Update Syn Flag = Y for DestTop
                    GlbSynchronize.UpdateSynFlagDeskTopData(dtds)
                End If

            End If
            dtds.Dispose()

            'STEP 2
            'Get Unsynchronize Of Customer & Vendor Data In Chunks
            tempDs = GlbSynchronize.GetUnsynchronizeCVData
            r = RowCount(tempDs)

            For i = 1 To r
                If r > 1 Then
                    dtds = New DataSet
                    tempDs = New DataSet

                    dtds = GlbSynchronize.GetCustomerVendorDataFromDeskTop

                    TblFlag = ChkExistingTable(dtds)

                    If TblFlag = True Then
                        ''Send Data desktop to Web for Synchronize
                        SynFlag = IEDService.SynchronizeFromWeb(dtds)

                        If SynFlag = True Then
                            'Update Syn Flag = Y for DestTop
                            GlbSynchronize.UpdateSynFlagDeskTopDataForCV(dtds)
                        End If

                    End If
                    dtds.Dispose()

                    tempDs = GlbSynchronize.GetUnsynchronizeCVData
                    r = RowCount(tempDs)
                Else
                    tempDs.Dispose()
                    Exit For
                End If
            Next

            'STEP 3
            'Get Unsynchronize Data From Transacation ,Ledger & Outstanding
            tempDs = GlbSynchronize.GetUnsynchronizeTLOData
            r = RowCount(tempDs)

            For i = 1 To r
                If r > 1 Then
                    dtds = New DataSet
                    tempDs = New DataSet

                    'Get un Synchronize Data from DeskTop Application In Chunks
                    dtds = GlbSynchronize.GetCustomerVendorTransactionDataFromDeskTop

                    TblFlag = ChkExistingTable(dtds)

                    If TblFlag = True Then
                        ''Send Data desktop to Web for Synchronize
                        SynFlag = IEDService.SynchronizeFromWebTransaction(dtds)

                        If SynFlag = True Then
                            'Update Syn Flag = Y for DestTop
                            GlbSynchronize.UpdateSynFlagDeskTopDataForT(dtds)
                        End If

                    End If
                    dtds.Dispose()
                    tempDs = GlbSynchronize.GetUnsynchronizeTLOData
                    r = RowCount(tempDs)
                Else
                    tempDs.Dispose()
                    Exit For
                End If
            Next

            '###########################################################
            ' Synchronize Data From InfiniAccouns Express(WEB BASE SYSTEM) To
            ' InfiniAccounts Express(DESKTOP SYSTEM)
            '###########################################################

            'STEP 1 
            'Get DateTime For Web Synchronization
            wds = GlbSynchronize.GetSynchronizeTableData()

            '''Get unSynchronize Data from Web Service
            dtds = New DataSet
            dtds = IEDService.SynchronizeFromDeskTop(wds)

            'Manipulation Data in DeskTop Application
            SynFlag = GlbSynchronize.DeskTopSynchronizeData(dtds)

            If SynFlag = True Then
                'GlbSynchronize.UpdateAfterValidation() 'Update Synchronize DateTime
                ''''Update Syn Flag = Y for Web
                'SynFlag = IEDService.UpdateSynchronizeFlag(wds)
            End If


            ''STEP 2
            wds = New DataSet
            wds = GlbSynchronize.GetSynchronizeTableData()

            tempDs = IEDService.GetUnSynCVDate(wds)
            r = RowCount(tempDs)

            If r > 0 Then
                For i = 1 To r
                    If r > 0 Then
                        dtds = New DataSet
                        dtds = IEDService.SynchronizeCustomerAndVendor(wds)

                        SynFlag = GlbSynchronize.DeskTopSynchronizeWithCVWEB(dtds)

                        If SynFlag = True Then
                            'Update Synchronize DateTime
                            GlbSynchronize.UpdateAfterValidation()
                            'Update Syn Flag = Y for Web
                            dtds.Merge(wds)
                            SynFlag = IEDService.UpdateSynchronizeFlagCVWebData(dtds)
                        End If
                        dtds.Dispose()
                        tempDs = IEDService.GetUnSynCVDate(wds)
                        r = RowCount(tempDs)
                    Else
                        Exit For
                    End If
                Next
            Else
                tempDs.Dispose()
                wds.Dispose()
            End If

            'STEP 3
            wds.Dispose()
            wds = GlbSynchronize.GetSynchronizeTableData()

            'Get Unsynchronize All Transactions Data From Web
            tempDs = IEDService.GetTotalTOLData(wds)
            r = RowCount(tempDs)

            If r > 0 Then
                For i = 1 To r
                    If r > 0 Then
                        'Get Unsynchronize Data From Web in Chunks
                        dtds = New DataSet
                        dtds = IEDService.SynchronizeTransacation(wds)

                        dtds.WriteXml("C:\SAk.xml")
                        'Insert Data Fetch From Web
                        SynFlag = GlbSynchronize.DeskTopSynchronizeWithTransactionWebData(dtds)

                        If SynFlag = True Then
                            'Update Synchronize DateTime
                            'GlbSynchronize.UpdateAfterValidation()
                            'Update Syn Flag = Y for Web
                            dtds.Merge(wds)
                            SynFlag = IEDService.UpdateSynchronizeFlagTWebData(dtds)
                        End If
                        dtds.Dispose()
                        tempDs = IEDService.GetTotalTOLData(wds)
                        r = RowCount(tempDs)
                    Else
                        Exit For
                    End If
                Next
            Else
                tempDs.Dispose()
                wds.Dispose()
            End If

            dtds = Nothing
            wds = Nothing
            MsgBox("Synchronization has been completed.", MsgBoxStyle.Information, "Infini Accounts Express")

        Catch e1 As Exception
            dtds = Nothing
            wds = Nothing
            IEDService = New InfiniExpress.InfiniWebService.InfiniService
            MsgBox("Data transfer failed...  " + e1.Message + "", MsgBoxStyle.Critical, "Infini Accounts Express")
        End Try


    End Sub

    Private Sub CmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCancel.Click
        Me.Dispose(True)
    End Sub

    Private Sub CmdSynch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSynch.Click
        SynchronizeData()
    End Sub

    Private Function ChkExistingTable(ByVal _tempDs As DataSet) As Boolean
        Dim myTable As DataTable
        Dim myRow As DataRow
        Dim th As String, BoolChk As Boolean

        For Each myTable In _tempDs.Tables
            For Each myRow In myTable.Rows

                th = myTable.TableName

                If th <> "TblSynchronize" Then
                    BoolChk = True
                    Exit For
                End If

            Next myRow
        Next myTable

        Return BoolChk
    End Function
    Private Function RowCount(ByVal tempDS As DataSet) As Integer
        Dim result As Integer
        Dim dt As DataTable
        For Each dt In tempDS.Tables
            result += dt.Rows.Count
        Next
        Return result
    End Function

End Class
