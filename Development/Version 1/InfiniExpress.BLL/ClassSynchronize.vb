'Option Strict On
Imports System
Imports System.Data
Imports System.Xml
Imports InfiniExpress.DAL
Imports System.IO
Imports Path = System.IO.Path

Public Class ClassSynchronize

    Public Shared cmdCollection As New Collection
    Dim GlbGlobal As New ClassGlobal
    Dim UserId As Integer
    Dim TranDateTime As String

    Public Function UpdateUserSynchronizeInfo(ByVal CustId As Integer) As Boolean

        Dim cmd As New InfiniCommand("UPDATESYNCHRONIZEID")
        cmd.AddParameter("@CustId", CustId)
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

        Return GlbGlobal.ExecuteCommand()

    End Function

    Public Sub UpdateAfterValidation()

        Dim cmd As New InfiniCommand("UPDATESYNCHRONIZEDATETIME")
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

        GlbGlobal.ExecuteCommand()

    End Sub

    Public Function GetMiscellaneousDataFromDeskTop() As DataSet

        Dim ts As DataSet
        Dim dr As IDataReader

        Dim cmd0 As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd0)

        'Step 1
        Dim cmd1 As New InfiniCommand("TblBank")
        cmdCollection.Add(cmd1)

        Dim cmd2 As New InfiniCommand("TblCompany")
        cmdCollection.Add(cmd2)

        Dim cmd4 As New InfiniCommand("TblFinancialYear")
        cmdCollection.Add(cmd4)

        Dim cmd7 As New InfiniCommand("TblTaxIds")
        cmdCollection.Add(cmd7)

        ts = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ts

    End Function

    Public Function GetCustomerVendorDataFromDeskTop() As DataSet
        Dim ts As DataSet

        Dim cmd0 As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd0)

        Dim cmd3 As New InfiniCommand("TblCustomer")
        cmdCollection.Add(cmd3)

        Dim cmd9 As New InfiniCommand("TblVendor")
        cmdCollection.Add(cmd9)

        ts = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ts

    End Function

    Public Function GetCustomerVendorTransactionDataFromDeskTop() As DataSet
        Dim ts As DataSet
        Dim _ds As IDataReader

        'Get Top 10 ParentID
        Dim cmd101 As New InfiniCommand("GETTRANSACTIONPARENTID")

        _ds = AccessHelper.ExecuteReader(cmd101)

        Do While _ds.Read = True

            Dim cmd0 As New InfiniCommand("TblSynchronize")
            cmdCollection.Add(cmd0)

            'Step 3
            Dim cmd8 As New InfiniCommand("TBLTRANSACTION10")
            cmd8.AddParameter("Parentid", _ds.Item("ParentID"))
            cmdCollection.Add(cmd8)

            Dim cmd5 As New InfiniCommand("TBLLEDGER10")
            cmd5.AddParameter("Parentid", _ds.Item("ParentID"))
            cmdCollection.Add(cmd5)

            Dim cmd6 As New InfiniCommand("TblOutStanding")
            cmdCollection.Add(cmd6)

        Loop

        'Dim cmd10 As New InfiniCommand("TblCreditCardInfo")
        'cmdCollection.Add(cmd10)

        ts = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ts

    End Function


    Public Function GetSynchronizeTableData() As DataSet

        Dim wds As DataSet
        Dim cmd As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd)

        wds = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)

        cmdCollection = New Collection

        Return wds

    End Function

    Public Function DeskTopSynchronizeData(ByVal DTds As DataSet) As Boolean

        Try
            TblBankTableSynchronize(DTds.Tables("TblBank"))

            TblCompanyTableSynchronize(DTds.Tables("TblCompany"))

            TblFinancialYearTableSynchronize(DTds.Tables("TblFinancialYear"))

            TblTaxIdsTableSynchronize(DTds.Tables("TblTaxIds"))

            ExecuteCommand() 'For Synchronize Updates Tables

        Catch sqle As DataException
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & sqle.Message)
            Return False

        Catch e As Exception
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & e.Message)
            Return False

        End Try

        Return True

    End Function

    Public Function DeskTopSynchronizeWithCVWEB(ByVal DTds As DataSet) As Boolean

        Try

            TblCustomerTableSynchronize(DTds.Tables("TblCustomer"))

            ExecuteSynchornizeCommand() 'Synchornize TblCustomer Table

            TblVendorTableSynchronize(DTds.Tables("TblVendor"))

            ExecuteSynchornizeCommand() 'Synchornize TblVendor Table

        Catch sqle As DataException
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & sqle.Message)
            Return False

        Catch e As Exception
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & e.Message)
            Return False

        End Try

        Return True

    End Function
    Public Function DeskTopSynchronizeWithTransactionWebData(ByVal DTds As DataSet) As Boolean

        Try
            TblTransactionTableSynchronize(DTds.Tables("TblTransaction"))

            ExecuteSynchornizeCommand() 'Synchornize TblTransaction Table

            TblLedgerTableSynchronize(DTds.Tables("TblLedger"))

            ExecuteSynchornizeCommand() 'Synchornize TblLedger Table

            TblOutStandingTableSynchronize(DTds.Tables("TblOutStanding"))

            ExecuteSynchornizeCommand() 'Synchornize TblOutStanding Table

        Catch sqle As DataException
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & sqle.Message)
            Return False

        Catch e As Exception
            cmdCollection = New Collection
            Throw New Exception("Synchronize Method : " & e.Message)
            Return False

        End Try

        Return True

    End Function



    Private Function ExecuteCommand() As Boolean
        Return AccessHelper.ExecuteCollection(cmdCollection)
        cmdCollection = New Collection
    End Function

    Private Function ExecuteSynchornizeCommand() As Boolean
        Return AccessHelper.ExecuteSynchronizeCollect(cmdCollection)
        cmdCollection = New Collection
    End Function

    Private Sub TblSynchronizeTableSynchronize(ByVal tb As DataTable)

        Dim dr As DataRow

        For Each dr In tb.Rows

            TranDateTime = CDate(dr("SynDateTime"))
            UserId = CInt(dr("UserId"))

        Next

    End Sub

    Private Sub TblBankTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim cmd As New InfiniCommand("TblBank")

            For Each dr In tb.Rows

                cmd.AddParameter("@BankId", CStr(dr("BankId")))
                cmd.AddParameter("@BankName", CStr(dr("BankName")))
                cmd.AddParameter("@Street", CStr(dr("Street")))
                cmd.AddParameter("@Town", CStr(dr("Town")))
                cmd.AddParameter("@Country", CStr(dr("Country")))
                cmd.AddParameter("@PostalCode", CStr(dr("PostalCode")))
                cmd.AddParameter("@AccountName", CStr(dr("AccountName")))
                cmd.AddParameter("@AccountNumber", CStr(dr("AccountNumber")))
                cmd.AddParameter("@SortCode", CStr(dr("SortCode")))
                cmdCollection.Add(cmd)

            Next

        End If

    End Sub

    Private Sub TblCompanyTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim cmd1 As New InfiniCommand("TblCompany")

            For Each dr In tb.Rows

                cmd1.AddParameter("@Name", CStr(dr("Name")))
                cmd1.AddParameter("@Street", CStr(dr("Street")))
                cmd1.AddParameter("@City", CStr(dr("City")))
                cmd1.AddParameter("@PostalCode", CStr(dr("PostalCode")))
                cmd1.AddParameter("@Country", CStr(dr("Country")))
                cmd1.AddParameter("@Telephone", CStr(dr("Telephone")))
                cmd1.AddParameter("@Fax", CStr(dr("Fax")))
                cmd1.AddParameter("@Email", CStr(dr("Email")))
                cmd1.AddParameter("@VatNo", CStr(dr("VatNo")))
                cmd1.AddParameter("@LogoName", CStr(dr("LogoName")))
                cmdCollection.Add(cmd1)

            Next

        End If

    End Sub

    Private Sub TblFinancialYearTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow

            For Each dr In tb.Rows

                Dim cmd2 As New InfiniCommand("TblFinancialYear")
                cmd2.AddParameter("@MyNo", CInt(dr("MyNo")))
                cmd2.AddParameter("@FMonth", CStr(dr("FMonth")))
                cmd2.AddParameter("@FYear", CStr(dr("FYear")))
                cmdCollection.Add(cmd2)

            Next

        End If

    End Sub

    Private Sub TblTaxIdsTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow

            For Each dr In tb.Rows

                Dim cmd3 As New InfiniCommand("TblTaxIds")
                cmd3.AddParameter("@TaxId", CStr(dr("TaxId")))
                cmd3.AddParameter("@TaxRate", CDbl(dr("TaxRate")))
                cmdCollection.Add(cmd3)

            Next

        End If

    End Sub

    Private Sub TblCustomerTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim CustId As String, ChFlag As String
            Dim Reader As IDataReader

            For Each dr In tb.Rows

                CustId = CStr(dr("Cart_Customer_UId"))

                Dim chcmd As New InfiniCommand("SynTblCustomer")
                chcmd.AddParameter("@CustomerId", CustId)

                Reader = ClassGlobal.GetDataReader(chcmd)
                If Reader.Read Then
                    ChFlag = "Y"
                Else
                    ChFlag = "N"
                End If
                Reader.Close()

                Dim cmd4 As New InfiniCommand("TblCustomer")
                cmd4.AddParameter("@CustomerId", CustId)
                cmd4.AddParameter("@CustomerName", CStr(dr("Name")))
                cmd4.AddParameter("@Street", CStr(dr("Street")))
                cmd4.AddParameter("@Town", CStr(dr("Town")))
                cmd4.AddParameter("@Country", CStr(dr("Country")))
                cmd4.AddParameter("@PostalCode", CStr(dr("PostCode")))
                cmd4.AddParameter("@ContactName", CStr(dr("Cont_Name")))
                cmd4.AddParameter("@Email", CStr(dr("Cart_Customer_Email")))
                cmd4.AddParameter("@Telephone", CStr(dr("Telephone")))
                cmd4.AddParameter("@Fax", CStr(dr("Fax")))
                cmd4.AddParameter("@Vat", CStr(dr("Vat")))
                cmd4.AddParameter("@CreditLimit", CDbl(dr("CreditLimit")))
                cmd4.AddParameter("@DFlag", CStr(dr("DFlag")))
                cmd4.AddParameter("@Update", ChFlag)
                cmdCollection.Add(cmd4)

            Next

        End If

    End Sub

    Private Sub TblVendorTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim VendId As String, ChFlag As String
            Dim Reader As IDataReader

            For Each dr In tb.Rows

                VendId = CStr(dr("VendorId"))

                Dim chcmd As New InfiniCommand("SynTblVendor")
                chcmd.AddParameter("@VendorId", VendId)

                Reader = ClassGlobal.GetDataReader(chcmd)
                If Reader.Read Then
                    ChFlag = "Y"
                Else
                    ChFlag = "N"
                End If
                Reader.Close()

                Dim cmd5 As New InfiniCommand("TblVendor")
                cmd5.AddParameter("@VendorId", VendId)
                cmd5.AddParameter("@VendorName", CStr(dr("VendorName")))
                cmd5.AddParameter("@Street", CStr(dr("Street")))
                cmd5.AddParameter("@Town", CStr(dr("Town")))
                cmd5.AddParameter("@Country", CStr(dr("Country")))
                cmd5.AddParameter("@PostalCode", CStr(dr("PostalCode")))
                cmd5.AddParameter("@ContactName", CStr(dr("ContactName")))
                cmd5.AddParameter("@Email", CStr(dr("Email")))
                cmd5.AddParameter("@Telephone", CStr(dr("Telephone")))
                cmd5.AddParameter("@Fax", CStr(dr("Fax")))
                cmd5.AddParameter("@Vat", CStr(dr("Vat")))
                cmd5.AddParameter("@CreditLimit", CDbl(dr("CreditLimit")))
                cmd5.AddParameter("@DFlag", CStr(dr("DFlag")))
                cmd5.AddParameter("@Update", ChFlag)
                cmdCollection.Add(cmd5)

            Next

        End If

    End Sub

    Private Sub TblTransactionTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim ChFlag As String
            Dim Reader As IDataReader

            For Each dr In tb.Rows

                Dim chcmd As New InfiniCommand("SynTblTransaction")
                chcmd.AddParameter("@ParentId", CInt(dr("ParentId")))

                Reader = ClassGlobal.GetDataReader(chcmd)
                If Reader.Read Then
                    ChFlag = "Y"
                Else
                    ChFlag = "N"
                End If
                'Reader.Close()
                Reader.Dispose()

                Dim cmd8 As New InfiniCommand("TblTransaction")
                cmd8.AddParameter("@ParentId", CInt(dr("ParentId")))
                cmd8.AddParameter("@InvTP", CStr(dr("InvTP")))
                cmd8.AddParameter("@TaxId", CStr(dr("TaxId")))
                cmd8.AddParameter("@InvDate", CStr(dr("InvDate")))
                cmd8.AddParameter("@InvDetails", CStr(dr("InvDetails")))
                cmd8.AddParameter("@InvNet", CDbl(dr("InvNet")))
                cmd8.AddParameter("@InvVat", CDbl(dr("InvVat")))
                cmd8.AddParameter("@Refund", CStr(dr("Refund")))
                cmd8.AddParameter("@Mark", CStr(dr("Mark")))
                cmd8.AddParameter("@V", CStr(dr("V")))
                cmd8.AddParameter("@TranDate", CStr(dr("TranDate")))
                cmd8.AddParameter("@Group", CInt(dr("GroupNo")))
                cmd8.AddParameter("@Req", CInt(dr("rid")))
                cmd8.AddParameter("@Update", ChFlag)
                cmdCollection.Add(cmd8)

            Next

        End If

    End Sub

    Private Sub TblLedgerTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim ChFlag As String
            Dim Reader As IDataReader

            For Each dr In tb.Rows

                Dim chcmd As New InfiniCommand("SynTblLedger")
                chcmd.AddParameter("@MyAutoNo", CInt(dr("MyAutoNo")))

                Reader = ClassGlobal.GetDataReader(chcmd)
                If Reader.Read Then
                    ChFlag = "Y"
                Else
                    ChFlag = "N"
                End If
                Reader.Dispose()
                'Reader.Close()

                Dim cmd9 As New InfiniCommand("TblLedger")
                cmd9.AddParameter("@ParentId", CInt(dr("ParentId")))
                cmd9.AddParameter("@ChildId", CInt(dr("ChildId")))
                cmd9.AddParameter("@TransId", CStr(dr("TransId")))
                cmd9.AddParameter("@Debit", CDbl(dr("Debit")))
                cmd9.AddParameter("@Credit", CDbl(dr("Credit")))
                cmd9.AddParameter("@TranType", CStr(dr("TranType")))
                cmd9.AddParameter("@TranRef", CStr(dr("TranRef")))
                cmd9.AddParameter("@PayStatus", CStr(dr("PayStatus")))
                cmd9.AddParameter("@MyAutoNo", CInt(dr("MyAutoNo")))
                cmd9.AddParameter("@Update", ChFlag)
                cmdCollection.Add(cmd9)

            Next

        End If

    End Sub

    Private Sub TblOutStandingTableSynchronize(ByVal tb As DataTable)

        If Not tb Is Nothing Then

            Dim dr As DataRow
            Dim ChFlag As String
            Dim Reader As IDataReader

            For Each dr In tb.Rows

                Dim chcmd As New InfiniCommand("SynTblOutStanding")
                chcmd.AddParameter("@OSNo", CInt(dr("OSNo")))
                'chcmd.AddParameter("@ParentId", CInt(dr("ParentId")))
                'chcmd.AddParameter("@TransId", CStr(dr("TransId")))
                'chcmd.AddParameter("@TranType", CStr(dr("TranType")))
                Reader = ClassGlobal.GetDataReader(chcmd)
                If Reader.Read Then
                    ChFlag = "Y"
                Else
                    ChFlag = "N"
                End If
                'Reader.Close()
                Reader.Dispose()

                Dim cmd7 As New InfiniCommand("TblOutStanding")
                cmd7.AddParameter("@MyAuto", CInt(dr("MyAuto")))
                cmd7.AddParameter("@ParentId", CInt(dr("ParentId")))
                cmd7.AddParameter("@ChildId", CInt(dr("ChildId")))
                cmd7.AddParameter("@TransId", CStr(dr("TransId")))
                cmd7.AddParameter("@Debit", CDbl(dr("Debit")))
                cmd7.AddParameter("@Credit", CDbl(dr("Credit")))
                cmd7.AddParameter("@TranType", CStr(dr("TranType")))
                cmd7.AddParameter("@Details", CStr(dr("Details")))
                cmd7.AddParameter("@TranRef", CStr(dr("TranRef")))
                cmd7.AddParameter("@OSDate", CStr(dr("OSDate")))
                cmd7.AddParameter("@OSFrom", CInt(dr("OSFrom")))
                cmd7.AddParameter("@OSTo", CInt(dr("OSTo")))
                cmd7.AddParameter("@OSFT", CStr(dr("OSFT")))
                cmd7.AddParameter("@OSNo", CInt(dr("OSNo")))
                cmd7.AddParameter("@Update", ChFlag)
                cmdCollection.Add(cmd7)

            Next

        End If

    End Sub

    Public Sub UpdateSynFlagDeskTopData(ByVal _tempDS As DataSet)

        Dim Id As Integer

        'Step 1
        Dim cmd As New InfiniCommand("TblBankSynFlag")
        Id = ClassGlobal.ExecuteNonQuery(cmd)

        Dim cmd1 As New InfiniCommand("TblCompanySynFlag")
        Id = ClassGlobal.ExecuteNonQuery(cmd1)

        Dim cmd2 As New InfiniCommand("TblFinancialYearSynFlag")
        Id = ClassGlobal.ExecuteNonQuery(cmd2)

        Dim cmd3 As New InfiniCommand("TblTaxIdsSynFlag")
        Id = ClassGlobal.ExecuteNonQuery(cmd3)

        GlbGlobal.ExecuteCommand()

    End Sub
    Public Sub UpdateSynFlagDeskTopDataForCV(ByVal _tempDS As DataSet)

        Dim Id As Integer

        'Step 2
        UpdateCustomerSynFlag(_tempDS.Tables("TblCustomer"))
        UpdateVendorSynFlag(_tempDS.Tables("TblVendor"))

        GlbGlobal.ExecuteCommand()

    End Sub

    Public Sub UpdateSynFlagDeskTopDataForT(ByVal _tempDS As DataSet)

        Dim Id As Integer

        'Step 3
        UpdateTransactionSynFlag(_tempDS.Tables("TblTransaction10"))

        UpdateLedgerSynFlag(_tempDS.Tables("TblLedger10"))

        UpdateOutStandingSynFlag(_tempDS.Tables("TblOutStanding"))

        'UpdateCCInfoSynFlag(_tempDS.Tables("TblOutStanding"))
        'Dim cmd9 As New InfiniCommand("TblCreditCardInfoFlag")
        'Id = ClassGlobal.ExecuteNonQuery(cmd9)

        GlbGlobal.ExecuteCommand()

    End Sub


    Public Sub UpdateCustomerSynFlag(ByVal _tempDS As DataTable)
        Dim id As Integer
        If Not _tempDS Is Nothing Then
            Dim dr As DataRow

            For Each dr In _tempDS.Rows
                Dim cmd4 As New InfiniCommand("TblCustomerSynFlag")
                With cmd4
                    .AddParameter("CustomerID", dr.Item("CustomerId"))
                End With
                id = ClassGlobal.ExecuteNonQuery(cmd4)
            Next
        End If
    End Sub

    Public Sub UpdateVendorSynFlag(ByVal _tempDS As DataTable)
        Dim id As Integer
        If Not _tempDS Is Nothing Then
            Dim dr As DataRow

            For Each dr In _tempDS.Rows
                Dim cmd4 As New InfiniCommand("TblVendorSynFlag")
                With cmd4
                    .AddParameter("VenoderID", dr.Item("VendorID"))
                End With
                id = ClassGlobal.ExecuteNonQuery(cmd4)
            Next
        End If
    End Sub
    Public Sub UpdateTransactionSynFlag(ByVal _tempDS As DataTable)
        Dim id As Integer
        If Not _tempDS Is Nothing Then
            Dim dr As DataRow

            For Each dr In _tempDS.Rows
                Dim cmd4 As New InfiniCommand("TblTransactionSynFlag")
                With cmd4
                    .AddParameter("ParentID", dr.Item("ParentID"))
                End With
                id = ClassGlobal.ExecuteNonQuery(cmd4)
            Next
        End If
    End Sub

    Public Sub UpdateLedgerSynFlag(ByVal _tempDS As DataTable)
        Dim id As Integer
        If Not _tempDS Is Nothing Then
            Dim dr As DataRow

            For Each dr In _tempDS.Rows
                Dim cmd4 As New InfiniCommand("TblLedgerSynFlag")
                With cmd4
                    .AddParameter("ParentID", dr.Item("ParentID"))
                End With
                id = ClassGlobal.ExecuteNonQuery(cmd4)
            Next
        End If
    End Sub

    Public Sub UpdateOutStandingSynFlag(ByVal _tempDS As DataTable)
        Dim id As Integer
        If Not _tempDS Is Nothing Then
            Dim dr As DataRow

            For Each dr In _tempDS.Rows
                Dim cmd4 As New InfiniCommand("TblOutStandingSynFlag")
                With cmd4
                    .AddParameter("ParentID", dr.Item("ParentID"))
                End With
                id = ClassGlobal.ExecuteNonQuery(cmd4)
            Next
        End If
    End Sub

    Public Function GetUnsynchronizeDataFromDeskTop() As DataSet

        Dim ts As DataSet
        Dim dr As IDataReader

        Dim cmd0 As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd0)

        'Step 1
        Dim cmd1 As New InfiniCommand("TblBank")
        'cmd1.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd1)

        Dim cmd2 As New InfiniCommand("TblCompany")
        'cmd2.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd2)

        Dim cmd4 As New InfiniCommand("TblFinancialYear")
        'cmd4.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd4)

        Dim cmd7 As New InfiniCommand("TblTaxIds")
        'cmd7.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd7)

        'Step2
        Dim cmd3 As New InfiniCommand("TblCustomer")
        'cmd3.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd3)

        Dim cmd9 As New InfiniCommand("TblVendor")
        'cmd9.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd9)

        'Step 3
        Dim cmd8 As New InfiniCommand("TblTransaction")
        'cmd8.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd8)

        Dim cmd5 As New InfiniCommand("TblLedger")
        'cmd5.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd5)

        Dim cmd6 As New InfiniCommand("TblOutStanding")
        'cmd6.AddParameter("@TranDateTime", TranDateTime)
        cmdCollection.Add(cmd6)

        'Dim cmd10 As New InfiniCommand("TblCreditCardInfo")
        ''cmd10.AddParameter("@TranDateTime", TranDateTime)
        'cmdCollection.Add(cmd10)

        ts = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ts

    End Function
    Public Function GetUnsynchronizeCVData() As DataSet
        Dim ds1 As DataSet

        Dim cmd0 As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd0)

        Dim cmd1 As New InfiniCommand("TBLSELECTALLCUSTOMER")
        cmdCollection.Add(cmd1)

        Dim cmd2 As New InfiniCommand("TBLSELECTALLVENDOR")
        cmdCollection.Add(cmd2)

        ds1 = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ds1

    End Function

    Public Function GetUnsynchronizeTLOData() As DataSet
        Dim ds1 As DataSet

        Dim cmd0 As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd0)

        Dim cmd8 As New InfiniCommand("TBLSELECTALLTRANSACTION")
        cmdCollection.Add(cmd8)

        Dim cmd5 As New InfiniCommand("TBLSELECTALLLEDGER")
        cmdCollection.Add(cmd5)

        Dim cmd6 As New InfiniCommand("TBLSELECTALLOUTSTANDING")
        cmdCollection.Add(cmd6)

        'Dim cmd10 As New InfiniCommand("TblCreditCardInfo")
        ''cmd10.AddParameter("@TranDateTime", TranDateTime)
        'cmdCollection.Add(cmd10)

        ds1 = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)
        cmdCollection = New Collection
        Return ds1

    End Function

End Class



