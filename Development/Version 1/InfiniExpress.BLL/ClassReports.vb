Option Strict On
Imports System
Imports System.Data
Imports InfiniExpress.DAL

Public Class ClassReports

    Public Function GetCustID() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_CUSTOMERID")
        ' Dim cmd As New InfiniCommand("REPORT")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetVendID() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_VENDORID")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetLedgerID() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_LEDGERID")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetCustList() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETCUSTOMERSLIST")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetCustActivity(ByVal CustID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETCUSTOMERACTIVITY")
        cmd.AddParameter("@CustID", CustID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function
    Public Function GetCompanyInfo() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETCOMPANYINFO")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetSalesInvoice(ByVal CustID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETSALESINVOICE")
        cmd.AddParameter("@CustID", CustID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetSalesReceipts(ByVal CustID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETSALESRECEIPTS")
        cmd.AddParameter("@CustID", CustID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetVendorList() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETVENDORSLIST")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetVendorActivity(ByVal VendID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETVENDORACTIVITY")
        cmd.AddParameter("@VendID", VendID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetPurchaseInvoice(ByVal VendID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETPURCHASEINVOICE")
        cmd.AddParameter("@VendID", VendID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetPurchasePayments(ByVal VendID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETPURCHASEPAYMENTS")
        cmd.AddParameter("@VendID", VendID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetLedgerBalance() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETLEDGERBALANCE")
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetLedgerActivity(ByVal LedgerID As String) As DataSet
        Dim cmd As New InfiniCommand("REPORTS_GETLEDGERACTIVITY")
        cmd.AddParameter("@LedgerID", LedgerID)
        Return ClassGlobal.GetDataSet(cmd)
    End Function

    Public Function GetInvoiceDetail() As DataSet
        Dim cmd As New InfiniCommand("REPORTS_INVOICEDETAIL")
        Return ClassGlobal.GetDataSet(cmd)
    End Function
    
    Public Function GetParentIDno(ByVal tempPID As String) As String
        Dim Str As String
        Dim Dr As IDataReader
        Dim Cmd As New InfiniCommand("GETTOTPID")
        With Cmd
            .AddParameter("GroupNo", CInt(tempPID))
        End With

        Dr = ClassGlobal.GetDataReader(Cmd)

        Do While Dr.Read = True

            If Str = "" Then
                Str = Convert.ToString(Dr.Item("parentid"))
            Else
                Str = Str & "," & Convert.ToString(Dr.Item("parentid"))
            End If

        Loop
        Return Str

    End Function

    Public Function ShowCreditCardInfo(ByVal tempCustID As String) As DataSet
        Dim ds As DataSet
        Dim cmd As New InfiniCommand("GETCCINFO")
        With cmd
            .AddParameter("CustomerID", tempCustID)
        End With
        ds = ClassGlobal.GetDataSet(cmd)
        Return ds
    End Function

End Class
