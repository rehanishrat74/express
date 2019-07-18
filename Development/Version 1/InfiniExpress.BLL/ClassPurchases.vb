Imports System.Data
Imports InfiniExpress.DAL

Public Class ClassPurchases

    Dim ClsGlobal As New ClassGlobal()
    Dim strsql As String
    Private PurchaseID As Integer
    Dim _dr As IDataReader
    Dim _Dts As DataSet
    Public Function InsertOnlineVendor(ByVal VendorID As String, ByVal VendorName As String, ByVal Street As String, ByVal Town As String, _
                                     ByVal Country As String, ByVal PostalCode As String, ByVal ContactName As String, _
                                     ByVal Email As String, ByVal Telephone As String, ByVal Fax As String, _
                                     ByVal VatNo As String, ByVal CreditLimit As Double)

        Dim _cmd As New InfiniCommand("INSERTONLINEVEDNOR")

        With _cmd
            .AddParameter("VendorID", VendorID)
            .AddParameter("VendorName", VendorName)
            .AddParameter("Street", Street)
            .AddParameter("Town", Town)
            .AddParameter("Country", Country)
            .AddParameter("PostCode", PostalCode)
            .AddParameter("ContactName", ContactName)
            .AddParameter("Email", Email)
            .AddParameter("Telephone", Telephone)
            .AddParameter("Fax", Fax)
            .AddParameter("Vat", VatNo)
            .AddParameter("CreditLimit", CreditLimit)
            .AddParameter("TranDateTime", Date.Now)
        End With

        ClassGlobal.ExecuteNonQuery(_cmd)


    End Function

    Public Sub InsertNewVendor(ByVal VendorName As String, ByVal Email As String, ByVal DateTime As String)

        'Insert New Vendor
        Dim _cmd As New InfiniCommand("INSERTNEWVENDOR")
        With _cmd
            .AddParameter("VendorName", VendorName)
            .AddParameter("Email", Email)
            .AddParameter("DateTime", DateTime)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub

    Public Function CreateVendorID(ByVal VendorName As String, ByVal Email As String, ByVal DateTime As String) As Integer

        Dim Maxid As Integer = 1
        Dim _cmd As New InfiniCommand("CREATEVENDORID")
        With _cmd
            .AddParameter("VendorName", VendorName)
            .AddParameter("Email", Email)
            .AddParameter("DateTime", DateTime)
        End With

        _dr = ClassGlobal.GetDataReader(_cmd)
        If _dr.Read Then
            Maxid = CInt(_dr("UniqueNo"))
        End If
        _dr.Close()
        Return Maxid

    End Function

    Public Sub UpdateVendor(ByVal VendorId As String, ByVal VendorName As String, ByVal Street As String, _
                ByVal Town As String, ByVal Country As String, ByVal PostalCode As String, ByVal ContactName As String, _
                ByVal Email As String, ByVal Telephone As String, ByVal Fax As String, ByVal Vat As String, ByVal CreditLimit As Double, ByVal DateTime As String, ByVal UniqueNo As Integer)

        Dim _cmd1 As New InfiniCommand("UPDATEVENDOR")
        With _cmd1
            .AddParameter("VendorID", VendorId)
            .AddParameter("VendorName", VendorName)
            .AddParameter("Street", Street)
            .AddParameter("Town", Town)
            .AddParameter("Country", Country)
            .AddParameter("PostalCode", PostalCode)
            .AddParameter("Email", Email)
            .AddParameter("ContactName", ContactName)
            .AddParameter("Telephone", Telephone)
            .AddParameter("Fax", Fax)
            .AddParameter("Vat", Vat)
            .AddParameter("CreditLimit", CreditLimit)
            .AddParameter("DateTime", DateTime)
            .AddParameter("UniqueNo", UniqueNo)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd1)

    End Sub

    Public Function GetDCAgainstTransaction(ByVal Transid As Integer, ByVal InvTp As String) As Double

        Dim _cmd As New InfiniCommand("GetDCAgainstTransaction")
        Dim GetDC As Double, INAT As Double, IVAT As Double
        With _cmd
            .AddParameter("osto", Transid)
            .AddParameter("Invtp", InvTp)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        GetDC = 0
        While _dr.Read()
            INAT = CDbl(_dr(6))
            IVAT = CDbl(_dr(7))
            GetDC = GetDC + INAT + IVAT
        End While
        Return GetDC

    End Function

    Public Function VendorInvoiceRefund(ByVal TPid As Integer) As IDataReader

        Dim _cmd As New InfiniCommand("VendorInvoiceRefund")
        With _cmd
            .AddParameter("VendorID", TPid)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function GetVendorDetails(ByVal _tempID As String) As IDataReader

        Dim _cmd As New InfiniCommand("GetVendorDetail")
        With _cmd
            .AddParameter("VendorID", _tempID)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function LoadVendorID() As IDataReader

        Dim _cmd As New InfiniCommand("LoadVendorIDs")
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function GetVendorActSummary(ByVal VendId As String) As DataSet

        Dim _cmd As New InfiniCommand("GETVENDORACTSUMMARY")
        With _cmd
            .AddParameter("VendorID", VendId)
        End With
        _Dts = ClassGlobal.GetDataSet(_cmd)
        Return _Dts

    End Function

    Public Function LoadVendorPayments(ByVal tempCID As String) As IDataReader
        Dim _dr As IDataReader
        Dim _cmd As New InfiniCommand("LOADVENDORPAYMENTS")
        With (_cmd)
            .AddParameter("VendorID", tempCID)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function PurchaseRefund(ByVal _id As String) As IDataReader

        Dim _cmd As New InfiniCommand("PURCHASEREFUND")
        With _cmd
            .AddParameter("VendorID", _id)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function GetPurchaseInvoiveDetail(ByVal VendID As String, ByVal ParentIDs As String) As DataSet

        Dim TransIDs(5) As String
        Dim transid As String
        Dim i As Integer = 0
        Dim _cmd As New InfiniCommand("GetPurchaseInvoiceDetail")
        _cmd.AddParameter("VendID", VendID)
        TransIDs = ParentIDs.Split(","c)
        For Each transid In TransIDs
            i += 1
            _cmd.AddParameter("ParentID" & Convert.ToString(i), transid)
        Next
        _Dts = ClassGlobal.GetDataSet(_cmd)
        Return _Dts

    End Function

    Public Function VendorDeleted(ByVal VendorId As String) As Integer

        Dim Flag As Integer = 0
        Dim _cmd As New InfiniCommand("VENDORDELETE")
        With _cmd
            .AddParameter("VendorID", VendorId)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)

        Dim _cmd1 As New InfiniCommand("CONFDELETE")
        If _dr.Read = False Then
            With _cmd1
                .AddParameter("VendorID", VendorId)
            End With
            ClassGlobal.ExecuteNonQuery(_cmd1)
            ClassGlobal.ExecuteCommand()
            Flag = 1
        End If
        Return Flag

    End Function

    Public Function VendorSummary() As DataSet

        Dim _cmd As New InfiniCommand("LOADVENDORTSUMMARY")
        _Dts = ClassGlobal.GetDataSet(_cmd)
        Return _Dts

    End Function

    Public Function GetSynchronizeTableData() As DataSet

        Dim cmdCollection As New Collection
        Dim wds As DataSet
        Dim cmd As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd)

        wds = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)

        cmdCollection = New Collection

        Return wds

    End Function

End Class
