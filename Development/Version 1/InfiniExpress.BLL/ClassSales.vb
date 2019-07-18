Option Strict On
Imports System
Imports System.Data
Imports InfiniExpress.DAL

Public Class ClassSales

    Dim _strsql As String
    Dim _dr As IDataReader
    Dim Ds As DataSet

    Sub New()
    End Sub

    Public Function CustomerDeleted(ByVal cmbId As Integer) As Integer

        Dim _dr As IDataReader
        Dim Flag As Integer = 0
        Dim _cmd As New InfiniCommand("CUSTOMERDELETE")

        _cmd.AddParameter("Customer_ID", cmbId)
        _dr = ClassGlobal.GetDataReader(_cmd)

        Dim _cmd1 As New InfiniCommand("CONFDELETECUST")
        If _dr.Read = False Then
            With _cmd1
                .AddParameter("Customer_ID", cmbId)
            End With
            ClassGlobal.ExecuteNonQuery(_cmd1)
            ClassGlobal.ExecuteCommand()
            Flag = 1
        End If
        Return Flag

    End Function

    Public Sub InsertNewCustomer(ByVal CustomerName As String, ByVal Email As String, ByVal DateTime As String)

        'Insert New Customer
        Dim _cmd As New InfiniCommand("INSERTNEWCUSTOMER")
        With _cmd
            .AddParameter("CustomerName", CustomerName)
            .AddParameter("Email", Email)
            .AddParameter("DateTime", DateTime)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub

    Public Function CreateCustomerID(ByVal CustomerName As String, ByVal Email As String, ByVal DateTime As String) As Integer

        Dim Maxid As Integer = 1
        Dim _cmd As New InfiniCommand("CREATECUSTOMERID")
        With _cmd
            .AddParameter("CustomerName", CustomerName)
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

    Public Sub UpdateCustomer(ByVal Customer_Id As Integer, ByVal CustomerName As String, ByVal Street As String, _
                ByVal Town As String, ByVal Country As String, ByVal PostalCode As String, ByVal ContactName As String, _
                ByVal Email As String, ByVal Telephone As String, ByVal Fax As String, ByVal Vat As String, _
                 ByVal CreditLimit As Double, ByVal UniqueNo As Integer)

        'Update Customer Details
        Dim _cmd As New InfiniCommand("UPDATECUSTOMER")
        With _cmd
            .AddParameter("CustomerID", Customer_Id)
            .AddParameter("CustomerName", CustomerName)
            .AddParameter("Street", Street)
            .AddParameter("Town", Town)
            .AddParameter("Country", Country)
            .AddParameter("PostCode", PostalCode)
            .AddParameter("ContactName", ContactName)
            .AddParameter("Email", Email)
            .AddParameter("Telephone", Telephone)
            .AddParameter("Fax", Fax)
            .AddParameter("Vat", Vat)
            .AddParameter("CreditLimit", CreditLimit)
            .AddParameter("UniqueNo", UniqueNo)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub

    Public Sub AmendCustomer(ByVal Customer_Id As String, ByVal Name As String, ByVal Street As String, _
                    ByVal Town As String, ByVal Country As String, ByVal PostalCode As String, ByVal ContactName As String, _
                    ByVal Email As String, ByVal Telephone As String, ByVal Fax As String, ByVal Vat As String, _
                    ByVal CreditLimit As Double) ', ByVal DateTime As String, ByVal UniqueNo As Integer)

        'Update Customer Details
        Dim _cmd As New InfiniCommand("AMENDCUSTOMER")
        With _cmd
            .AddParameter("Customer_ID", Customer_Id)
            .AddParameter("Name", Name)
            .AddParameter("Street", Street)
            .AddParameter("Town", Town)
            .AddParameter("Country", Country)
            .AddParameter("PostCode", PostalCode)
            .AddParameter("Cont_Name", ContactName)
            .AddParameter("Cart_Customer_Email", Email)
            .AddParameter("Telephone", Telephone)
            .AddParameter("Fax", Fax)
            .AddParameter("Vat", Vat)
            .AddParameter("CreditLimit", CreditLimit)
            '.AddParameter("DateTime", DateTime)
            ' .AddParameter("UniqueNo", UniqueNo)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub


    Public Function LoadCustomerID() As IDataReader
        Dim _cmd As New InfiniCommand("LoadCustomerIDS")
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr
    End Function

    Public Function GetCustInfo(ByVal _tempID As String) As IDataReader

        Dim _cmd As New InfiniCommand("GetCustomerDetail")
        _cmd.AddParameter("CustomerID", _tempID)
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function LoadCustSummary() As DataSet
        Dim _cmd As New InfiniCommand("LoadCustomerSummary")
        Ds = ClassGlobal.GetDataSet(_cmd)
        Return Ds
    End Function

    Public Function GetSalesInvoiveDetail(ByVal CustID As String, ByVal ParentIDs As String) As DataSet

        Dim TransIDs(5) As String
        Dim transid As String
        Dim i As Integer = 0
        Dim _cmd As New InfiniCommand("GetSalesInvoiceDetail")
        _cmd.AddParameter("CustID", CustID)
        TransIDs = ParentIDs.Split(","c)
        For Each transid In TransIDs
            i += 1
            _cmd.AddParameter("ParentID" & Convert.ToString(i), transid)
        Next
        Ds = ClassGlobal.GetDataSet(_cmd)
        Return Ds

    End Function

    Public Function LoadCustomerReceipts(ByVal tempCID As String) As IDataReader
        Dim Dr As IDataReader
        Dim _cmd As New InfiniCommand("LOADCUSTOMERRECEIPTS")
        _cmd.AddParameter("CustomerID", tempCID)
        Dr = ClassGlobal.GetDataReader(_cmd)
        'Ds = ClassGlobal.GetDataSet(_cmd)
        Return Dr

    End Function

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
        _dr.Close()
        Return GetDC

    End Function

    Public Function GetCustomerActSummary(ByVal CId As String) As DataSet

        Dim _cmd As New InfiniCommand("GETCUSTOMERACTSUMMARY")
        With _cmd
            .AddParameter("Customer_ID", CId)
        End With
        Ds = ClassGlobal.GetDataSet(_cmd)
        Return Ds

    End Function

    Public Function SalesRefund(ByVal _id As String) As IDataReader

        Dim _cmd As New InfiniCommand("SALESREFUND")
        With _cmd
            .AddParameter("CustomerID", _id)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Function CustomerInvoiceRefund(ByVal TPid As Integer) As IDataReader

        Dim _cmd As New InfiniCommand("CustomerInvoiceRefund")
        With _cmd
            .AddParameter("CustomerID", TPid)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        Return _dr

    End Function

    Public Sub InsertCustomer(ByVal _email As String, ByVal _tTime As Date, ByVal _custname As String)
        Dim _cmd As New InfiniCommand("INSERTCUSTOMER")
        With _cmd
            .AddParameter("Email", _email)
            .AddParameter("TranDateTime", _tTime)
            .AddParameter("CustomerName", _custname)
        End With
        Dim _id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)

        If ClassGlobal.ExecuteCommand = False Then Exit Sub
    End Sub

    Public Function GetCustomer_ID(ByVal _dtime As Date, ByVal _email As String) As Integer
        Dim _dr As IDataReader
        Dim _cmd As New InfiniCommand("GetCustomerAutoNo")
        With _cmd
            .AddParameter("trandatetime", _dtime)
            .AddParameter("Email", _email)
        End With
        _dr = ClassGlobal.GetDataReader(_cmd)
        _dr.Read()
        Return CInt(_dr.Item("uniqueno"))

    End Function

    Public Sub MCKey(ByVal _logID As Integer, ByVal _logkey As String)
        Dim _cmd As New InfiniCommand("INSERTMCKEY")
        With _cmd

            .AddParameter("LogID", _logID)
            .AddParameter("LogKey", _logkey)
            .AddParameter("LogType", "C")

        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)
        If ClassGlobal.ExecuteCommand = False Then Exit Sub

    End Sub
    Public Function getPdfinvNo() As Integer
        Dim _dr As IDataReader
        Dim _cmd As New InfiniCommand("GETPDFINVNO")
        Dim tempVar As Integer
        _dr = ClassGlobal.GetDataReader(_cmd)

        _dr.Read()
        If _dr.IsDBNull(0) = True Then
            getPdfinvNo = 1
        Else
            tempVar = CInt(_dr.Item("Ino"))
            getPdfinvNo = tempVar + 1

        End If
        _dr = Nothing
        Return getPdfinvNo

    End Function
    Public Sub Insertpdfinvoice(ByVal _datafile As String, ByVal _invdate As Date, ByVal _invtp As String, ByVal _invno As Integer, ByVal _custID As String, ByVal _invNet As Double, ByVal _tranDT As String)
        Dim _cmd As New InfiniCommand("INSERTPDF")
        With _cmd
            .AddParameter("Datafile", _datafile)
            .AddParameter("Invdate", _invdate)
            .AddParameter("Invtp", _invtp)
            .AddParameter("Invno", _invno)
            .AddParameter("CustID", _custID)
            .AddParameter("Invnet", _invNet)
            .AddParameter("TranDateTime", _tranDT)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)
        If ClassGlobal.ExecuteCommand = False Then Exit Sub
    End Sub
    Public Sub updateGroupno(ByVal _invno As Integer, ByVal _parentno As Integer)

        Dim _cmd As New InfiniCommand("UPDATEGROUPNO")
        With _cmd
            .AddParameter("invno", _invno)
            .AddParameter("ParentID", _parentno)
        End With
        Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub
    Public Sub InsertOnlineCustomer(ByVal Customer_Id As Integer, ByVal CustomerName As String, ByVal Street As String, _
                ByVal Town As String, ByVal Country As String, ByVal PostalCode As String, ByVal ContactName As String, _
                ByVal Email As String, ByVal Telephone As String, ByVal Fax As String, ByVal Vat As String, _
                 ByVal CreditLimit As Double)

        Dim _cmd As New InfiniCommand("INSERTONLINECUSTOMER")
        With _cmd
            .AddParameter("CustomerID", Customer_Id)
            .AddParameter("CustomerName", CustomerName)
            .AddParameter("Street", Street)
            .AddParameter("Town", Town)
            .AddParameter("Country", Country)
            .AddParameter("PostCode", PostalCode)
            .AddParameter("ContactName", ContactName)
            .AddParameter("Email", Email)
            .AddParameter("Telephone", Telephone)
            .AddParameter("Fax", Fax)
            .AddParameter("Vat", Vat)
            .AddParameter("CreditLimit", CreditLimit)
            .AddParameter("TranDateTime", Now.Now)
        End With
        ClassGlobal.ExecuteNonQuery(_cmd)

    End Sub

    Public Function GetSynchronizeTableData() As DataSet

        Dim cmdCollection As New Collection
        Dim wds As DataSet
        Dim cmd As New InfiniCommand("TblSynchronize")
        cmdCollection.Add(cmd)

        wds = AccessHelper.ExecuteSynchronizeCollection(cmdCollection)

        cmdCollection = New Collection

        Return wds

    End Function

    Public Sub InsertCardInfo(ByVal _tempInvno As Integer, ByVal _tempdate As String, ByVal _tempAmt As Double, ByVal _tempShipComp As String, _
                                        ByVal _tempTrackNo As String, ByVal _tempDeliveryAddress As String, ByVal _tempCardType As String, _
                                        ByVal _tempCardHN As String, ByVal _tempCardNo As String, ByVal _tempExpDate As String, ByVal _tempSecurityCode As String, _
                                        ByVal _tempCardAddress As String, ByVal _tranDT As DateTime, ByVal _tempAc As String, ByVal _tempInvdetail As String, ByVal _tempInvtype As String, ByVal _tempPayment As String, ByVal _tempCustID As String)

        Dim Cmd As New InfiniCommand("CreditCardInfo")
        With Cmd
            .AddParameter("Inv_no", _tempInvno)
            .AddParameter("Inv_date", _tempdate)
            .AddParameter("Debit", _tempAmt)
            .AddParameter("ShippingComy", _tempShipComp)
            .AddParameter("TrackingNumber", _tempTrackNo)
            .AddParameter("DeliveryAddress", _tempDeliveryAddress)
            .AddParameter("CardType", _tempCardType)
            .AddParameter("CardHoldername", _tempCardHN)
            .AddParameter("CardNumber", _tempCardNo)
            .AddParameter("ExpiryDate", _tempExpDate)
            .AddParameter("SecurityCode", _tempSecurityCode)
            .AddParameter("CardAddress", _tempCardAddress)
            .AddParameter("TranDateTime", _tranDT)
            .AddParameter("Ac", _tempAc)
            .AddParameter("Invdetails", _tempInvdetail)
            .AddParameter("Invtype", _tempInvtype)
            .AddParameter("Paymentmode", _tempPayment)
            .AddParameter("CustomerID", _tempCustID)
        End With
        Dim id As Integer = ClassGlobal.ExecuteNonQuery(Cmd)

    End Sub

    Public Sub insertCreditCardInfo(ByVal _tempInvno As Integer, ByVal _tempdate As String, ByVal _tempAmt As Double, ByVal _tempDT As DateTime, ByVal _tempAc As String, ByVal _tempInvdetail As String, ByVal _tempInvtype As String, ByVal _tempPayment As String, ByVal _tempCustID As String)

        Dim Cmd As New InfiniCommand("CNCreditCardInfo")
        With Cmd
            .AddParameter("Inv_no", _tempInvno)
            .AddParameter("Inv_date", _tempdate)
            .AddParameter("Debit", _tempAmt)
            .AddParameter("TranDateTime", _tempDT)
            .AddParameter("Ac", _tempAc)
            .AddParameter("Invdetails", _tempInvdetail)
            .AddParameter("Invtype", _tempInvtype)
            .AddParameter("Paymentmode", _tempPayment)
            .AddParameter("CustomerID", _tempCustID)
        End With
        Dim id As Integer = ClassGlobal.ExecuteNonQuery(Cmd)

    End Sub


    Public Sub insertCreditCardInfoFORSC(ByVal _tempInvno As Integer, ByVal _tempdate As String, ByVal _tempAmt As Double, ByVal _tempDT As DateTime, ByVal _tempAc As String, ByVal _tempInvdetail As String, ByVal _tempInvtype As String, ByVal _tempPayment As String, ByVal _tempCustID As String)

        Dim Cmd As New InfiniCommand("CNCreditCardInfoFORSC")
        With Cmd
            .AddParameter("Inv_no", _tempInvno)
            .AddParameter("Inv_date", _tempdate)
            .AddParameter("Credit", _tempAmt)
            .AddParameter("TranDateTime", _tempDT)
            .AddParameter("Ac", _tempAc)
            .AddParameter("Invdetails", _tempInvdetail)
            .AddParameter("Invtype", _tempInvtype)
            .AddParameter("Paymentmode", _tempPayment)
            .AddParameter("CustomerID", _tempCustID)
        End With
        Dim id As Integer = ClassGlobal.ExecuteNonQuery(Cmd)

    End Sub

    Public Function GetSynCustID() As Integer

        Dim dr As IDataReader
        Dim cmd As New InfiniCommand("TBLSYNCHRONIZE")
        dr = ClassGlobal.GetDataReader(cmd)

        dr.Read()
        Return CInt(dr("UserID"))

    End Function

    Public Function GetCustomerservice(ByVal _tempCustomerID As String) As DataSet

        Dim cmd As New InfiniCommand("GETSERSTATUS")
        With cmd
            .AddParameter("Customerid", _tempCustomerID)
        End With
        Return ClassGlobal.GetDataSet(cmd)

    End Function

    Public Function GetRefundedInvNo(ByVal _tempID As Integer) As Integer

        Dim Dr As IDataReader, Invno As Integer
        Dim cmd As New InfiniCommand("GETREFUNDEDINVNO")
        With cmd
            .AddParameter("Parentid", _tempID)
        End With
        Dr = ClassGlobal.GetDataReader(cmd)
        Dr.Read()
        Invno = CInt(Dr("GroupNo"))
        Dr.Close()
        Return Invno

    End Function

    Public Sub SendRefundInvoiceForServiceCollection(ByVal invDetail As String, ByVal InvType As String, ByVal InvNo As Integer, ByVal Ac As String, _
                                                                        ByVal PayMode As String, ByVal InvDate As Date, ByVal InvAmount As Double, ByVal DT As DateTime)
        Dim cmd09 As New InfiniCommand("INSERTSYNREFUNDINV")
        With cmd09
            .AddParameter("Inv_no", InvNo)
            .AddParameter("Inv_date", InvDate)
            .AddParameter("Debit", InvAmount)
            .AddParameter("TranDateTime", DT)
            .AddParameter("Ac", "@ACC" & Ac)
            .AddParameter("Invdetails", invDetail)
            .AddParameter("Invtype", InvType)
            .AddParameter("Paymentmode", PayMode)
            .AddParameter("CustomerID", Ac)
        End With
        Dim id As Integer = ClassGlobal.ExecuteNonQuery(cmd09)

    End Sub

    Public Function checkCCNo(ByVal _tempInv As Integer) As String
        Dim _dr As IDataReader
        Dim cmd091 As New InfiniCommand("GETCCNO")
        Dim _str As String
        With cmd091
            .AddParameter("Inv_no", _tempInv)
        End With
        _dr = ClassGlobal.GetDataReader(cmd091)

        Do While _dr.Read = True
            Return "CC"
        Loop
        Return "CB"

    End Function

End Class