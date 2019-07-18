Option Strict On
#Region "Import Libraries"
Imports System
Imports System.Data
#End Region

Public Class AccessQueries

    Private Shared _SqlStr As String
    Private Shared QTime As String
    Private Shared QLast As Date

    Public Shared Function GetQueryString(ByRef _cmd As InfiniCommand) As String

        Dim SPName As String, SPQuery As String

        QTime = CStr(QLast.Now)

        SPName = _cmd.CommandName

        Select Case UCase(SPName)

            'Case "GETCCINFO"
            '    SPQuery = "Select * from tblcreditcardinfo"
        Case "GETCCNO"
                SPQuery = GETCCNO(_cmd)
            Case "GETCCINFO"
                SPQuery = GETCCINFO(_cmd)
            Case "INSERTONLINEVEDNOR"
                SPQuery = InsertOnlineVendor(_cmd)
            Case "INSERTONLINECUSTOMER"
                SPQuery = InsertOnlineCustomer(_cmd)
            Case "VENDORDELETE"
                SPQuery = VendorDelete(_cmd)
            Case "CONFDELETE"
                SPQuery = CONFDELETE(_cmd)
            Case "GETPURCHASEMAXID"
                SPQuery = "SELECT Max(UniqueNo)+1 As VendNo FROM Tblvendor"
            Case "GETLASTSYNCHRONIZEDATE"
                SPQuery = "SELECT * FROM TblSynchronize Where SynId=1"
            Case "UPDATESYNCHRONIZEID"
                SPQuery = UpdateSynchronizeId(_cmd)
            Case "UPDATESYNCHRONIZEDATETIME"
                SPQuery = UpdateSynchronizeDateTime()
            Case "INSERTNEWCUSTOMER"
                SPQuery = InsertNewCustomer(_cmd)
            Case "CREATECUSTOMERID"
                SPQuery = CreateCustomerId(_cmd)
            Case "UPDATECUSTOMER"
                SPQuery = UpdateCustomer(_cmd)
            Case "LOADCUSTOMERIDS"
                SPQuery = "Select CustomerId From tblCustomer Where DFlag='N' Order By CustomerId"
            Case "LOADCUSTOMERSUMMARY"
                SPQuery = LoadCustomerSummary()
            Case "GETCUSTOMERDETAIL"
                SPQuery = GetCustomerDetail(_cmd)
            Case "CONFDELETECUST"
                SPQuery = CONFDELETECUST(_cmd)
            Case "CUSTOMERDELETE"
                SPQuery = CustomerDelete(_cmd)
            Case "INSERTNEWVENDOR"
                SPQuery = InsertNewVendor(_cmd)
            Case "CREATEVENDORID"
                SPQuery = CreateVendorId(_cmd)
            Case "UPDATEVENDOR"
                SPQuery = UpdateVendor(_cmd)
            Case "LOADVENDORIDS"
                SPQuery = "Select Vendorid from tblvendor where dflag='N' order by VendorId"
            Case "GETVENDORDETAIL"
                SPQuery = GetVendorDetail(_cmd)
            Case "LOADVENDORTSUMMARY"
                SPQuery = LoadVendorSummary()
            Case "GETFINANCIALYEAR"
                SPQuery = "Select * From TblFinancialYear Order By MyNo"
            Case "CHECKFINANCIALYEARTRANSACTION"
                SPQuery = "Select  ParentId , InvTp , TaxId , InvDate ,  InvDetails  , InvNet , InvVAT , Refund  , Mark , V , TranDate From TblTransaction"
            Case "LOADTAXIDS"
                SPQuery = "Select Top 9 * From TblTaxIds Order By CountId"
            Case "GETTRANSACTIONNO"
                SPQuery = "SELECT Max(ParentId)+1 As ParentNo FROM TblTransaction"
            Case "GETLEDGERAUTONO"
                SPQuery = "SELECT Max(MyAutoNo)+1 As AutoNo FROM TblLedger"

            Case "GETMAXOSNO"
                SPQuery = "SELECT IIF(ISNULL(Max(OSNo)+1),1,Max(OSNo)+1) As OSMaxNo FROM TblOutStanding"
            Case "GETEXISTLEDGERAUTONO"
                SPQuery = GetLedgerAutoNo(_cmd)
            Case "LOADCUSTOMERRECEIPTS"
                SPQuery = LoadCustomerReceipts(_cmd)
            Case "LOADVENDORPAYMENTS"
                SPQuery = LOADVENDORPAYMENTS(_cmd)
            Case "BANKINFO"
                SPQuery = LoadBankDetail(_cmd)
            Case "UPDATEBANK"
                SPQuery = UpdateBankDetail(_cmd)
            Case "UPDATELEDGERID"
                SPQuery = UpdateLedgerId(_cmd)
            Case "GETCOMPANYINFO"
                SPQuery = "SELECT * FROM TblCompany WHERE Id=1"
            Case "UPDATECOMPANYINFO"
                SPQuery = UpdateCompanyDetail(_cmd)
            Case "REMOVECOMPANYLOGO"
                SPQuery = "Update TblCompany Set LogoName='' Where [Id]=1"
            Case "LEDGERSUMMARY"
                SPQuery = "SELECT tblledger.transid, (select name from tblledgerids where ledgerid=tblledger.transid) As tname,  IIF((SUM(tblledger.debit)-SUM(tblledger.credit))>0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TDR, IIF((SUM(tblledger.debit)-SUM(tblledger.credit))<0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TCR FROM TblLedger Where tblledger.trantype='T' OR tblledger.trantype='U' OR tblledger.trantype='B' Group by tblledger.transid "
            Case "UPDATETAXCODE"
                SPQuery = UpdateTaxCodes(_cmd)
            Case "UPDATEVENDOR"
                SPQuery = UpdateVendor(_cmd)
            Case "ADDMONTH"
                SPQuery = UpdateFinancialYear(_cmd)
            Case "VATBETWEENTRANSACTION"
                SPQuery = VATBetweenTransaction(_cmd)
            Case "VATBEFORETRANSACTION"
                SPQuery = VATBeforeTransaction(_cmd)
            Case "VATQUERY1"
                SPQuery = CStr(_cmd.Parameter(0))
            Case "VATQUERY2"
                SPQuery = CStr(_cmd.Parameter(0))
            Case "VATDETAILS"
                SPQuery = CStr(_cmd.Parameter(0))
            Case "RECONCILETRANS"
                SPQuery = CStr(_cmd.Parameter(0))
            Case "UPDATERECONCILETRANS"
                SPQuery = ReconcileTransaction(_cmd)
            Case "GETCUSTOMERACTSUMMARY"
                SPQuery = GETCUSTOMERACTSUMMARY(_cmd)
            Case "GETVENDORACTSUMMARY"
                SPQuery = GETVENDORACTSUMMARY(_cmd)
            Case "LOADCUSTSUMMARY"
                SPQuery = LoadCustSummary(_cmd)
            Case "GETTRANSACTIONNO"
                SPQuery = GetTransactionNo(_cmd)
            Case "INSERTTBLTRANSACTION"
                SPQuery = InsertTransaction(_cmd)
            Case "INSERTTBLLEDGER"
                SPQuery = InsertLedger(_cmd)
            Case "INSERTTBLOUTSTANDING"
                SPQuery = InsertOutStanding(_cmd)
            Case "SYNTBLCUSTOMER" 'For Synchornization
                SPQuery = "Select * From TblCustomer Where CustomerId='" & CStr(_cmd.Parameter(0)) & "'"
            Case "SYNTBLVENDOR"
                SPQuery = "Select * From TblVendor Where VendorId='" & CStr(_cmd.Parameter(0)) & "'"
            Case "SYNTBLTRANSACTION"
                SPQuery = "Select * From TblTransaction Where ParentId=" & CInt(_cmd.Parameter(0)) & ""
            Case "SYNTBLLEDGER"
                'SPQuery = "Select * From TblLedger Where ParentId=" & CInt(_cmd.Parameter(0)) & " And TransId='" & CStr(_cmd.Parameter(1)) & "' And TranType='" & CStr(_cmd.Parameter(2)) & "'"
                SPQuery = "Select * From TblLedger Where MyAutoNo=" & CInt(_cmd.Parameter(0)) & ""
            Case "SYNTBLOUTSTANDING"
                'SPQuery = "Select * From TblOutStanding Where MyAuto=" & CInt(_cmd.Parameter(0)) & " And ParentId=" & CInt(_cmd.Parameter(1)) & " And TransId='" & CStr(_cmd.Parameter(2)) & "' And TranType='" & CStr(_cmd.Parameter(3)) & "'"
                SPQuery = "Select * From TblOutStanding Where OSNo=" & CInt(_cmd.Parameter(0)) & ""
            Case "UPDATEPAYSTATUS"
                SPQuery = UpdateLedgerPayStatus(_cmd)
            Case "GETLEDGERAUTONO"
                SPQuery = GetLedgerAutoNo(_cmd)
            Case "UPDATEOUTSTANDING"
                SPQuery = UpdateOutstanding(_cmd)
            Case "CHECKREFUND"
                SPQuery = ChkRefund(_cmd)
            Case "PAYSTATUSFLAG"
                SPQuery = PaystatusFlag(_cmd)
            Case "GETBANKBALANCE"
                SPQuery = GetBankBalance(_cmd)
            Case "SALESREFUND"
                SPQuery = SalesRefund(_cmd)
            Case "PURCHASEREFUND"
                SPQuery = PurchaseRefund(_cmd)
            Case "CUSTOMERINVOICEREFUND"
                SPQuery = CUSTOMERINVOICEREFUND(_cmd)
            Case "VENDORINVOICEREFUND"
                SPQuery = VENDORINVOICEREFUND(_cmd)
            Case "GETDCAGAINSTTRANSACTION"
                SPQuery = GetDCAgainstTransaction(_cmd)
            Case "REPORTS_CUSTOMERID"
                SPQuery = GetCustID()
            Case "REPORTS_VENDORID"
                SPQuery = GetVendID()
            Case "REPORTS_LEDGERID"
                SPQuery = GetLedgerID()
            Case "REPORTS_GETCUSTOMERSLIST"
                SPQuery = GetCustList()
            Case "REPORTS_GETCUSTOMERACTIVITY"
                SPQuery = GetCustActivity(_cmd)
            Case "REPORTS_GETSALESINVOICE"
                SPQuery = GetSalesInvoice(_cmd)
            Case "REPORTS_GETSALESRECEIPTS"
                SPQuery = GetSalesReceipts(_cmd)
            Case "REPORTS_GETVENDORSLIST"
                SPQuery = GetVendorList()
            Case "REPORTS_GETVENDORACTIVITY"
                SPQuery = GetVendorActivity(_cmd)
            Case "REPORTS_GETPURCHASEINVOICE"
                SPQuery = GetPurchaseInvoice(_cmd)
            Case "REPORTS_GETPURCHASEPAYMENTS"
                SPQuery = GetPurchasePayments(_cmd)
            Case "REPORTS_GETLEDGERBALANCE"
                SPQuery = GetLedgerBalance()
            Case "REPORTS_GETLEDGERACTIVITY"
                SPQuery = GetLedgerActivity(_cmd)
            Case "REPORTS_GETCOMPANYINFO"
                SPQuery = "SELECT Name,TCmpny.Street,(TCmpny.city + ' ' + TCmpny.Country) as CityCountry,Logo,LogoName from TblCompany as TCmpny"
            Case "GETSALESINVOICEDETAIL"
                SPQuery = GetSalesInvoiceDetail(_cmd)
            Case "GETPURCHASEINVOICEDETAIL"
                SPQuery = GetPurchaseInvoiceDetail(_cmd)
            Case "RECONCILEVAT"
                SPQuery = Reconcilevat(_cmd)
            Case "SELECTALLCUSTOMER" 'Import Export Quries
                SPQuery = "Select * from tblcustomer"
            Case "SELECTALLVENDOR"
                SPQuery = "select * from tblvendor"
            Case "SELECTCOMPANY"
                SPQuery = "select * from tblCompany"
            Case "SELECTBANK"
                SPQuery = "select * from tblbank"
            Case "SELECTF_YEAR"
                'SPQuery = "select myno,fmonth,fyear from tblfinancialyear"
                SPQuery = "select * from tblfinancialyear"
            Case "SELECTTAXID"
                SPQuery = "select top 9 * from tbltaxids"
            Case "SELECTTRANSACTION"
                SPQuery = "select parentid from tbltransaction"
            Case "SELECTLEDGER"
                SPQuery = "select distinct parentid from tblledger"
            Case "SELECTRECORD"
                SPQuery = selectrecord(_cmd)
            Case "SELECTLEDGERRECORD"
                SPQuery = selectledgerrecord(_cmd)
            Case "SELECTOUTSTANDING"
                SPQuery = "Select * from tbloutstanding"
            Case "RSELECTLEDGER"
                SPQuery = "Select * from tblledger"
            Case "RSELECTTRANSACTION"
                SPQuery = "Select * from tbltransaction"
            Case "DELETECUSTOMER"
                SPQuery = "delete from tblCustomer"
            Case "DELETECUSTOMER_PRO"
                SPQuery = "delete from Customer"
            Case "DELETESUPPLIER_PRO"
                SPQuery = "delete from supplier"
            Case "DELETETAX_CODE"
                SPQuery = "delete from tax_code"
            Case "DELETEBANK"
                SPQuery = "delete from bank"
            Case "DELETECOMPANY"
                SPQuery = "delete from company"
            Case "DELETEF_YEAR"
                SPQuery = "delete from F_year"
            Case "DELETETRANSACTION"
                SPQuery = "delete from [TRANSACTION]"
            Case "DELETELEDGER"
                SPQuery = "delete from LEDGER"
            Case "DELETEOUTSTANDING"
                SPQuery = "delete from outstanding"
            Case "INSERTALLCUSTOMER"
                SPQuery = insertallcustomer(_cmd)
            Case "INSERTALLVENDOR"
                SPQuery = insertallvendor(_cmd)
            Case "INSERTCOMPANY"
                SPQuery = insertcompany(_cmd)
            Case "INSERTBANK"
                SPQuery = insertbank(_cmd)
            Case "INSERTF_YEAR"
                SPQuery = insertf_year(_cmd)
            Case "UPDATETAXID"
                SPQuery = UpdateTaxID(_cmd)
            Case "INSERTTRANSACTION_PRO"
                SPQuery = INSERTTRANSACTION_PRO(_cmd)
            Case "INSERTLEDGER_PRO"
                SPQuery = INSERTLEDGER_PRO(_cmd)
            Case "INSERTOUTSTANDING_PRO"
                SPQuery = INSERTOUTSTANDING_PRO(_cmd)
            Case "EDELETECUSTOMER" 'Restore Quries
                SPQuery = "delete from tblCustomer"
            Case "EDELETEVENDOR"
                SPQuery = "delete from tblvendor"
            Case "EDELETETAX_CODE"
                SPQuery = "delete from tbltaxid"
            Case "EDELETEBANK"
                SPQuery = "delete from tblbank"
            Case "EDELETEF_YEAR"
                SPQuery = "delete from TblFinancialYear"
            Case "EDELETETRANSACTION"
                SPQuery = "delete from tblTRANSACTION"
            Case "EDELETELEDGER"
                SPQuery = "delete from tblLEDGER"
            Case "EDELETEOUTSTANDING"
                SPQuery = "delete from tbloutstanding"
            Case "EINSERTALLCUSTOMER"
                SPQuery = Einsertallcustomer(_cmd)
            Case "EINSERTALLVENDOR"
                SPQuery = Einsertallvendor(_cmd)
            Case "EINSERTCOMPANY"
                SPQuery = Eupdatecompany(_cmd)
            Case "EINSERTBANK"
                SPQuery = Einsertbank(_cmd)
            Case "EINSERTF_YEAR"
                SPQuery = Einsertf_year(_cmd)
            Case "EUPDATETAXID"
                SPQuery = EUpdateTaxID(_cmd)
            Case "EINSERTTRANSACTION"
                SPQuery = EINSERTTRANSACTION(_cmd)
            Case "EINSERTLEDGER"
                SPQuery = EINSERTLEDGER(_cmd)
            Case "EINSERTOUTSTANDING"
                SPQuery = EINSERTOUTSTANDING(_cmd)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "DELETEINVOICE"
                SPQuery = "delete from invoice"
            Case "DELETEINVOICE_DET"
                SPQuery = "delete from invoice_det"
            Case "DELETESALESOD"
                SPQuery = "delete from salesod"
            Case "DELETESALESORDER"
                SPQuery = "delete from salesorder"
            Case "DELETEPURACHASEOD"
                SPQuery = "delete from purchaseod"
            Case "DELETEPURCHASEORDER"
                SPQuery = "delete from purchaseorder"
            Case "DELETESUBSCRIPTION"
                SPQuery = "delete from subscription"
            Case "DELETESUBSCRIPTION_DET"
                SPQuery = "delete from SUBSCRIPTION_DET"
            Case "DELETERECURRING"
                SPQuery = "delete from recurring"
            Case "DELETEPRODUCT"
                SPQuery = "delete from product"
            Case "DELETEPRODUCT_DEF"
                SPQuery = "delete from PRODUCT_DEF"
            Case "DELETEPRODUCT_BOM"
                SPQuery = "delete from PRODUCT_BOM"
            Case "DELETEPROD_SALES"
                SPQuery = "delete from PROD_SALES"
            Case "DELETEPROD_LINK"
                SPQuery = "delete from PROD_LINKS"
            Case "DELETEPRODUCT_DISCOUNT"
                SPQuery = "delete from PROD_DISCOUNT"
            Case "DELETEPROD_ACTIVITY"
                SPQuery = "delete from PROD_ACTIVITY"
            Case "DELETEPREPAYMENT"
                SPQuery = "delete from PREPAYMENTS"
            Case "DELETECART_ITEMS"
                SPQuery = "delete from CART_ITEMS"
            Case "DELETECART_ORDER"
                SPQuery = "delete from CART_ORDER"
            Case "DELETEACCURAL"
                SPQuery = "delete from ACCRUAL"
            Case "DELETECUST_PROD_PRICE"
                SPQuery = "Delete From CUST_PROD_PRICE"
            Case "GETSYSOBJECTS"
                SPQuery = "SELECT * from msysobjects"
            Case "DELETEQUESTION"
                SPQuery = "Delete from question"
            Case "DELETEQANSWER"
                SPQuery = "Delete from qanswer"
            Case "DELETEQTABLE"
                SPQuery = "Delete from qtable"
            Case "DELETECATEGORY"
                SPQuery = "Delete from P_CATEGORY"
            Case "DELETENRECORD"
                SPQuery = "Delete from n_record"
            Case "UPDATENRECORD"
                SPQuery = UPDATENRECORD(_cmd)
            Case "NRECORD"
                SPQuery = "Select count(myauto) as Num from n_record"
            Case "SELECTNRECORD"
                'SPQuery = "select tbltransaction.parentid,tbltransaction.invtp,tbltransaction.invdate,tbltransaction.invnet,tbltransaction.invvat,tblledger.transid,tblledger.trantype From tbltransaction left join tblledger on tbltransaction.parentid=tblledger.parentid where tblledger.transid='10000'"
                SPQuery = "select tbltransaction.parentid,tbltransaction.invtp,tbltransaction.invdate,tbltransaction.invnet,tbltransaction.invvat,tblledger.transid,tblledger.trantype From tbltransaction left join tblledger on tbltransaction.parentid=tblledger.parentid where tblledger.transid='10000' or tblledger.transid='70200'  or tblledger.transid='49999' Or tblledger.transid='20000' order by transid"
            Case "FATCHRECORD"
                SPQuery = fatchrecord(_cmd)
            Case "UPNRECORD"
                SPQuery = UPNRECORD(_cmd)
            Case "INSERTN_RECORD"
                SPQuery = INSERTN_RECORD(_cmd)
            Case "GETPROMONTH"
                SPQuery = "select f_month,f_year from f_year"
            Case "INSERTCUSTOMER"
                SPQuery = INSERTCUSTOMER(_cmd)
            Case "GETCUSTOMERAUTONO"
                SPQuery = GetCustomerAutoNo(_cmd)
            Case "INSERTMCKEY"
                SPQuery = INSERTMCKEY(_cmd)
            Case "AMENDCUSTOMER"
                SPQuery = AmendCustomer(_cmd)
            Case "GETPDFINVNO"
                SPQuery = "select max(invno) as Ino from tblpdfinvoices"
            Case "INSERTPDF"
                SPQuery = INSERTPDF(_cmd)
            Case "UPDATEGROUPNO"
                SPQuery = UPDATEGROUPNO(_cmd)
            Case "CREDITCARDINFO"
                SPQuery = creditcardinfo(_cmd)
            Case "CNCREDITCARDINFO"
                SPQuery = CNCreditCardInfo(_cmd)
            Case "CNCREDITCARDINFOFORSC"
                SPQuery = CNCreditCardInfoFORSC(_cmd)
            Case "TBLSYNCHRONIZE"
                SPQuery = "SELECT * FROM TblSynchronize Where SynId=1"
            Case "GETREFUNDEDINVNO"
                SPQuery = GETREFUNDEDINVNO(_cmd)
            Case "INSERTSYNREFUNDINV"
                SPQuery = INSERTSYNREFUNDINV(_cmd)
            Case "GETTRANSACTIONPARENTID"
                SPQuery = "Select top 10 parentid from tbltransaction where Syn='N' order by parentid"

        End Select

        Return SPQuery

    End Function

    Public Shared Function GetUnSynchronizeDataQueryString(ByRef cmd As InfiniCommand) As String

        Dim SPName As String, SPQuery As String
        Dim dtLast As Date

        SPName = cmd.CommandName

        Select Case SPName
            Case "TblSynchronize"
                SPQuery = "SELECT * FROM TblSynchronize Where SynId=1"
            Case "TblBank"
                SPQuery = "SELECT * FROM TblBank Where BankId='70200' And Syn='N'" ' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#"
            Case "TblCompany"
                SPQuery = "SELECT * FROM TblCompany WHERE Id=1 And Syn='N'" ' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#"
            Case "TblCustomer"
                'SPQuery = "SELECT top 2 * FROM TblCustomer Where Syn='N' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "# Order By UniqueNo"
                SPQuery = "SELECT top 10 * FROM TblCustomer Where Syn='N' Order By UniqueNo"
            Case "TblFinancialYear"
                SPQuery = "SELECT * FROM TblFinancialYear Where Syn='N' Order By MyNo" 'And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#
            Case "TblLedger"
                SPQuery = "SELECT top 10 * FROM TblLedger Where Syn='N' Order By ParentId" ' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#
            Case "TblOutStanding"
                'SPQuery = "SELECT * FROM TblOutStanding Where Syn='N' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "# Order By MyAuto"
                SPQuery = "SELECT top 10 * FROM TblOutStanding Where Syn='N' order by parentid"
            Case "TblTaxIds"
                SPQuery = "Select Top 9 * From TblTaxIds Where Syn='N' Order By CountId" 'And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#
            Case "TblTransaction"
                SPQuery = "SELECT top 10 * FROM TblTransaction Where Syn='N' Order By ParentId" 'And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "# 
            Case "TblVendor"
                'SPQuery = "SELECT top 2 * FROM TblVendor Where Syn='N' And TranDateTime >=#" & CDate(cmd.Parameter(0)) & "# Order By UniqueNo"
                SPQuery = "SELECT top 10 * FROM TblVendor Where Syn='N' Order By UniqueNo"
            Case "TblCreditCardInfo"
                SPQuery = "Select * From TblCreditCardInfo Where Syn='N' Order By Inv_No" 'and TranDateTime >=#" & CDate(cmd.Parameter(0)) & "#
            Case "TBLSELECTALLCUSTOMER" 'Import Export Quries
                SPQuery = "Select * from tblcustomer where syn='N' Order By UniqueNo"
            Case "TBLSELECTALLVENDOR"
                SPQuery = "select * from tblvendor where syn='N' Order By UniqueNo"
            Case "TBLSELECTALLTRANSACTION"
                SPQuery = "select  * from tbltransaction where syn='N' Order By ParentID"
            Case "TBLSELECTALLLEDGER"
                SPQuery = "select * from tblledger where syn='N' Order By ParentID"
            Case "TBLSELECTALLOUTSTANDING"
                SPQuery = "select * from tbloutstanding where syn='N' order by myauto"


            Case "TBLTRANSACTION10"
                SPQuery = TblTransaction10(cmd)
            Case "TBLLEDGER10"
                SPQuery = TblLedger10(cmd)


        End Select

        Return SPQuery

    End Function

    Private Shared Function TblTransaction10(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "Select * from tbltransaction where parentid=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function TblLedger10(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "Select * from tblledger where parentid=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function
    Public Shared Function GetSynchronizeQueryString(ByRef cmd As InfiniCommand) As String

        Dim SPName As String, SPQuery As String

        SPName = cmd.CommandName

        Select Case SPName
            Case "TblBank"
                SPQuery = SynchronizeTblBank(cmd)
            Case "TblCompany"
                SPQuery = SynchronizeTblCompany(cmd)
            Case "TblFinancialYear"
                SPQuery = SynchronizeTblFinancialYear(cmd)
            Case "TblTaxIds"
                SPQuery = SynchronizeTblTaxIds(cmd)
            Case "TblCustomer"
                SPQuery = SynchronizeTblCustomer(cmd)
            Case "TblVendor"
                SPQuery = SynchronizeTblVendor(cmd)
            Case "TblTransaction"
                SPQuery = SynchronizeTblTransaction(cmd)
            Case "TblLedger"
                SPQuery = SynchronizeTblLedger(cmd)
            Case "TblOutStanding"
                SPQuery = SynchronizeTblOutStanding(cmd)

            Case "TblBankSynFlag"
                SPQuery = "UPDATE TblBank SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblCompanySynFlag"
                SPQuery = "UPDATE TblCompany SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblFinancialYearSynFlag"
                SPQuery = "UPDATE TblFinancialYear SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblTaxIdsSynFlag"
                SPQuery = "UPDATE TblTaxIds SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblCustomerSynFlag"
                SPQuery = TblCustomerSynFlag(cmd) '"UPDATE TblCustomer SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblVendorSynFlag"
                SPQuery = TblVendorSynFlag(cmd)
                'SPQuery = "UPDATE TblVendor SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblTransactionSynFlag"
                SPQuery = TblTransactionSynFlag(cmd)
                'SPQuery = "UPDATE TblTransaction SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblLedgerSynFlag"
                SPQuery = TblLedgerSynFlag(cmd)
                'SPQuery = "UPDATE TblLedger SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblOutStandingSynFlag"
                SPQuery = TblOutStandingSynFlag(cmd)
                'SPQuery = "UPDATE TblOutStanding SET Syn='Y', TranDateTime=#" & Date.Now & "#"
            Case "TblCreditCardInfoFlag"
                SPQuery = "UPDATE TblCreditCardInfo SET Syn='Y', TranDateTime=#" & Date.Now & "#"

        End Select

        Return SPQuery

    End Function

    Private Shared Function SynchronizeTblCustomer(ByRef _cmd As InfiniCommand) As String

        Dim ChRecord As String
        ChRecord = CStr(_cmd.Parameter(13))
        If ChRecord = "Y" Then
            _SqlStr = "UPDATE TblCustomer SET CustomerName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', PostalCode='" & CStr(_cmd.Parameter(5)) & "', ContactName='" & CStr(_cmd.Parameter(6)) & "',Email='" & CStr(_cmd.Parameter(7)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & ", DFlag='" & CStr(_cmd.Parameter(12)) & "', TaxId='T1', LedgerId='10000', TranDateTime=#" & Date.Now & "#, Syn='Y' WHERE CustomerId='" & CStr(_cmd.Parameter(0)) & "'"
        Else
            _SqlStr = "INSERT INTO TblCustomer(CustomerId,CustomerName,Street,Town,Country,PostalCode,ContactName,Email,Telephone,Fax,Vat,CreditLimit,DFlag,Syn,TranDateTime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','Y',#" & Date.Now & "#)"
        End If
        Return _SqlStr

    End Function

    Private Shared Function SynchronizeTblVendor(ByRef _cmd As InfiniCommand) As String

        Dim ChRecord As String
        ChRecord = CStr(_cmd.Parameter(13))
        If ChRecord = "Y" Then
            _SqlStr = "UPDATE TblVendor SET VendorName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', PostalCode='" & CStr(_cmd.Parameter(5)) & "', ContactName='" & CStr(_cmd.Parameter(6)) & "',Email='" & CStr(_cmd.Parameter(7)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & ", DFlag='" & CStr(_cmd.Parameter(12)) & "', TaxId='T1', LedgerId='20000', TranDateTime=#" & Date.Now & "#, Syn='Y' WHERE VendorId='" & CStr(_cmd.Parameter(0)) & "'"
        Else
            _SqlStr = "INSERT INTO TblVendor(VendorId,VendorName,Street,Town,Country,PostalCode,ContactName,Email,Telephone,Fax,Vat,CreditLimit,DFlag,Syn,TranDateTime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','Y',#" & Date.Now & "#)"
        End If
        Return _SqlStr

    End Function

    Private Shared Function SynchronizeTblTransaction(ByRef _cmd As InfiniCommand) As String

        Dim ChRecord As String
        ChRecord = CStr(_cmd.Parameter(13))
        If ChRecord = "Y" Then
            _SqlStr = "Update TblTransaction Set InvTP='" & CStr(_cmd.Parameter(1)) & "', TaxId='" & CStr(_cmd.Parameter(2)) & "', InvDate='" & CStr(_cmd.Parameter(3)) & "', InvDetails='" & CStr(_cmd.Parameter(4)) & "', InvNet=" & CDbl(_cmd.Parameter(5)) & ", InvVAT=" & CDbl(_cmd.Parameter(6)) & ", Refund='" & CStr(_cmd.Parameter(7)) & "', Mark='" & CStr(_cmd.Parameter(8)) & "', V='" & CStr(_cmd.Parameter(9)) & "', TranDate=#" & CDate(_cmd.Parameter(10)) & "#, Syn='Y', TranDateTime=#" & Date.Now & "#, GroupNo=" & CInt(_cmd.Parameter(11)) & ", rid=" & CInt(_cmd.Parameter(12)) & " Where ParentId=" & CInt(_cmd.Parameter(0)) & ""
        Else
            _SqlStr = "Insert Into TblTransaction (ParentId,InvTP,TaxId,InvDate,InvDetails,InvNet,InvVAT,Refund,Mark,V,TranDate,Syn,TranDateTime,GroupNo,rid) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "'," & CDbl(_cmd.Parameter(5)) & "," & CDbl(_cmd.Parameter(6)) & ",'" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "',#" & CDate(_cmd.Parameter(10)) & "#,'Y',#" & Date.Now & "#," & CInt(_cmd.Parameter(11)) & "," & CInt(_cmd.Parameter(12)) & ")"
        End If
        Return _SqlStr

    End Function

    Private Shared Function SynchronizeTblLedger(ByRef _cmd As InfiniCommand) As String

        Dim ChRecord As String
        ChRecord = CStr(_cmd.Parameter(9))
        If ChRecord = "Y" Then
            _SqlStr = "Update TblLedger Set Debit=" & CDbl(_cmd.Parameter(3)) & ", Credit=" & CDbl(_cmd.Parameter(4)) & ", TranRef='" & CStr(_cmd.Parameter(6)) & "', PayStatus='" & CStr(_cmd.Parameter(7)) & "', MyAutoNo=" & CInt(_cmd.Parameter(8)) & ", Syn='Y', TranDateTime=#" & Date.Now & "# Where MyAutoNo=" & CInt(_cmd.Parameter(0)) & ""
        Else
            _SqlStr = "Insert Into TblLedger (ParentId, ChildId, TransId, Debit, Credit, TranType, TranRef, PayStatus, MyAutoNo, Syn, TranDateTime) Values (" & CInt(_cmd.Parameter(0)) & "," & CInt(_cmd.Parameter(1)) & ",'" & CStr(_cmd.Parameter(2)) & "'," & CDbl(_cmd.Parameter(3)) & "," & CDbl(_cmd.Parameter(4)) & ",'" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "'," & CInt(_cmd.Parameter(8)) & ",'Y',#" & Date.Now & "#)"
        End If
        Return _SqlStr

    End Function

    Private Shared Function SynchronizeTblOutStanding(ByRef _cmd As InfiniCommand) As String

        Dim ChRecord As String
        ChRecord = CStr(_cmd.Parameter(14))
        If ChRecord = "Y" Then
            _SqlStr = "Update TblOutStanding Set Debit=" & CDbl(_cmd.Parameter(4)) & ", Credit=" & CDbl(_cmd.Parameter(5)) & ", Details='" & CStr(_cmd.Parameter(7)) & "', TranRef='" & CStr(_cmd.Parameter(8)) & "', OSDate='" & CStr(_cmd.Parameter(9)) & "', OSFrom=" & CInt(_cmd.Parameter(10)) & ", OSTo=" & CInt(_cmd.Parameter(11)) & ", OSFT='" & CStr(_cmd.Parameter(12)) & "', Syn='Y', TranDateTime=#" & Date.Now & "# Where OSNo=" & CInt(_cmd.Parameter(13)) & ""
        Else
            _SqlStr = "Insert Into TblOutStanding (MyAuto, ParentId, ChildId, TransId, Debit, Credit, TranType, Details, TranRef, OSDate, OSFrom, OSTo, OSFT, Syn, TranDateTime, OSNo) Values (" & CInt(_cmd.Parameter(0)) & "," & CInt(_cmd.Parameter(1)) & "," & CInt(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "'," & CDbl(_cmd.Parameter(4)) & "," & CDbl(_cmd.Parameter(5)) & ",'" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CInt(_cmd.Parameter(10)) & "," & CInt(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','Y',#" & Date.Now & "#," & CInt(_cmd.Parameter(13)) & ")"
        End If
        Return _SqlStr

    End Function

    Private Shared Function SynchronizeTblBank(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblBank SET BankName='" & CStr(_cmd.Parameter(1)) & "', Street= '" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', PostalCode='" & CStr(_cmd.Parameter(5)) & "', AccountName='" & CStr(_cmd.Parameter(6)) & "', AccountNumber='" & CStr(_cmd.Parameter(7)) & "', SortCode='" & CStr(_cmd.Parameter(8)) & "', Syn='Y', TranDateTime=#" & Date.Now & "# WHERE BankId ='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function SynchronizeTblCompany(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblCompany Set [Name]='" & CStr(_cmd.Parameter(0)) & "', Street='" & CStr(_cmd.Parameter(1)) & "', City='" & CStr(_cmd.Parameter(2)) & "', PostalCode='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', Telephone='" & CStr(_cmd.Parameter(5)) & "', Fax='" & CStr(_cmd.Parameter(6)) & "', Email='" & CStr(_cmd.Parameter(7)) & "', VatNo='" & CStr(_cmd.Parameter(8)) & "', LogoName='" & CStr(_cmd.Parameter(10)) & "', Syn='Y', TranDateTime=#" & Date.Now & "# Where [Id]=1"
        Return _SqlStr
    End Function

    Private Shared Function SynchronizeTblFinancialYear(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblFinancialYear SET FMonth='" & CStr(_cmd.Parameter(1)) & "', FYear='" & CStr(_cmd.Parameter(2)) & "', Syn='Y', TranDateTime=#" & Date.Now & "# WHERE MyNo=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function SynchronizeTblTaxIds(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblTaxIds SET TaxRate=" & CDbl(_cmd.Parameter(1)) & ", Syn='Y', TranDateTime=#" & Date.Now & "# WHERE TaxId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function UpdateSynchronizeId(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblSynchronize SET UserId=" & CInt(_cmd.Parameter(0)) & " WHERE SynId=1"
        Return _SqlStr
    End Function

    Private Shared Function UpdateSynchronizeDateTime() As String
        _SqlStr = "UPDATE TblSynchronize SET SynDateTime=#" & Date.Now & "# WHERE SynId=1"
        '_SqlStr = "UPDATE TblSynchronize SET SynDateTime=#" & Date.Now & "#, UserId=0 WHERE SynId=1"
        Return _SqlStr
    End Function

    Private Shared Function VendorDelete(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "Select TransId From TblLedger Where TransId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function CustomerDelete(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "Select TransId From TblLedger Where TransId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function CONFDELETE(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "Update TblVendor Set DFlag='Y' Where vendorId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function CONFDELETECUST(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "Update tblCustomer Set DFlag='Y' Where CustomerId=" & CInt(_cmd.Parameter(0)) & ""
        _SqlStr = "Update tblCustomer Set DFlag='Y' Where CustomerId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function GetDCAgainstTransaction(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TBLOutStanding.myauto, TBLOutStanding.parentid, TBLOutStanding.osTo, TBLOutStanding.osFT, TBLTRANSACTION.invtp, " & _
                  "TBLTRANSACTION.Refund,TBLTRANSACTION.invnet, TBLTRANSACTION.invvat " & _
                  "FROM TBLOutStanding LEFT JOIN TBLTRANSACTION ON TBLOutStanding.parentid = TBLTRANSACTION.parentid " & _
                  "WHERE TBLOutStanding.osTo=" & CStr(_cmd.Parameter(0)) & " AND TBLOutStanding.osFT='T' AND TBLTRANSACTION.invtp='" & CStr(_cmd.Parameter(1)) & "' AND TBLTRANSACTION.Refund='N'"
        Return _SqlStr
    End Function

    Private Shared Function GetTransactionNo(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT Max(ParentId)+1 As ParentNo FROM TblTransaction"
        Return _SqlStr
    End Function

    Private Shared Function CUSTOMERINVOICEREFUND(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT tbltransaction.parentid, tbltransaction.invtp, tblledger.tranref AS LTref, tblledger.trantype, tblledger.debit AS LDebit," & _
                  "tblledger.credit AS LCredit, tblledger.PayStatus, tblOutStanding.Debit AS ODebit, tblOutStanding.credit AS OCredit, " & _
                  "tblOutStanding.details, tblOutStanding.osFrom, tblOutStanding.osTo, tblOutStanding.myauto, tbltransaction_1.invtp as Fromtp " & _
                  "FROM ((tbltransaction LEFT JOIN tblledger ON tbltransaction.parentid = tblledger.parentid) LEFT JOIN tblOutStanding ON tblledger.myautono = tblOutStanding.myauto) LEFT JOIN tbltransaction AS tbltransaction_1 ON tblOutStanding.osFrom = tbltransaction_1.parentid " & _
                  "WHERE tbltransaction.parentid=1 AND tbltransaction.invtp='SI' AND tblledger.trantype='C'AND tblOutStanding.osFrom>0 AND tblOutStanding.osTo=0 " & _
                  "ORDER BY tblOutStanding.osFrom"
        Return _SqlStr
    End Function

    Private Shared Function VENDORINVOICEREFUND(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT tbltransaction.parentid, tbltransaction.invtp, tblledger.tranref AS LTref, tblledger.trantype, tblledger.debit AS LDebit," & _
                  "tblledger.credit AS LCredit, tblledger.PayStatus, tblOutStanding.Debit AS ODebit, tblOutStanding.credit AS OCredit, " & _
                  "tblOutStanding.details, tblOutStanding.osFrom, tblOutStanding.osTo, tblOutStanding.myauto, tbltransaction_1.invtp as Fromtp " & _
                  "FROM ((tbltransaction LEFT JOIN tblledger ON tbltransaction.parentid = tblledger.parentid) LEFT JOIN tblOutStanding ON tblledger.myautono = tblOutStanding.myauto) LEFT JOIN tbltransaction AS tbltransaction_1 ON tblOutStanding.osFrom = tbltransaction_1.parentid " & _
                  "WHERE tbltransaction.parentid=1 AND tbltransaction.invtp='PI' AND tblledger.trantype='S'AND tblOutStanding.osFrom>0 AND tblOutStanding.osTo=0 " & _
                  "ORDER BY tblOutStanding.osFrom"
        Return _SqlStr
    End Function

    Private Shared Function SalesRefund(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "SELECT TBLTRANSACTION.parentid, TBLTRANSACTION.invdetails, TBLTRANSACTION.invtp,TBLTransaction_1.invtp as fromtp, " & _
                  "TBLTRANSACTION.taxid, TBLTRANSACTION.invdate, TBLTRANSACTION.invnet, TBLTRANSACTION.invvat, TBLLEDGER.TransId, " & _
                  "TBLTRANSACTION.taxid,TBLLEDGER.PayStatus,TBLLEDGER.MYAUTONO,TBLOUTSTANDING.MYAUTO,TBLLEDGER.debit as LDebit, " & _
                  "TBLLEDGER.credit as LCredit, TBLOutStanding.Debit as oDebit,TBLOutStanding.credit as OCredit " & _
                  "FROM ((TBLTRANSACTION LEFT JOIN TBLLEDGER ON TBLTRANSACTION.parentid = TBLLEDGER.parentid) " & _
                  "LEFT JOIN TBLOutStanding ON TBLLEDGER.myautono = TBLOutStanding.myauto)" & _
                  "LEFT JOIN TBLtransaction as TBLTransaction_1 on TBLOutstanding.osfrom = TBLTransaction_1.parentid " & _
                  "WHERE TBLTRANSACTION.invtp='SI' AND TBLLEDGER.transid='" & CStr(_cmd.Parameter(0)) & "'  AND TBLOutStanding.osFrom>0 AND " & _
                  "TBLOutStanding.osTo=0 AND TBLOutStanding.osFT='F' AND TBLLEDGER.PayStatus='f' AND " & _
                  "TBLTRANSACTION.Refund='N' " & _
                  "GROUP BY TBLTRANSACTION.parentid, TBLTRANSACTION.invdetails,TBLTRANSACTION.invtp, " & _
                  "TBLTRANSACTION.taxid, TBLTRANSACTION.invdate, TBLTRANSACTION.invnet, TBLTRANSACTION.invvat, " & _
                  "TBLLEDGER.TransId, TBLTRANSACTION.taxid, TBLLEDGER.PayStatus, TBLLEDGER.MYAUTONO, " & _
                  "TBLOUTSTANDING.MYAUTO,TBLLEDGER.debit, TBLLEDGER.credit, TBLOutStanding.Debit, TBLOutStanding.credit, TBLtransaction_1.InvTP " & _
                  "Order by tbltransaction.parentid Asc"

        Return _SqlStr
    End Function

    Private Shared Function PurchaseRefund(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "SELECT TBLTRANSACTION.parentid, TBLTRANSACTION.invdetails, TBLTRANSACTION.invtp,TBLTransaction_1.invtp as fromtp, " & _
                  "TBLTRANSACTION.taxid, TBLTRANSACTION.invdate, TBLTRANSACTION.invnet, TBLTRANSACTION.invvat, TBLLEDGER.TransId, " & _
                  "TBLTRANSACTION.taxid,TBLLEDGER.PayStatus,TBLLEDGER.MYAUTONO,TBLOUTSTANDING.MYAUTO,TBLLEDGER.debit as LDebit, " & _
                  "TBLLEDGER.credit as LCredit, TBLOutStanding.Debit as oDebit,TBLOutStanding.credit as OCredit " & _
                  "FROM ((TBLTRANSACTION LEFT JOIN TBLLEDGER ON TBLTRANSACTION.parentid = TBLLEDGER.parentid) " & _
                  "LEFT JOIN TBLOutStanding ON TBLLEDGER.myautono = TBLOutStanding.myauto)" & _
                  "LEFT JOIN TBLtransaction as TBLTransaction_1 on TBLOutstanding.osfrom = TBLTransaction_1.parentid " & _
                  "WHERE TBLTRANSACTION.invtp='PI' AND TBLLEDGER.transid='" & CStr(_cmd.Parameter(0)) & "'  AND TBLOutStanding.osFrom>0 AND " & _
                  "TBLOutStanding.osTo=0 AND TBLOutStanding.osFT='F' AND TBLLEDGER.PayStatus='f' AND " & _
                  "TBLTRANSACTION.Refund='N' " & _
                  "GROUP BY TBLTRANSACTION.parentid, TBLTRANSACTION.invdetails,TBLTRANSACTION.invtp, " & _
                  "TBLTRANSACTION.taxid, TBLTRANSACTION.invdate, TBLTRANSACTION.invnet, TBLTRANSACTION.invvat, " & _
                  "TBLLEDGER.TransId, TBLTRANSACTION.taxid, TBLLEDGER.PayStatus, TBLLEDGER.MYAUTONO, " & _
                  "TBLOUTSTANDING.MYAUTO,TBLLEDGER.debit, TBLLEDGER.credit, TBLOutStanding.Debit, TBLOutStanding.credit, TBLtransaction_1.InvTP " & _
                  "Order by tbltransaction.parentid Asc"

        Return _SqlStr
    End Function

    Private Shared Function GETCUSTOMERACTSUMMARY(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "Select T.ParentID, T.InvTP, T.InvDate, T.InvDetails, Sum(L.Credit) As LCredit, Sum(L.Debit) As LDebit," & _
                  "IIF(L.PayStatus='*', (LCredit + LDebit),IIF(L.PayStatus='f', 0, ((LCredit + LDebit) - (select SUM(Debit) + SUM(Credit) from TblOutstanding as O  where O.myauto = L.myautono)))) As OSAmt,IIF(L.PayStatus='f','',L.PayStatus ) AS PayStatus From TblTransaction As T, TblLedger AS L " & _
                  "Where (L.ParentID=T.ParentID) and (L.Debit<>L.Credit)and (L.TransId='" & CStr(_cmd.Parameter(0)) & "') and (L.TranType='C') " & _
                  "Group By T.ParentID, T.InvTP, T.InvDate, T.InvDetails, L.PayStatus, L.MyAutoNo"
        Return _SqlStr
    End Function

    Private Shared Function GETVENDORACTSUMMARY(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "Select T.ParentID, T.InvTP, T.InvDate, T.InvDetails, Sum(L.Credit) As LCredit, Sum(L.Debit) As LDebit," & _
                 "IIF(L.PayStatus='*', (LCredit + LDebit),IIF(L.PayStatus='f', 0, ((LCredit + LDebit) - (select SUM(Debit) + SUM(Credit) from TblOutstanding as O  where O.myauto = L.myautono)))) As OSAmt,IIF(L.PayStatus='f','',L.PayStatus ) AS PayStatus From TblTransaction As T, TblLedger AS L " & _
                 "Where (L.ParentID=T.ParentID) and (L.Debit<>L.Credit) and (L.TransId='" & CStr(_cmd.Parameter(0)) & "') and (L.TranType='S') " & _
                 "Group By T.ParentID, T.InvTP, T.InvDate, T.InvDetails, L.PayStatus, L.MyAutoNo"

        Return _SqlStr
    End Function

    Private Shared Function LoadCustSummary(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "SELECT C1.CustomerID, UniqueNo, CustomerName, ContactName , Email, " & _
                    "IIF((SELECT SUM(DEBIT) -SUM(CREDIT)  FROM TBLLEDGER WHERE TRANTYPE='C' AND TRANSID=C1.CUSTOMERID),'0.00') AS SUMBAL " & _
                    "FROM TBLCUSTOMER C1 Where C1.DFlag='N'  ORDER BY UniqueNo"
        Return _SqlStr
    End Function

    Private Shared Function InsertNewCustomer(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "INSERT INTO TblCustomer(CustomerName, Email, TranDateTime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CDate(_cmd.Parameter(2)) & "')"
        Return _SqlStr

    End Function

    Private Shared Function CreateCustomerId(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "Select * From TblCustomer Where CustomerName='" & CStr(_cmd.Parameter(0)) & "' And Email='" & CStr(_cmd.Parameter(1)) & "' And TranDateTime=#" & CDate(_cmd.Parameter(2)) & "#"
        Return _SqlStr

    End Function

    Private Shared Function UpdateCustomer(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "UPDATE TblCustomer SET CustomerId=" & CInt(_cmd.Parameter(0)) & ", CustomerName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "' , Country='" & CStr(_cmd.Parameter(4)) & "' , PostalCode= '" & CStr(_cmd.Parameter(5)) & "', ContactName= '" & CStr(_cmd.Parameter(6)) & "',Email='" & CStr(_cmd.Parameter(7)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & " WHERE UniqueNo=" & CInt(_cmd.Parameter(12)) & ""
        Return _SqlStr

    End Function
    Private Shared Function AmendCustomer(ByRef _cmd As InfiniCommand) As String

        '_SqlStr = "UPDATE TblCustomer SET CustomerId='" & CStr(_cmd.Parameter(0)) & "', CustomerName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "' , Country='" & CStr(_cmd.Parameter(4)) & "' , PostalCode= '" & CStr(_cmd.Parameter(5)) & "', ContactName= '" & CStr(_cmd.Parameter(6)) & "',Email='" & CStr(_cmd.Parameter(7)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & ", TranDateTime=#" & CDate(_cmd.Parameter(12)) & "#, Syn='N' WHERE UniqueNo=" & CInt(_cmd.Parameter(13)) & ""
        _SqlStr = "UPDATE TblCustomer SET CustomerName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "' , Country='" & CStr(_cmd.Parameter(4)) & "' , PostalCode= '" & CStr(_cmd.Parameter(5)) & "', ContactName= '" & CStr(_cmd.Parameter(6)) & "',Email='" & CStr(_cmd.Parameter(7)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & ", TranDateTime=#" & Date.Now & "#, Syn='N' WHERE CustomerId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr

    End Function

    Private Shared Function InsertNewVendor(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "INSERT INTO TblVendor(VendorName, Email, TranDateTime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CDate(_cmd.Parameter(2)) & "')"
        Return _SqlStr

    End Function

    Private Shared Function CreateVendorId(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "Select * From TblVendor Where VendorName='" & CStr(_cmd.Parameter(0)) & "' And Email='" & CStr(_cmd.Parameter(1)) & "' And TranDateTime=#" & CDate(_cmd.Parameter(2)) & "#"
        Return _SqlStr

    End Function

    Private Shared Function UpdateVendor(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "UPDATE TblVendor SET VendorId='" & CStr(_cmd.Parameter(0)) & "', VendorName='" & CStr(_cmd.Parameter(1)) & "', Street='" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "' , Country='" & CStr(_cmd.Parameter(4)) & "' , PostalCode= '" & CStr(_cmd.Parameter(5)) & "', ContactName= '" & CStr(_cmd.Parameter(7)) & "',Email='" & CStr(_cmd.Parameter(6)) & "', Telephone= '" & CStr(_cmd.Parameter(8)) & "' , Fax= '" & CStr(_cmd.Parameter(9)) & "', Vat='" & CStr(_cmd.Parameter(10)) & "', CreditLimit=" & CDbl(_cmd.Parameter(11)) & ", TranDateTime=#" & CDate(_cmd.Parameter(12)) & "#, Syn='N' WHERE UniqueNo=" & CInt(_cmd.Parameter(13)) & ""
        Return _SqlStr

    End Function


    Private Shared Function GetCustomerDetail(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Select * from tblcustomer where customerid='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function GetVendorDetail(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Select * from tblvendor where vendorid='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function LoadCustomerSummary() As String

        '_SqlStr = "SELECT CustomerID, UniqueNo,CustomerName,ContactName,email,IIF( ISNULL( SUM(TBLLEDGER.DEBIT)-SUM(TBLLEDGER.CREDIT)),0,SUM(TBLLEDGER.DEBIT)-SUM(TBLLEDGER.CREDIT)) as Balance,SUM(TBLLEDGER.DEBIT)AS DEBIT,SUM(TBLLEDGER.CREDIT) AS CREDIT,TBLLEDGER.TRANTYPE  FROM TBLCUSTOMER LEFT JOIN TBLLEDGER ON TBLCUSTOMER.CUSTOMERID=TBLLEDGER.TRANSID GROUP BY  CUSTOMERID,CUSTOMERNAME,email,CONTACTNAME,UNIQUENO, TRANTYPE,DFLAG  HAVING TBLLEDGER.TRANTYPE='C' AND TBLCUSTOMER.DFLAG='N'"
        _SqlStr = "SELECT C1.CustomerId, C1.UniqueNo, C1.CustomerName, C1.ContactName, C1.Email, (SELECT IIf(IsNull(Sum(DEBIT)-Sum(CREDIT)),0,Sum(DEBIT)-Sum(CREDIT)) FROM TBLLEDGER WHERE TRANTYPE='C' AND TRANSID=C1.customerid) AS Balance FROM TBLCUSTOMER C1 Where C1.DFlag='N' Order By C1.UniqueNo"
        Return _SqlStr

    End Function

    Private Shared Function LoadVendorSummary() As String

        '_SqlStr = "SELECT VendorID, UniqueNo,VendorName,ContactName,Email,IIF( ISNULL( SUM(TBLLEDGER.DEBIT)-SUM(TBLLEDGER.CREDIT)),0,SUM(TBLLEDGER.CREDIT)-SUM(TBLLEDGER.DEBIT)) as Balance,SUM(TBLLEDGER.DEBIT)AS DEBIT,SUM(TBLLEDGER.CREDIT) AS CREDIT,TBLLEDGER.TRANTYPE  FROM TBLVendor LEFT JOIN TBLLEDGER ON TBLVendor.VENDORID=TBLLEDGER.TRANSID GROUP BY  VENDORID,vendorNAME,CONTACTNAME,Email,UNIQUENO, TRANTYPE,DFLAG HAVING TBLLEDGER.TRANTYPE='S' AND TBLvendor.DFLAG='N'"
        _SqlStr = "SELECT V1.VendorId, V1.UniqueNo, V1.VendorName, V1.ContactName, V1.Email, (SELECT IIf(IsNull(Sum(DEBIT)-Sum(CREDIT)),0,Sum(DEBIT)-Sum(CREDIT)) FROM TBLLEDGER WHERE TRANTYPE='S' AND TRANSID=V1.VendorId) AS Balance FROM TBLVendor V1 Where V1.DFlag='N' Order By V1.UniqueNo"
        Return _SqlStr

    End Function

    Private Shared Function LOADVENDORPAYMENTS(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "SELECT '0.00' as NullCol1,'0.00' as NullCol2, tbltransaction.ParentId, tbltransaction.InvTP, tbltransaction.TaxId, tbltransaction.InvDate, tbltransaction.InvDetails,tblledger.MyAutono,tblledger.PayStatus,(tblledger.Debit+tblledger.credit) AS Expr1, Sum(tbloutstanding.Debit+tbloutstanding.Credit) AS Expr2,  iif(not isnull((expr1-expr2)),expr1-expr2,expr1) AS Amount " & _
                  "FROM (tbltransaction LEFT JOIN tblledger ON tbltransaction.ParentId=tblledger.ParentId) LEFT JOIN tbloutstanding ON tblledger.MyAutoNo=tbloutstanding.MyAuto " & _
                  "GROUP BY tbltransaction.ParentId,(tblledger.Debit+tblledger.credit), tbltransaction.InvTP, tbltransaction.TaxId, tbltransaction.InvDate, tbltransaction.InvDetails,tblledger.MyAutono, tblledger.PayStatus, tbltransaction.Refund, tblledger.TransId " & _
                  "HAVING (((tbltransaction.InvTP)='PI' Or (tbltransaction.InvTP)='PC' Or (tbltransaction.InvTP)='PA') AND ((tblledger.PayStatus)='*' Or (tblledger.PayStatus)='P') AND ((tbltransaction.Refund)='N') AND ((tblledger.TransId)='" & CStr(_cmd.Parameter(0)) & "'))"
        Return _SqlStr

    End Function

    Private Shared Function LoadCustomerReceipts(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "SELECT '0.00' as NullCol1,'0.00' as NullCol2, tbltransaction.ParentId, tbltransaction.InvTP, tbltransaction.TaxId, tbltransaction.InvDate, tbltransaction.InvDetails,tblledger.MyAutono,tblledger.PayStatus,(tblledger.Debit+tblledger.credit) AS Expr1, Sum(tbloutstanding.Debit+tbloutstanding.Credit) AS Expr2,  iif(not isnull((expr1-expr2)),expr1-expr2,expr1) AS Amount " & _
                  "FROM (tbltransaction LEFT JOIN tblledger ON tbltransaction.ParentId=tblledger.ParentId) LEFT JOIN tbloutstanding ON tblledger.MyAutoNo=tbloutstanding.MyAuto " & _
                  "GROUP BY tbltransaction.ParentId,(tblledger.Debit+tblledger.credit), tbltransaction.InvTP, tbltransaction.TaxId, tbltransaction.InvDate, tbltransaction.InvDetails,tblledger.MyAutono, tblledger.PayStatus, tbltransaction.Refund, tblledger.TransId " & _
                  "HAVING (((tbltransaction.InvTP)='SI' Or (tbltransaction.InvTP)='SC' Or (tbltransaction.InvTP)='SA') AND ((tblledger.PayStatus)='*' Or (tblledger.PayStatus)='P') AND ((tbltransaction.Refund)='N') AND ((tblledger.TransId)='" & CStr(_cmd.Parameter(0)) & "'))"
        Return _SqlStr

    End Function

    Private Shared Function UpdateCompanyDetail(ByRef _cmd As InfiniCommand) As String

        Dim Check As Boolean
        Check = CBool(_cmd.Parameter(9))

        If Check = True Then
            _SqlStr = "Update TblCompany Set [Name]='" & CStr(_cmd.Parameter(0)) & "', Street='" & CStr(_cmd.Parameter(1)) & "', City='" & CStr(_cmd.Parameter(2)) & "', PostalCode='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', Telephone='" & CStr(_cmd.Parameter(5)) & "', Fax='" & CStr(_cmd.Parameter(6)) & "', Email='" & CStr(_cmd.Parameter(7)) & "', VatNo='" & CStr(_cmd.Parameter(8)) & "', LogoName='" & CStr(_cmd.Parameter(10)) & "', Syn='N', TranDateTime='" & CDate(QTime) & "' Where [Id]=1"
        Else
            _SqlStr = "Update TblCompany Set [Name]='" & CStr(_cmd.Parameter(0)) & "', Street='" & CStr(_cmd.Parameter(1)) & "', City='" & CStr(_cmd.Parameter(2)) & "', PostalCode='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', Telephone='" & CStr(_cmd.Parameter(5)) & "', Fax='" & CStr(_cmd.Parameter(6)) & "', Email='" & CStr(_cmd.Parameter(7)) & "', VatNo='" & CStr(_cmd.Parameter(8)) & "', Syn='N', TranDateTime='" & CDate(QTime) & "' Where [Id]=1"
        End If

        Return _SqlStr

    End Function

    Private Shared Function UpdateTaxCodes(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblTaxIds SET TaxRate=" & CDbl(_cmd.Parameter(1)) & ", Syn='N', TranDateTime='" & CDate(QTime) & "' WHERE TaxId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function UpdateFinancialYear(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblFinancialYear SET FMonth='" & CStr(_cmd.Parameter(1)) & "', FYear='" & CStr(_cmd.Parameter(2)) & "' WHERE MyNo=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function LoadBankDetail(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblLedgerIds.LedgerId,TblBank.BankId ,TblBank.BankName as BankName,TblBank.Street as Street ,TblBank.Town as Town,TblBank.Country as Country ,TblBank.PostalCode as PostalCode, TblBank.AccountName as AccountName,TblBank.AccountNumber as AccountNumber,TblBank.SortCode FROM TblLedgerIds LEFT JOIN TblBank ON TblLedgerIds.LedgerId = TblBank.BankId WHERE LedgerId='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function UpdateBankDetail(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblBank SET BankName='" & CStr(_cmd.Parameter(1)) & "', Street= '" & CStr(_cmd.Parameter(2)) & "', Town='" & CStr(_cmd.Parameter(3)) & "', Country='" & CStr(_cmd.Parameter(4)) & "', PostalCode='" & CStr(_cmd.Parameter(5)) & "', AccountName='" & CStr(_cmd.Parameter(6)) & "', AccountNumber='" & CStr(_cmd.Parameter(7)) & "', SortCode='" & CStr(_cmd.Parameter(8)) & "', Syn='N', TranDateTime='" & CDate(QTime) & "' WHERE BankId ='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function UpdateLedgerId(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblLedgerIds SET Name='" & CStr(_cmd.Parameter(1)) & "', Syn='N', TranDateTime='" & CDate(QTime) & "' WHERE LedgerId ='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function AddFinancialYear(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblFinancialYear Set FMonth='" & CStr(_cmd.Parameter(1)) & "', FYear='" & CStr(_cmd.Parameter(2)) & "' Where MyNo=" & Val(_cmd.Parameter(0))
        Return _SqlStr
    End Function

    Private Shared Function VATBetweenTransaction(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT Count(TBLTRANSACTION.Invtp) as Nor From TBLTRANSACTION WHERE (TBLTRANSACTION.Invtp<>'SR' And TBLTRANSACTION.Invtp<>'SA' And TBLTRANSACTION.Invtp<>'PP' And TBLTRANSACTION.Invtp<>'PA') AND (TBLTRANSACTION.V" & CStr(_cmd.Parameter(0)) & ")"
        Return _SqlStr
    End Function

    Private Shared Function VATBeforeTransaction(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT Count(TBLTRANSACTION.Invtp) as Nor From TBLTRANSACTION WHERE (TBLTRANSACTION.Invtp<>'SR' And TBLTRANSACTION.Invtp<>'SA' And TBLTRANSACTION.Invtp<>'PP' And TBLTRANSACTION.Invtp<>'PA' ) AND (TBLTRANSACTION.V" & CStr(_cmd.Parameter(0)) & ")"
        Return _SqlStr
    End Function

    Private Shared Function ReconcileTransaction(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblTransaction Set TblTransaction.v='R' where TblTransaction.ParentID=" & CInt(_cmd.Parameter(0)) & " AND TblTransaction.Invtp='" & CStr(_cmd.Parameter(1)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function Reconcilevat(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblTransaction Set TblTransaction.v='R' where TBLTRANSACTION.V = 'N' "
        Return _SqlStr
    End Function
    '------------ ClassGlobal Methods
    Private Shared Function InsertTransaction(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "INSERT INTO TblTransaction(ParentId, InvTP, TaxId, InvDate, InvDetails, InvNet, InvVat,Refund,Mark,V,Trandate, TranDateTime) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "', '" & CStr(_cmd.Parameter(2)) & "', '" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "', " & CDbl(_cmd.Parameter(5)) & ", " & CDbl(_cmd.Parameter(6)) & ",'" & CStr(_cmd.Parameter(7)) & "' ,'" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CDate(_cmd.Parameter(10)) & "','" & CDate(_cmd.Parameter(11)) & "')"
        Return _SqlStr
    End Function
    Private Shared Function InsertLedger(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "INSERT INTO TblLedger(ParentId,Childid,transid,debit,credit,trantype,tranref,paystatus,myautono, TranDateTime) values (" & CInt(_cmd.Parameter(0)) & "," & CInt(_cmd.Parameter(1)) & ", '" & CStr(_cmd.Parameter(2)) & "'," & CDbl(_cmd.Parameter(3)) & ", " & CDbl(_cmd.Parameter(4)) & ",'" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "' ,'" & CStr(_cmd.Parameter(7)) & "'," & CInt(_cmd.Parameter(8)) & ",'" & CDate(_cmd.Parameter(9)) & "')"
        Return _SqlStr
    End Function
    Private Shared Function InsertOutStanding(ByRef _cmd As InfiniCommand) As String
        '_SqlStr = "INSERT INTO TblOutStanding(MyAuto, ParentId, ChildId, TransId, Debit, Credit, TranType, Details, TranRef, OSDate, OSFrom, OSTo, OSFT, TranDateTime) values (" & CInt(_cmd.Parameter(0)) & "," & CInt(_cmd.Parameter(1)) & "," & CInt(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "'," & CDbl(_cmd.Parameter(4)) & "," & CDbl(_cmd.Parameter(5)) & ",'" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CInt(_cmd.Parameter(10)) & "," & CInt(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CDate(_cmd.Parameter(13)) & "')"
        _SqlStr = "INSERT INTO TblOutStanding(MyAuto, ParentId, ChildId, TransId, Debit, Credit, TranType, Details, TranRef, OSDate, OSFrom, OSTo, OSFT, TranDateTime, OSNo) values (" & CInt(_cmd.Parameter(0)) & "," & CInt(_cmd.Parameter(1)) & "," & CInt(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "'," & CDbl(_cmd.Parameter(4)) & "," & CDbl(_cmd.Parameter(5)) & ",'" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CInt(_cmd.Parameter(10)) & "," & CInt(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CDate(_cmd.Parameter(13)) & "'," & CInt(_cmd.Parameter(14)) & ")"
        Return _SqlStr
    End Function

    Private Shared Function UpdateLedgerPayStatus(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblLedger SET PayStatus='" & CStr(_cmd.Parameter(1)) & "', Syn='N', TranDateTime='" & CDate(_cmd.Parameter(2)) & "'  WHERE MyAutoNo =" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function
    Private Shared Function GetLedgerAutoNo(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT MyAutoNo From TblLedger Where ChildId=" & CInt(_cmd.Parameter(0)) & "  And TranType='" & CStr(_cmd.Parameter(1)) & "'"
        Return _SqlStr
    End Function
    Private Shared Function UpdateOutstanding(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblOutStanding Set Tranref='REFUND', Syn='N', TranDateTime='" & CDate(_cmd.Parameter(1)) & "' Where MyAuto=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function ChkRefund(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "Update TblTransaction Set Refund='" & CStr(_cmd.Parameter(0)) & "',Mark='" & CStr(_cmd.Parameter(1)) & "', Syn='N', TranDateTime='" & CDate(_cmd.Parameter(3)) & "'  Where parentid=" & CInt(_cmd.Parameter(2)) & ""
        Return _SqlStr
    End Function
    Private Shared Function PaystatusFlag(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TBlLEDGER.Parentid, Sum(TBlLEDGER.debit) AS LDEBIT, Sum(TBlLEDGER.credit) AS LCREDIT,TBlLEDGER.trantype," & _
                  "Sum(TblOutStanding.Debit) AS ODEBIT,Sum(TblOutStanding.Credit) AS OCredit,Sum(TBlLEDGER.debit)+Sum(TBlLEDGER.credit) AS LedgerVal,Sum(TblOutStanding.Debit) + Sum(TblOutStanding.Credit)  AS osvalue,TBlLEDGER.myautono " & _
                  "FROM TBlLEDGER LEFT JOIN TblOutStanding ON TBlLEDGER.myautono = TblOutStanding.myauto " & _
                  "GROUP BY TBlLEDGER.Parentid, TBlLEDGER.trantype, TBlLEDGER.myautono " & _
                  "HAVING (((TBlLEDGER.Parentid)=" & CInt(_cmd.Parameter(0)) & ") AND ((TBlLEDGER.trantype)='" & CStr(_cmd.Parameter(1)) & "'))"
        Return _SqlStr
    End Function
    Private Shared Function GetBankBalance(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TransId,(Sum(debit)-Sum(Credit)) AS tBal FROM TblLEDGER GROUP BY TransId HAVING (TransId='" & CStr(_cmd.Parameter(0)) & "')"
        Return _SqlStr
    End Function

    '---------- Class Reports Queries
    Private Shared Function GetCustID() As String
        _SqlStr = "Select CustomerID, CustomerName from TblCustomer where DFlag='N' order by CustomerID"
        Return _SqlStr
    End Function

    Private Shared Function GetVendID() As String
        _SqlStr = "Select VendorID, VendorName from TblVendor where DFlag='N' order by VendorID"
        Return _SqlStr
    End Function

    Private Shared Function GetLedgerID() As String
        _SqlStr = "Select distinct ( LedgerID) , Name  from TblLedgerIds, TblLedger where TblLedgerIds.LedgerId = TblLedger.TransId and (TblLedger.TranType = 'T' or TblLedger.TranType = 'U' or TblLedger.TranType = 'B' ) order by TblLedgerIds.LedgerID"
        Return _SqlStr
    End Function

    Private Shared Function GetCustList() As String
        _SqlStr = "SELECT CustomerId, CustomerName  ,  ( TblCustomer.Street +' '  +  Town  + ' '  + TblCustomer.Country + ' '  + TblCustomer.PostalCode )   As Address  ,  ContactName,TblCustomer.Email,TblCustomer.Telephone ,TblCustomer.Fax,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,TC.Logo from TblCompany as TC , TblCustomer where TblCustomer.DFlag = 'N' "
        Return _SqlStr
    End Function

    Private Shared Function GetCustActivity(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "Select T.ParentID, T.InvTP, T.InvDate, T.InvDetails,T.TaxID, Sum(L.Credit) As LCredit, Sum(L.Debit) As LDebit, IIF(L.PayStatus='*', (LCredit + LDebit),IIF(L.PayStatus='f', 0, ((LCredit + LDebit) - (select SUM(Debit) + SUM(Credit) from TblOutstanding as O  where O.myauto = L.myautono)))) As OSAmt,IIF(L.PayStatus='f','',L.PayStatus ) AS PayStatus,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry, First(TC.Logo) AS Logo from TblCompany TC, TblTransaction As T, TblLedger AS L Where (L.ParentID=T.ParentID) and (L.Debit<>L.Credit)and (L.TransId='" & CStr(_cmd.Parameter(0)) & "') and (L.TranType='C') Group By T.ParentID, T.InvTP, T.InvDate, T.InvDetails,T.TaxID, L.PayStatus, L.MyAutoNo,TC.Name,TC.Street,TC.city,TC.Country "
        Return _SqlStr

    End Function

    Private Shared Function GetSalesInvoice(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  ,TblTransaction.InvVAT , TblTransaction.InvNet , TblTransaction.InvTP  ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet As Gross, TblCompany.Name,TblCompany.Street,(TblCompany.city + ' ' + TblCompany.Country) as CityCountry,TblCompany.Logo from TblCompany, TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND (TblTRANSACTION.InvTP='SI') AND ((TblLEDGER.TransId)='" & CStr(_cmd.Parameter(0)) & "') AND ((TblLEDGER.trantype)='C') AND (TblTransaction.Mark='N') ) ORDER BY TBLTRANSACTION.ParentId"
        Return _SqlStr
    End Function

    Private Shared Function GetSalesReceipts(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  , TblTransaction.InvTP , TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet   As Gross ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,First(TC.Logo) from TblCompany TC, TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND ((TblLEDGER.TransId)='" & CStr(_cmd.Parameter(0)) & "' ) AND ((TblLEDGER.trantype)='C')) AND ((TblTransaction.Invtp)='SR') Group By TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  , TblTransaction.InvTP , TblLedger.TranType , (TblTransaction.InvVAT+TblTransaction.InvNet) ,TC.Name,TC.Street,TC.city, TC.Country "
        Return _SqlStr
    End Function

    Private Shared Function GetVendorList() As String
        _SqlStr = "SELECT  TV.VendorID ,TV.VendorName, TV.ContactName  ,TV.Email , ( TV.Street +' ' + TV.Town   +' ' + TV.Country  + ' ' + TV.PostalCode )  As Address ,TV.Telephone ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,TC.Logo from TblCompany as TC , TblVendor as TV where TV.DFlag = 'N' "
        Return _SqlStr
    End Function

    Private Shared Function GetVendorActivity(ByRef _cmd As InfiniCommand) As String

        _SqlStr = "Select T.ParentID, T.InvTP, T.InvDate, T.InvDetails,T.TaxID, Sum(L.Credit) As LCredit, Sum(L.Debit) As LDebit, IIF(L.PayStatus='*', (LCredit + LDebit),IIF(L.PayStatus='f', 0, ((LCredit + LDebit) - (select SUM(Debit) + SUM(Credit) from TblOutstanding as O  where O.myauto = L.myautono)))) As OSAmt,IIF(L.PayStatus='f','',L.PayStatus ) AS PayStatus,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry, First(TC.Logo) AS Logo from TblCompany TC, TblTransaction As T, TblLedger AS L Where (L.ParentID=T.ParentID) and (L.Debit<>L.Credit)and (L.TransId='" & CStr(_cmd.Parameter(0)) & "') and (L.TranType='S') Group By T.ParentID, T.InvTP, T.InvDate, T.InvDetails,T.TaxID, L.PayStatus, L.MyAutoNo,TC.Name,TC.Street,TC.city,TC.Country "
        Return _SqlStr

    End Function

    Private Shared Function GetPurchaseInvoice(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId ,TblTransaction.InvDetails  ,TblTransaction.InvVAT , TblTransaction.InvNet , TblTransaction.InvTP ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet   As Gross ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,TC.Logo from TblCompany as TC, TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId])  AND ((TblLEDGER.TransId)='" & CStr(_cmd.Parameter(0)) & "' ) AND ((TblLEDGER.trantype)='S') AND ((TblTRANSACTION.InvTP)='PI')  AND (TblTransaction.Mark='N') ) ORDER BY TBLTRANSACTION.ParentId"
        Return _SqlStr
    End Function

    Private Shared Function GetPurchasePayments(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate , TblTransaction.ParentId , TblTransaction.InvDetails  , TblTransaction.InvTP ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet As Gross ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,First(TC.Logo) from TblCompany as TC, TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND ((TblLEDGER.TransId)='" & CStr(_cmd.Parameter(0)) & "' ) AND ((TblLEDGER.trantype)='S')) AND((TblTransaction.Invtp)='PP') Group By TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  , TblTransaction.InvTP , TblLedger.TranType , (TblTransaction.InvVAT+TblTransaction.InvNet) ,TC.Name,TC.Street,TC.city, TC.Country "
        Return _SqlStr
    End Function

    Private Shared Function GetLedgerBalance() As String
        _SqlStr = "SELECT tblledger.transid, (select name from tblledgerids where ledgerid=tblledger.transid) As tname,  IIF((SUM(tblledger.debit)-SUM(tblledger.credit))>0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TDR, IIF((SUM(tblledger.debit)-SUM(tblledger.credit))<0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TCR ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,First(TC.Logo) AS Logo from TblCompany as TC , TblLedger Where tblledger.trantype='T' OR tblledger.trantype='U' OR tblledger.trantype='B' Group by TC.Name,TC.Street,TC.city,TC.Country,tblledger.transid "
        Return _SqlStr
    End Function

    Private Shared Function GetLedgerActivity(ByRef _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TL.ParentId , TL.TransId , SUM (TL.Debit) As Debit ,SUM (TL.Credit) AS Credit , TL.TRANTYPE ,  TT.TaxId , TT.Invtp, TT.InvDate, InvDetails, iif(SUM (TL.Debit)  > SUM (TL.Credit) ,SUM (TL.Debit)  ,SUM (TL.Credit) ) as Amount ,TC.Name,TC.Street,(TC.city + ' ' + TC.Country) as CityCountry,First(TC.Logo) AS Logo from TblCompany as TC, TblLedger as TL  INNER JOIN  TblTransaction as TT  ON TL.ParentID=TT.ParentId WHERE ( TL.TRANSID='" & CStr(_cmd.Parameter(0)) & "' AND TL.TRANTYPE IN ('S' ,'T','U' , 'B') ) GROUP By TL.ParentId , TL.TransId , TT.TaxId ,TL.TRANTYPE , TT.Invtp , TT.InvDate , InvDetails,TC.Name,TC.Street,TC.city,TC.Country Order By TL.ParentID "
        Return _SqlStr
    End Function
    Private Shared Function GetSalesInvoiceDetail(ByRef _cmd As InfiniCommand) As String
        'ORG
        _SqlStr = "SELECT TblTransaction.TaxId, TblTransaction.InvDetails, Sum(TblTransaction.InvNet) AS InvNet, Sum(TblTransaction.InvVAT) AS InvVAT, TblLedger.ParentId, (TblTransaction.InvNet + TblTransaction.InvVat) AS Amount, TblCustomer.CustomerId, TblCustomer.CustomerName, TblCustomer.Street AS TStreet, TblCustomer.Town, TblCustomer.Country AS TCountry, TblCustomer.PostalCode, TblCustomer.ContactName, TblCustomer.Email, TblCustomer.Telephone, TblCustomer.Fax, TblCustomer.Vat, TblCustomer.CreditLimit, TblCompany.Name, TblCompany.Street, First(TblCompany.Logo) AS Logo, (TblCompany.City + ' ' + TblCompany.Country) as citycountry FROM TblCompany, TblCustomer INNER JOIN (TblLedger INNER JOIN TblTransaction ON TblLedger.ParentId = TblTransaction.ParentId) ON Cstr(TblCustomer.CustomerId) = TblLedger.TransId WHERE (((TblLedger.TransId)='" & CStr(_cmd.Parameter(0)) & "') AND ((TblTransaction.ParentId)=" & CStr(_cmd.Parameter(1)) & "))"
        Dim i As Integer
        For i = 2 To _cmd.Parameters.Count - 1
            _SqlStr = _SqlStr & " OR (((TblTransaction.ParentId)=" & CStr(_cmd.Parameter(i)) & "))"
        Next
        _SqlStr = _SqlStr & " GROUP BY TblTransaction.TaxId, TblTransaction.InvDetails, TblLedger.ParentId, TblCustomer.CustomerId, TblCustomer.CustomerName, TblCustomer.Street, TblCustomer.Town, TblCustomer.Country, TblCustomer.PostalCode, TblCustomer.ContactName, TblCustomer.Email, TblCustomer.Telephone, TblCustomer.Fax, TblCustomer.Vat, TblCustomer.CreditLimit, TblCompany.Name, TblCompany.Street, TblCompany.City, TblCompany.Country ,(TblTransaction.InvNet + TblTransaction.InvVat) order by TblLedger.ParentId"
        Return _SqlStr

    End Function

    Private Shared Function GetPurchaseInvoiceDetail(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT TblTransaction.TaxId, TblTransaction.InvDetails, Sum(TblTransaction.InvNet) AS InvNet, Sum(TblTransaction.InvVAT) AS InvVAT, TblLedger.ParentId, (TblTransaction.InvNet + TblTransaction.InvVat) AS Amount, TblVendor.VendorId as CustomerId, TblVendor.VendorName as CustomerName, TblVendor.Street AS TStreet, TblVendor.Town, TblVendor.Country AS TCountry, TblVendor.PostalCode, TblVendor.ContactName, TblVendor.Email, TblVendor.Telephone, TblVendor.Fax, TblVendor.Vat, TblVendor.CreditLimit, TblCompany.Name, TblCompany.Street, First(TblCompany.Logo) AS Logo, (TblCompany.City + ' ' + TblCompany.Country) as citycountry FROM TblCompany, TblVendor INNER JOIN (TblLedger INNER JOIN TblTransaction ON TblLedger.ParentId = TblTransaction.ParentId) ON TblVendor.VendorId = TblLedger.TransId WHERE (((TblLedger.TransId)='" & CStr(_cmd.Parameter(0)) & "') AND ((TblTransaction.ParentId)=" & CStr(_cmd.Parameter(1)) & "))"
        Dim i As Integer
        For i = 2 To _cmd.Parameters.Count - 1
            _SqlStr = _SqlStr & " OR (((TblTransaction.ParentId)=" & CStr(_cmd.Parameter(i)) & "))"
        Next
        _SqlStr = _SqlStr & " GROUP BY TblTransaction.TaxId, TblTransaction.InvDetails, TblLedger.ParentId, TblVendor.VendorId, TblVendor.VendorName, TblVendor.Street, TblVendor.Town, TblVendor.Country, TblVendor.PostalCode, TblVendor.ContactName, TblVendor.Email, TblVendor.Telephone, TblVendor.Fax, TblVendor.Vat, TblVendor.CreditLimit, TblCompany.Name, TblCompany.Street, TblCompany.City, TblCompany.Country ,(TblTransaction.InvNet + TblTransaction.InvVat) order by TblLedger.ParentId"
        Return _SqlStr
    End Function

    'Import Export Quries
    Public Shared Function insertallcustomer(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "INSERT INTO Customer(ac, Name, Street, Town, Country, PostCode, Cont_Name, vEmail, Telephone, Fax, Vat, Credit_Limit,std_tax_code,nc) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function insertallvendor(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "INSERT INTO supplier(ac, Name, Street, Town, Country, PostCode, Cont_Name, Telephone, Fax, Vat, Credit_Limit,std_tax_code,nc) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CDbl(_cmd.Parameter(10)) & ",'" & CStr(_cmd.Parameter(11)) & "','" & CStr(_cmd.Parameter(12)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function insertcompany(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE company set comp_Name='" & CStr(_cmd.Parameter(0)) & "',Street1='" & CStr(_cmd.Parameter(1)) & "',Twon='" & CStr(_cmd.Parameter(2)) & "',p_code='" & CStr(_cmd.Parameter(3)) & "',country='" & CStr(_cmd.Parameter(4)) & "',telephone='" & CStr(_cmd.Parameter(5)) & "',Fax='" & CStr(_cmd.Parameter(6)) & "',vat_no='" & CStr(_cmd.Parameter(7)) & "',vlogo='" & CStr(_cmd.Parameter(8)) & "' where id=1"
        Return _SqlStr
    End Function

    Public Shared Function insertbank(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update bank set name='" & CStr(_cmd.Parameter(2)) & "',street='" & CStr(_cmd.Parameter(3)) & "',town='" & CStr(_cmd.Parameter(4)) & "',country='" & CStr(_cmd.Parameter(5)) & "',post_code='" & CStr(_cmd.Parameter(6)) & "',acct_name='" & CStr(_cmd.Parameter(7)) & "',acct_number='" & CStr(_cmd.Parameter(8)) & "',sort_code='" & CStr(_cmd.Parameter(9)) & "',ac_type='" & CStr(_cmd.Parameter(10)) & "' where ac='" & CStr(_cmd.Parameter(0)) & "'"
        '_SqlStr = "update bank set nc='" & CStr(_cmd.Parameter(1)) & "',name='" & CStr(_cmd.Parameter(2)) & "',street='" & CStr(_cmd.Parameter(3)) & "',town='" & CStr(_cmd.Parameter(4)) & "',country='" & CStr(_cmd.Parameter(5)) & "',post_code='" & CStr(_cmd.Parameter(6)) & "',acct_number='" & CStr(_cmd.Parameter(7)) & "',,acct_number='" & CStr(_cmd.Parameter(8)) & "',sort_code='" & CStr(_cmd.Parameter(9)) & "',ac_type='" & CStr(_cmd.Parameter(10)) & "' where ac='" & CStr(_cmd.Parameter(0)) & "'"
        ' _SqlStr = "INSERT INTO bank(ac,nc,name,street,town,country,post_code,acct_name,acct_number,sort_code,ac_type) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function insertf_year(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE F_YEAR set f_month='" & CStr(_cmd.Parameter(1)) & "',f_year='" & CStr(_cmd.Parameter(2)) & "' where myno=" & CInt(_cmd.Parameter(0)) & ""
        '_SqlStr = "INSERT INTO F_YEAR(myno,f_month,f_year) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function UpdateTaxID(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update tax_code set rate='" & CStr(_cmd.Parameter(1)) & "' where code='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Public Shared Function selectrecord(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "select * from tbltransaction right join tblledger on tbltransaction.parentid=tblledger.parentid where tblledger.parentid=" & CInt(_cmd.Parameter(0)) & " and ((tblledger.trantype='C') or (tblledger.trantype='B') or (tblledger.trantype='S') or (tblledger.tranref='REFUND'))"
        Return _SqlStr

    End Function

    Public Shared Function selectledgerrecord(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "select * from tbltransaction right join tblledger on tbltransaction.parentid=tblledger.parentid where tblledger.parentid=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr

    End Function

    Public Shared Function INSERTTRANSACTION_PRO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into [Transaction](record_no,activity_no,tran_no,ac,nc,tc,details,net,vat,tp,v,transfor,w,markw,[date],ref) values (" & CDbl(_cmd.Parameter(0)) & "," & CDbl(_cmd.Parameter(1)) & "," & CDbl(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "'," & CDbl(_cmd.Parameter(7)) & "," & CDbl(_cmd.Parameter(8)) & ",'" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "','" & CStr(_cmd.Parameter(11)) & "','" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "','" & CStr(_cmd.Parameter(14)) & "','" & CStr(_cmd.Parameter(15)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function INSERTLEDGER_PRO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into Ledger(record_no,activity_no,tran_no,nc,debit,credit,tran_type,refn,pay_status,myauto,details) values (" & CDbl(_cmd.Parameter(0)) & "," & CDbl(_cmd.Parameter(1)) & "," & CDbl(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "'," & CDbl(_cmd.Parameter(4)) & "," & CDbl(_cmd.Parameter(5)) & ",'" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "'," & CDbl(_cmd.Parameter(9)) & ",'" & CStr(_cmd.Parameter(10)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function INSERTOUTSTANDING_PRO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into outstanding(myauto,r_no,a_no,t_no,nc,tran_type,debit,credit,refn,details,net,odate,[from],[to],ft) values (" & CDbl(_cmd.Parameter(0)) & "," & CDbl(_cmd.Parameter(1)) & "," & CDbl(_cmd.Parameter(2)) & "," & CDbl(_cmd.Parameter(3)) & ",'" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "'," & CDbl(_cmd.Parameter(6)) & "," & CDbl(_cmd.Parameter(7)) & ",'" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CDbl(_cmd.Parameter(10)) & ",'" & CStr(_cmd.Parameter(11)) & "'," & CDbl(_cmd.Parameter(12)) & "," & CDbl(_cmd.Parameter(13)) & ",'" & CStr(_cmd.Parameter(14)) & "')"
        Return _SqlStr
    End Function

    'Restore Process
    Public Shared Function Einsertallcustomer(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "INSERT INTO tblCustomer(Uniqueno,CustomerID, CustomerName, Street, Town, Country, PostalCode, ContactName, Email, Telephone, Fax, Vat, CreditLimit,syn,trandatetime) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "','" & CStr(_cmd.Parameter(11)) & "'," & CDbl(_cmd.Parameter(12)) & ",'" & CStr(_cmd.Parameter(13)) & "','" & CStr(_cmd.Parameter(14)) & "')"
        _SqlStr = "INSERT INTO tblCustomer(CustomerID, CustomerName, Street, Town, Country, PostalCode, ContactName, Email, Telephone, Fax, Vat, CreditLimit,syn,trandatetime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function Einsertallvendor(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "INSERT INTO tblvendor(Uniqueno,VendorID, VendorName, Street, Town, Country, PostalCode, ContactName, Telephone, Fax, Vat, CreditLimit,Email,syn,trandatetime) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "','" & CStr(_cmd.Parameter(14)) & "')"
        _SqlStr = "INSERT INTO tblvendor(VendorID, VendorName, Street, Town, Country, PostalCode, ContactName, Telephone, Fax, Vat, CreditLimit,Email,syn,trandatetime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CDbl(_cmd.Parameter(10)) & ",'" & CStr(_cmd.Parameter(11)) & "','" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function Eupdatecompany(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "UPDATE tblcompany set Name='" & CStr(_cmd.Parameter(0)) & "',Street='" & CStr(_cmd.Parameter(1)) & "',City='" & CStr(_cmd.Parameter(2)) & "',postalcode='" & CStr(_cmd.Parameter(3)) & "',country='" & CStr(_cmd.Parameter(4)) & "',telephone='" & CStr(_cmd.Parameter(5)) & "',Fax='" & CStr(_cmd.Parameter(6)) & "',vatno='" & CStr(_cmd.Parameter(7)) & "',logo='" & CStr(_cmd.Parameter(8)) & "',email='" & CStr(_cmd.Parameter(9)) & "' where id=1"
        _SqlStr = "UPDATE tblcompany set Name='" & CStr(_cmd.Parameter(0)) & "',Street='" & CStr(_cmd.Parameter(1)) & "',City='" & CStr(_cmd.Parameter(2)) & "',postalcode='" & CStr(_cmd.Parameter(3)) & "',country='" & CStr(_cmd.Parameter(4)) & "',telephone='" & CStr(_cmd.Parameter(5)) & "',Fax='" & CStr(_cmd.Parameter(6)) & "',vatno='" & CStr(_cmd.Parameter(7)) & "',email='" & CStr(_cmd.Parameter(8)) & "' where id=1"
        Return _SqlStr
    End Function

    Public Shared Function Einsertbank(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "INSERT INTO tblbank(BankID,BankName,street,town,country,postalcode,AccountName,AccountNumber,sortcode,syn,trandatetime) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "')"
        _SqlStr = "UPDATE tblbank set bankname='" & CStr(_cmd.Parameter(1)) & "',Street='" & CStr(_cmd.Parameter(2)) & "',town='" & CStr(_cmd.Parameter(3)) & "',Country='" & CStr(_cmd.Parameter(4)) & "',Postalcode='" & CStr(_cmd.Parameter(5)) & "',AccountName='" & CStr(_cmd.Parameter(6)) & "',AccountNumber='" & CStr(_cmd.Parameter(7)) & "',sortcode='" & CStr(_cmd.Parameter(8)) & "',syn='" & CStr(_cmd.Parameter(9)) & "',trandatetime='" & CStr(_cmd.Parameter(10)) & "' where Bankid='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Public Shared Function Einsertf_year(ByVal _cmd As InfiniCommand) As String
        '_SqlStr = "INSERT INTO tblFINANCIALYEAR(myno,fmonth,fyear,syn,trandatetime) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "')"
        _SqlStr = "UPDATE tblFINANCIALYEAR set fmonth='" & CStr(_cmd.Parameter(1)) & "', fyear='" & CStr(_cmd.Parameter(2)) & "',syn='" & CStr(_cmd.Parameter(3)) & "',trandatetime='" & CDate(_cmd.Parameter(4)) & "' where myno=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Public Shared Function EUpdateTaxID(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update tbltaxIDs set taxrate='" & CStr(_cmd.Parameter(1)) & "' where taxid='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Public Shared Function EINSERTTRANSACTION(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into tblTransaction(parentid,invtp,taxid,invdate,invdetails,invnet,invvat,refund,mark,v,trandate,syn,trandatetime) values (" & CDbl(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "'," & CDbl(_cmd.Parameter(5)) & "," & CDbl(_cmd.Parameter(6)) & ",'" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CDate(_cmd.Parameter(10)) & "','" & CStr(_cmd.Parameter(11)) & "','" & CStr(_cmd.Parameter(12)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function EINSERTLEDGER(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into tblLedger(Parentid,Childid,transid,debit,credit,trantype,tranref,paystatus,myautono,syn,trandatetime) values (" & CDbl(_cmd.Parameter(0)) & "," & CDbl(_cmd.Parameter(1)) & ",'" & CStr(_cmd.Parameter(2)) & "'," & CDbl(_cmd.Parameter(3)) & "," & CDbl(_cmd.Parameter(4)) & ",'" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function EINSERTOUTSTANDING(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into Tbloutstanding(myauto,Parentid,Childid,transid,debit,credit,trantype,details,tranref,osdate,osfrom,osto,osft,syn,trandatetime) values (" & CDbl(_cmd.Parameter(0)) & "," & CDbl(_cmd.Parameter(1)) & "," & CDbl(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "'," & CDbl(_cmd.Parameter(4)) & "," & CDbl(_cmd.Parameter(5)) & ",'" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "'," & CDbl(_cmd.Parameter(10)) & "," & CDbl(_cmd.Parameter(11)) & ",'" & CStr(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "','" & CStr(_cmd.Parameter(14)) & "')"
        Return _SqlStr
    End Function

    Public Shared Function UPDATENRECORD(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update n_record set actual='0.000' where myauto=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Public Shared Function fatchrecord(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "select actual from n_record where nc='" & CStr(_cmd.Parameter(0)) & "' and month='" & CStr(_cmd.Parameter(1)) & "' and year='" & CStr(_cmd.Parameter(2)) & "'"
        Return _SqlStr
    End Function
    Public Shared Function UPNRECORD(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update n_record set actual=" & CDbl(_cmd.Parameter(3)) & " where nc='" & CStr(_cmd.Parameter(0)) & "' and year='" & CStr(_cmd.Parameter(1)) & "' and month='" & CStr(_cmd.Parameter(2)) & "'"
        Return _SqlStr
    End Function

    Public Shared Function INSERTN_RECORD(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "INSERT INTO n_record(nc,[year],[month]) values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "')"
        '' _SqlStr = "INSERT INTO n_record(nc,[year],[month],actual)  ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "'," & CDbl(_cmd.Parameter(3)) & ")"
        Return _SqlStr
    End Function
    Public Shared Function INSERTCUSTOMER(ByVal _cmd As InfiniCommand) As String

        _SqlStr = "INSERT INTO tblcustomer(Email,TranDateTime,customername) values ('" & CStr(_cmd.Parameter(0)) & "','" & CDate(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "')"
        Return _SqlStr
    End Function
    Public Shared Function GetCustomerAutoNo(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "select UniqueNo from tblcustomer where trandatetime=" & "#" & CDate(_cmd.Parameter(0)) & "#" & " and email= '" & CStr(_cmd.Parameter(1)) & "'"
        Return _SqlStr
    End Function

    Public Shared Function INSERTMCKEY(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into mckey(logid,logkey,logtype) values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "')"
        Return _SqlStr
    End Function
    Public Shared Function INSERTPDF(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into tblpdfinvoices(datafile,invdate,invtp,invno,custid,invnet,trandatetime) Values ('" & CStr(_cmd.Parameter(0)) & "','" & CDate(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "'," & CInt(_cmd.Parameter(3)) & ",'" & CStr(_cmd.Parameter(4)) & "','" & CDbl(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "')"
        Return _SqlStr
    End Function
    Public Shared Function UPDATEGROUPNO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "update tbltransaction set groupno=" & CInt(_cmd.Parameter(0)) & " where parentid=" & CInt(_cmd.Parameter(1)) & ""
        Return _SqlStr
    End Function
    Private Shared Function InsertOnlineCustomer(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into tblcustomer(CustomerId,CustomerName,Street,Town,Country,PostalCode,ContactName,Email,Telephone,Fax,Vat,CreditLimit,TranDateTime) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CDate(_cmd.Parameter(12)) & "')"
        Return _SqlStr
    End Function
    Private Shared Function InsertOnlineVendor(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into tblvendor(VendorId,VendorName,Street,Town,Country,PostalCode,ContactName,Email,Telephone,Fax,Vat,CreditLimit,TranDateTime) Values ('" & CStr(_cmd.Parameter(0)) & "','" & CStr(_cmd.Parameter(1)) & "','" & CStr(_cmd.Parameter(2)) & "','" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "'," & CDbl(_cmd.Parameter(11)) & ",'" & CDate(_cmd.Parameter(12)) & "')"
        Return _SqlStr
    End Function

    Private Shared Function creditcardinfo(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into TblCreditCardInfo(inv_no,inv_date,Debit,ShippingComp,TrackingNumber,DeliveryAddress,CardType,CardHoldername,CardNumber,ExpiryDate,SecurityCode,CardAddress,TranDateTime,ac,invdetails,invtype,paymentmode,CustomerID) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "'," & CDbl(_cmd.Parameter(2)) & ",'" & CStr(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "','" & CStr(_cmd.Parameter(9)) & "','" & CStr(_cmd.Parameter(10)) & "','" & CStr(_cmd.Parameter(11)) & "','" & CDate(_cmd.Parameter(12)) & "','" & CStr(_cmd.Parameter(13)) & "','" & CStr(_cmd.Parameter(14)) & "','" & CStr(_cmd.Parameter(15)) & "','" & CStr(_cmd.Parameter(16)) & "','" & CStr(_cmd.Parameter(17)) & "')"
        Return _SqlStr
    End Function

    Private Shared Function CNCreditCardInfo(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into TblCreditCardInfo(inv_no,inv_date,Debit,TranDateTime,ac,invdetails,invtype,paymentmode,CustomerID) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "'," & CDbl(_cmd.Parameter(2)) & ",'" & CDate(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "')"
        Return _SqlStr
    End Function

    Private Shared Function CNCreditCardInfoFORSC(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into TblCreditCardInfo(inv_no,inv_date,Credit,TranDateTime,ac,invdetails,invtype,paymentmode,CustomerID) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "'," & CDbl(_cmd.Parameter(2)) & ",'" & CDate(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "')"
        Return _SqlStr
    End Function

    Private Shared Function GETREFUNDEDINVNO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "SELECT GroupNo From TblTransaction Where ParentId=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function INSERTSYNREFUNDINV(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "insert into TblCreditCardInfo(inv_no,inv_date,Debit,TranDateTime,ac,invdetails,invtype,paymentmode,CustomerID) Values (" & CInt(_cmd.Parameter(0)) & ",'" & CStr(_cmd.Parameter(1)) & "'," & CDbl(_cmd.Parameter(2)) & ",'" & CDate(_cmd.Parameter(3)) & "','" & CStr(_cmd.Parameter(4)) & "','" & CStr(_cmd.Parameter(5)) & "','" & CStr(_cmd.Parameter(6)) & "','" & CStr(_cmd.Parameter(7)) & "','" & CStr(_cmd.Parameter(8)) & "')"
        Return _SqlStr '"select tblcustomer.customerid,tblcustomer.customername,tblcreditcardinfo.inv_no,tblcreditcardinfo.inv_date,tblcreditcardinfo.invdetails,tblcreditcardinfo.cardtype,tblcreditcardinfo.cardholdername,tblcreditcardinfo.cardnumber,tblcreditcardinfo.ExpiryDate,tblcreditcardinfo.debit from tblcreditcardinfo left join tblcustomer on tblcreditcardinfo.customerid=tblcustomer.customerid"
    End Function
    Private Shared Function GETCCINFO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "select tblcustomer.customerid,tblcustomer.customername,tblcreditcardinfo.inv_no,tblcreditcardinfo.inv_date,tblcreditcardinfo.invdetails,tblcreditcardinfo.cardtype,tblcreditcardinfo.cardholdername,tblcreditcardinfo.cardnumber,tblcreditcardinfo.ExpiryDate,tblcreditcardinfo.debit from tblcreditcardinfo left join tblcustomer on tblcreditcardinfo.customerid=tblcustomer.customerid where tblcreditcardinfo.customerid='" & CStr(_cmd.Parameter(0)) & "'  and tblcreditcardinfo.paymentmode='CC'"
        Return _SqlStr
    End Function
    Private Shared Function GETCCNO(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "select cardnumber from tblcreditcardinfo where inv_no=" & CInt(_cmd.Parameter(0)) & " and paymentmode='CC'"
        Return _SqlStr
    End Function
    Private Shared Function TblCustomerSynFlag(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblCustomer SET Syn='Y', TranDateTime=#" & Date.Now & "# where customerID='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function
    Private Shared Function TblVendorSynFlag(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE Tblvendor SET Syn='Y', TranDateTime=#" & Date.Now & "# where VendorID='" & CStr(_cmd.Parameter(0)) & "'"
        Return _SqlStr
    End Function

    Private Shared Function TblTransactionSynFlag(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblTransaction SET Syn='Y', TranDateTime=#" & Date.Now & "# where ParentID=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function TblLedgerSynFlag(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblLedger SET Syn='Y', TranDateTime=#" & Date.Now & "# where ParentID=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function

    Private Shared Function TblOutStandingSynFlag(ByVal _cmd As InfiniCommand) As String
        _SqlStr = "UPDATE TblOutstanding SET Syn='Y', TranDateTime=#" & Date.Now & "# where ParentID=" & CInt(_cmd.Parameter(0)) & ""
        Return _SqlStr
    End Function
End Class
