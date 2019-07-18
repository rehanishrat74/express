Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class ReportQueries

    Public Function GetCustID() As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("Select CustomerID from TblCustomer", olecon)
        oleAdapt.Fill(myds)
        Return myds
    End Function

    Public Function GetVendID() As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("Select VendorID from TblVendor", olecon)
        oleAdapt.Fill(myds)
        Return myds
    End Function

    Public Function GetLedgerID() As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("Select LedgerID from TblLedgerIds", olecon)
        oleAdapt.Fill(myds)
        Return myds
    End Function

    Public Function GetCustList() As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT CustomerId,  CustomerName  ,  (Street +' '  +  Town  + ' '  + Country + ' '  + PostalCode )   As Address  ,  ContactName,Email,Telephone , Fax FROM TblCustomer ", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\CustList.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function



    Public Function GetCustActivity(ByVal CustID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT TL.ParentId , TL.TransId , SUM (TL.Debit) As Debit ,SUM (TL.Credit) AS Credit , TL.TRANTYPE ,  TT.TaxId , TT.Invtp, TT.InvDate,InvDetails ,  iif( SUM (TL.Debit) > SUM (TL.Credit), SUM (TL.Debit),SUM (TL.Credit)) as Amount FROM TblLedger as TL  INNER JOIN  TblTransaction as TT  ON TL.ParentID=TT.ParentId WHERE TL.TRANSID=" & "'" & CStr(CustID) & "'  AND TL.TRANTYPE='C' GROUP By TL.ParentId , TL.TransId , TT.TaxId ,TL.TRANTYPE , TT.Invtp , TT.InvDate , InvDetails Order By TL.ParentID", olecon)
        oleAdapt.Fill(myds)

        'Dim Filename As String = Application.StartupPath & "\CustActivity.xml"
        'Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        'Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        'myds.WriteXmlSchema(MyXmlTextWriter)
        'myFileStream.Close()
        'File.Delete(Filename)
        Return myds

    End Function

    Public Function GetSalesInvoice(ByVal CustID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  ,TblTransaction.InvVAT , TblTransaction.InvNet , TblTransaction.InvTP  ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet   As Gross FROM TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND (TblTRANSACTION.InvTP='SI') AND ((TblLEDGER.TransId)=" & "'" & CStr(CustID) & "') AND ((TblLEDGER.trantype)='C') AND (TblTransaction.Mark='N') ) ORDER BY TBLTRANSACTION.ParentId", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\SalesInvoice.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function

    Public Function GetSalesReceipts(ByVal CustID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT DISTINCT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId , TblTransaction.InvDetails  , TblTransaction.InvTP , TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet   As Gross FROM TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND ((TblLEDGER.TransId)=" & "'" & CStr(CustID) & "' ) AND ((TblLEDGER.trantype)='C')) AND ((TblTransaction.Invtp)='SR')", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\SalesReceipts.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function

    Public Function GetVendorList() As DataSet

        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT  VendorID ,VendorName, ContactName  ,Email , ( Street +' ' + Town   +' ' + Country  + ' ' + PostalCode )  As Address ,Telephone FROM TblVendor", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\VendorList.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function


    Public Function GetVendorActivity(ByVal VendID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT TL.ParentId , TL.TransId , SUM (TL.Debit) As Debit ,SUM (TL.Credit) AS Credit , TL.TRANTYPE ,TT.TaxId , TT.Invtp, TT.InvDate, InvDetails ,iif(SUM (TL.Debit)  > SUM (TL.Credit) ,SUM (TL.Debit)  ,SUM (TL.Credit) ) as Amount FROM TblLedger as TL  INNER JOIN  TblTransaction as TT  ON TL.ParentID=TT.ParentId WHERE TL.TRANSID=" & "'" & CStr(VendID) & "' AND TL.TRANTYPE='S' GROUP By TL.ParentId , TL.TransId , TT.TaxId ,TL.TRANTYPE , TT.Invtp , TT.InvDate,  InvDetails Order By TL.ParentID ", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\VendorActivity.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function


    Public Function GetPurchaseInvoice(ByVal VendID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate ,TblTransaction.ParentId ,TblTransaction.InvDetails  ,TblTransaction.InvVAT , TblTransaction.InvNet , TblTransaction.InvTP ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet   As Gross FROM TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId])  AND ((TblLEDGER.TransId)=" & "'" & CStr(VendID) & "' ) AND ((TblLEDGER.trantype)='S') AND ((TblTRANSACTION.InvTP)='PI')  AND (TblTransaction.Mark='N') ) ORDER BY TBLTRANSACTION.ParentId", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\PurchaseInvoice.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds

    End Function

    Public Function GetPurchasePayments(ByVal VendID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT DISTINCT TblLedger.TransId,TblTransaction.TaxId ,TblTransaction.InvDate , TblTransaction.ParentId , TblTransaction.InvDetails  ,TblTransaction.InvVAT , TblTransaction.InvNet , TblTransaction.InvTP ,TblLedger.TranType , TblTransaction.InvVAT+TblTransaction.InvNet As Gross FROM TblLEDGER , TblTRANSACTION  WHERE (((TblLEDGER.ParentId)=[TblTransaction].[ParentId]) AND ((TblLEDGER.TransId)=" & "'" & CStr(VendID) & "' ) AND ((TblLEDGER.trantype)='S')) AND((TblTransaction.Invtp)='PP') ", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\PurchasePayments.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds

    End Function

    Public Function GetLedgerBalance() As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT tblledger.transid, (select name from tblledgerids where ledgerid=tblledger.transid) As tname,  IIF((SUM(tblledger.debit)-SUM(tblledger.credit))>0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TDR, IIF((SUM(tblledger.debit)-SUM(tblledger.credit))<0,ABS(SUM(tblledger.debit)-SUM(tblledger.credit)),0) As TCR FROM TblLedger Where tblledger.trantype='T' OR tblledger.trantype='U' OR tblledger.trantype='B' Group by tblledger.transid ", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\LEDGERBALANCE.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds
    End Function

    Public Function GetLedgerActivity(ByVal LedgerID As String) As DataSet
        Dim myds As New DataSet()
        Dim olecon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InfiniAccountsExpress\\InfiniExpress\\DataBase\\ExpressDB.mdb")
        Dim oleAdapt As New OleDbDataAdapter("SELECT TL.ParentId , TL.TransId , SUM (TL.Debit) As Debit ,SUM (TL.Credit) AS Credit , TL.TRANTYPE ,  TT.TaxId , TT.Invtp, TT.InvDate, InvDetails, iif(SUM (TL.Debit)  > SUM (TL.Credit) ,SUM (TL.Debit)  ,SUM (TL.Credit) ) as Amount FROM TblLedger as TL  INNER JOIN  TblTransaction as TT  ON TL.ParentID=TT.ParentId WHERE ( TL.TRANSID=" & "'" & CStr(LedgerID) & "' AND TL.TRANTYPE IN ('S' ,'T','U' , 'B') ) GROUP By TL.ParentId , TL.TransId , TT.TaxId ,TL.TRANTYPE , TT.Invtp , TT.InvDate , InvDetails Order By TL.ParentID ", olecon)
        oleAdapt.Fill(myds)

        Dim Filename As String = Application.StartupPath & "\LedgerActivity.xml"
        Dim myFileStream As New System.IO.FileStream(Filename, System.IO.FileMode.Create)
        Dim MyXmlTextWriter As New System.Xml.XmlTextWriter(myFileStream, System.Text.Encoding.Unicode) ' from saira
        myds.WriteXmlSchema(MyXmlTextWriter)
        myFileStream.Close()
        File.Delete(Filename)
        Return myds

    End Function
End Class
