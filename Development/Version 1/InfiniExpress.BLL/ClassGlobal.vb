'Option Strict On
Imports System
Imports System.Data
Imports InfiniExpress.DAL

Public Class ClassGlobal

    Dim Dr As IDataReader
    Dim _Dts As DataSet
    Public Shared cmdCollection As New Collection()

#Region "Public Shared Methods"

    Public Overloads Shared Function GetDataSet(ByRef Cmmd As InfiniCommand) As DataSet

        Dim dset As DataSet
        Try
            dset = AccessHelper.ExecuteDataset(Cmmd)
        Catch sqle As DataException
            dset = Nothing
            Throw New Exception("GetDataSet Method :" & sqle.Message)
        Catch e As Exception
            dset = Nothing
            Throw New Exception("GetDataSet Method :" & e.Message)
        End Try
        Return dset

    End Function

    Public Overloads Shared Function GetDataReader(ByRef Cmmd As InfiniCommand) As IDataReader

        Dim Reader As IDataReader
        Try
            Reader = AccessHelper.ExecuteReader(Cmmd)
        Catch sqle As DataException
            Reader = Nothing
            Throw New Exception("GetDataReader Method :" & sqle.Message)
        Catch e As Exception
            Reader = Nothing
            Throw New Exception("GetDataReader Method :" & e.Message)
        End Try
        Return Reader

    End Function

    Public Overloads Shared Function GetDataReader_Pro(ByRef Cmmd As InfiniCommand) As IDataReader

        Dim Reader As IDataReader
        Try
            Reader = AccessHelper.ExecuteReader_Pro(Cmmd)
        Catch sqle As DataException
            Reader = Nothing
            Throw New Exception("GetDataReader Method :" & sqle.Message)
        Catch e As Exception
            Reader = Nothing
            Throw New Exception("GetDataReader Method :" & e.Message)
        End Try
        Return Reader

    End Function

    Public Shared Function ExecuteNonQuery(ByVal Cmmd As InfiniCommand) As Integer

        cmdCollection.Add(Cmmd)
        Return 1

    End Function

    Public Overloads Shared Function ExecuteCommand() As Boolean
        Return AccessHelper.ExecuteCollection(cmdCollection)
        cmdCollection = New Collection
    End Function

    Public Overloads Shared Function ExecuteCommandFor_Pro() As Boolean
        Return AccessHelper.ExecuteCollectionFor_Pro(cmdCollection)
        cmdCollection = New Collection
    End Function

#End Region

#Region "Global Methods"

    Public Function GetTransactionNo() As Integer

        Dim Tcmd As New InfiniCommand("GETTRANSACTIONNO")
        Dim TrnNo As Integer = 1
        Dr = ClassGlobal.GetDataReader(Tcmd)
        If Dr.Read Then
            If Not IsDBNull(Dr("Parentno")) Then
                TrnNo = CInt(Dr("parentno"))
            End If
        End If
        Dr.Close()
        Return TrnNo

    End Function

    Public Function GetLedgerAutoNo() As Integer

        Dim Lcmd As New InfiniCommand("GETLEDGERAUTONO")
        Dim LedNo As Integer = 1
        Dr = ClassGlobal.GetDataReader(Lcmd)
        If Dr.Read Then
            If Not IsDBNull(Dr("AutoNo")) Then
                LedNo = CInt(Dr("AutoNo"))
            End If
        End If
        Dr.Close()
        Return LedNo

    End Function

    Public Function GetOutStandingMaxOSNo() As Integer

        Dim OSMaxNo As Integer = 1
        Dim Oscmd As New InfiniCommand("GETMAXOSNO")
        _Dts = ClassGlobal.GetDataSet(Oscmd)
        If Not _Dts.Tables(0).Rows.Count = 0 Then
            OSMaxNo = _Dts.Tables(0).Rows(0).Item("OSMaxNo")
        End If
        _Dts.Dispose()

        Return OSMaxNo

    End Function

    Public Sub InsertTransactionData(ByVal ParentId As Double, ByVal InvTp As String, ByVal TaxId As String, ByVal InvDate As String, _
               ByVal InvDetails As String, ByVal InvNet As Double, ByVal InvVat As Double, ByVal TranDate As Date, ByVal TranDateTime As String, _
               Optional ByVal VAT As String = "-", Optional ByVal Refund As String = "N", Optional ByVal Mark As String = "N")

        Dim cmd8 As New InfiniCommand("INSERTTBLTRANSACTION")
        With cmd8
            .AddParameter("@ParentId", ParentId)
            .AddParameter("@InvTp", InvTp)
            .AddParameter("@TaxId", TaxId)
            .AddParameter("@InvDate", InvDate)
            .AddParameter("@InvDetails", InvDetails)
            .AddParameter("@InvNet", InvNet)
            .AddParameter("@InvVat", InvVat)
            .AddParameter("@Refund", Refund)
            .AddParameter("@Mark", Mark)
            .AddParameter("@VAT", VAT)
            .AddParameter("@TranDate", TranDate)
            .AddParameter("@TranDateTime", TranDateTime)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd8)

    End Sub

    Public Sub InsertLedger(ByVal ParentId As Double, ByVal ChildId As Double, ByVal TransId As String, ByVal Debit As Double, ByVal Credit As Double, _
               ByVal TranType As String, ByVal TranRef As String, ByVal PayStatus As String, ByVal AutoNo As Integer, ByVal TranDateTime As String)

        Dim cmd7 As New InfiniCommand("INSERTTBLLEDGER")
        With cmd7
            .AddParameter("@ParentId", ParentId)
            .AddParameter("@ChildId", ChildId)
            .AddParameter("@TransId", TransId)
            .AddParameter("@Debit", Debit)
            .AddParameter("@Credit", Credit)
            .AddParameter("@TranType", TranType)
            .AddParameter("@TranRef", TranRef)
            .AddParameter("@PayStatus", PayStatus)
            .AddParameter("@AutoNo", AutoNo)
            .AddParameter("@TranDateTime", TranDateTime)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd7)

    End Sub

    Public Sub InsertOutStanding(ByVal AutoNo As Integer, ByVal ParentId As Double, ByVal ChildId As Integer, ByVal TransId As String, ByVal Debit As Double, ByVal Credit As Double, _
                                 ByVal TranType As String, ByVal Details As String, ByVal TranRef As String, ByVal OSDate As String, ByVal OSFrom As Double, ByVal OSTo As Double, _
                                 ByVal OSFT As String, ByVal TranDateTime As String, ByVal OSMaxNo As Integer)

        Dim cmd6 As New InfiniCommand("INSERTTBLOUTSTANDING")
        With cmd6
            .AddParameter("@AutoNo", AutoNo)
            .AddParameter("@ParentId", ParentId)
            .AddParameter("@ChildId", ChildId)
            .AddParameter("@TransId", TransId)
            .AddParameter("@Debit", Debit)
            .AddParameter("@Credit", Credit)
            .AddParameter("@TranType", TranType)
            .AddParameter("@Details", Details)
            .AddParameter("@TranRef", TranRef)
            .AddParameter("@OSDate", OSDate)
            .AddParameter("@OSFrom", OSFrom)
            .AddParameter("@OSTo", OSTo)
            .AddParameter("@OSFT", OSFT)
            .AddParameter("@TranDateTime", TranDateTime)
            .AddParameter("@OSNo", OSMaxNo)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd6)

    End Sub

    Public Sub UpdateLedgerPayStatus(ByVal _tempAuto As Integer, ByVal _flag As String, ByVal TranDateTime As String)

        Dim cmd3 As New InfiniCommand("UPDATEPAYSTATUS")
        With cmd3
            .AddParameter("@MyAutoNo", _tempAuto)
            .AddParameter("@PayStatus", _flag)
            .AddParameter("@TranDateTime", TranDateTime)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd3)

    End Sub

    Public Function TaxCodeSummary() As DataSet

        Dim cmd As New InfiniCommand("LOADTAXIDS")
        _Dts = ClassGlobal.GetDataSet(cmd)
        Return _Dts

    End Function

    Public Function GetExistingLedgerAutoNo(ByVal ChildNo As Integer, ByVal TType As String) As Integer

        Dim LedgerAuto As Integer = 1
        Dim cmd0 As New InfiniCommand("GETEXISTLEDGERAUTONO")
        With cmd0
            .AddParameter("@ChildNo", ChildNo)
            .AddParameter("@TType", TType)
        End With
        Dr = ClassGlobal.GetDataReader(cmd0)
        If Dr.Read Then
            If Not IsDBNull(Dr("MyAutoNo")) Then
                LedgerAuto = CInt(Dr("MyAutoNo"))
            End If
        End If
        Dr.Close()
        Return LedgerAuto

    End Function

    Sub UpdateOutstanding(ByVal _MyAuto As Integer, ByVal TDateTime As String)

        Dim cmd2 As New InfiniCommand("UPDATEOUTSTANDING")
        With cmd2
            .AddParameter("@Tauto", _MyAuto)
            .AddParameter("@TranDateTime", TDateTime)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd2)

    End Sub

    Public Sub ChkRefund(ByVal ChkMark As String, ByVal Chk As String, ByVal Id As Integer, ByVal TDateTime As String)

        Dim cmd As New InfiniCommand("CHECKREFUND")
        With cmd
            .AddParameter("@ChkMark", ChkMark)
            .AddParameter("@Chk", Chk)
            .AddParameter("@ParentId", Id)
            .AddParameter("@TranDateTime", TDateTime)
        End With
        Dim tId As Integer = ClassGlobal.ExecuteNonQuery(cmd)

    End Sub

    Public Sub PaystatusFlag(ByVal Tparentid As Integer, ByVal TStr As String, ByVal TDateTime As String)

        Dim Tmyautono As Integer, Paystatus As String
        Dim cmd4 As New InfiniCommand("PAYSTATUSFLAG")
        With cmd4
            .AddParameter("@ParentId", Tparentid)
            .AddParameter("@TType", TStr)
        End With
        Dr = ClassGlobal.GetDataReader(cmd4)

        If Dr.Read = True Then
            Tmyautono = CInt(Dr("myautono"))
            If CDbl(Dr("osvalue")) = 0 Then
                Paystatus = "*"
            ElseIf CDbl(Dr("ledgerval")) = CDbl(Dr("osvalue")) Then
                Paystatus = "f"
            Else
                Paystatus = "p"
            End If
        End If
        Dr.Close()

        If Tmyautono > 0 Then
            UpdateLedgerPayStatus(Tmyautono, Paystatus, TDateTime)
        End If

    End Sub

    Public Function GetBankBalance(ByVal _tempBID As String) As Double

        Dim BValue As Double

        Dim cmd5 As New InfiniCommand("GETBANKBALANCE")
        With cmd5
            .AddParameter("@TransId", _tempBID)
        End With
        Dr = ClassGlobal.GetDataReader(cmd5)
        While Dr.Read
            BValue = CDbl(Dr("tbal"))
        End While
        Dr.Close()

        Return BValue

    End Function

#End Region

End Class
