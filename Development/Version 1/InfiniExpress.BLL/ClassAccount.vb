Option Strict On
Imports System
Imports InfiniExpress.DAL

Public Class ClassAccount

    Dim _Rd As IDataReader
    Dim _SqlStr As String

    Sub New()
    End Sub

    Public Function CompanyInfo() As IDataReader

        Dim cmd As New InfiniCommand("GETCOMPANYINFO")
        _Rd = ClassGlobal.GetDataReader(cmd)
        Return _Rd

    End Function

    Public Sub UpdateCompanyInfo(ByVal Name As String, ByVal Street As String, ByVal City As String, _
                                  ByVal PCode As String, ByVal Country As String, ByVal Telephone As String, _
                                  ByVal Fax As String, ByVal EMail As String, ByVal VatNo As String, _
                                  ByVal Logo As Object, ByVal LogoName As String, ByVal Checked As Boolean)

        Dim cmd As New InfiniCommand("UPDATECOMPANYINFO")
        With cmd
            .AddParameter("@Name", Name)
            .AddParameter("@Street", Street)
            .AddParameter("@City", City)
            .AddParameter("@PCode", PCode)
            .AddParameter("@Country", Country)
            .AddParameter("@Telephone", Telephone)
            .AddParameter("@Fax", Fax)
            .AddParameter("@EMail", EMail)
            .AddParameter("@VatNo", VatNo)
            .AddParameter("@Add", Checked)
            .AddParameter("@LogoName", LogoName)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

        If Checked = True Then
            Dim cmd1 As New InfiniCommand("UPDATECOMPANYLOGO")
            With cmd1
                .AddParameter("@Logo", Logo)
            End With
            Id = ClassGlobal.ExecuteNonQuery(cmd1)
        End If

    End Sub

    Public Sub RemoveCompanyLogo()

        'Remove Company Logo
        Dim Id As Integer
        Dim cmd2 As New InfiniCommand("REMOVECOMPANYLOGO")
        Id = ClassGlobal.ExecuteNonQuery(cmd2)

    End Sub

    Public Function BankInfo(ByVal id As String) As IDataReader

        Dim cmd As New InfiniCommand("BANKINFO")
        cmd.AddParameter("@LedgerId", id)
        _Rd = ClassGlobal.GetDataReader(cmd)
        Return _Rd

    End Function

    Public Sub BankUpdate(ByVal BankId As String, ByVal Bname As String, ByVal Bstreet As String, ByVal Btown As String, _
                                 ByVal Bcountry As String, ByVal Bpcode As String, ByVal BAname As String, ByVal BActNum As String, ByVal BSCode As String)

        'Update Bank details
        Dim cmd As New InfiniCommand("UPDATEBANK")
        With cmd
            .AddParameter("@BankId", BankId)
            .AddParameter("@BankName", Bname)
            .AddParameter("@BankStreet", Bstreet)
            .AddParameter("@BankTown", Btown)
            .AddParameter("@BankCountry", Bcountry)
            .AddParameter("@BankPCode", Bpcode)
            .AddParameter("@BankAccountName", BAname)
            .AddParameter("@BankAccountNumber", BActNum)
            .AddParameter("@BankSCode", BSCode)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

        'Update Ledger Id and Ledger Name
        Dim cmd2 As New InfiniCommand("UPDATELEDGERID")
        With cmd2
            .AddParameter("@BankId", BankId)
            .AddParameter("@BankName", Bname)
        End With
        Id = ClassGlobal.ExecuteNonQuery(cmd2)

    End Sub

    Public Function LedgerSummary() As DataSet

        Dim _dts As DataSet
        Dim cmd As New InfiniCommand("LEDGERSUMMARY")
        _dts = ClassGlobal.GetDataSet(cmd)
        Return _dts

    End Function

    Public Sub UpdateTaxCodeSummary(ByVal TCode As String, ByVal TRate As Double)

        Dim cmd As New InfiniCommand("UPDATETAXCODE")
        With cmd
            .AddParameter("@TCode", TCode)
            .AddParameter("@TRate", TRate)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

    End Sub

    Public Function FinancialYear() As IDataReader

        Dim cmd As New InfiniCommand("GETFINANCIALYEAR")
        _Rd = ClassGlobal.GetDataReader(cmd)
        Return _Rd

    End Function

    Public Function CheckFinancialYearTransaction() As Boolean

        Dim cmd As New InfiniCommand("CHECKFINANCIALYEARTRANSACTION")
        _Rd = ClassGlobal.GetDataReader(cmd)
        If _Rd.Read Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function FetchMonthYear() As IDataReader

        Dim cmd As New InfiniCommand("GETMAXCUSTOMERID")
        cmd.AddParameter("@MonthNumber", 1)
        _Rd = ClassGlobal.GetDataReader(cmd)
        Return _Rd

    End Function

    Public Sub ReconcileVAT()

        Dim cmd As New InfiniCommand("RECONCILEVAT")
        ClassGlobal.ExecuteNonQuery(cmd)

    End Sub

    Public Sub AddFinancialYear(ByVal StartMonth As String, ByVal CurrentYear As String, ByVal NextYear As String)

        Select Case StartMonth
            Case "January"
                AddMonth("January", CurrentYear, 1)
                AddMonth("February", CurrentYear, 2)
                AddMonth("March", CurrentYear, 3)
                AddMonth("April", CurrentYear, 4)
                AddMonth("May", CurrentYear, 5)
                AddMonth("June", CurrentYear, 6)
                AddMonth("July", CurrentYear, 7)
                AddMonth("August", CurrentYear, 8)
                AddMonth("September", CurrentYear, 9)
                AddMonth("October", CurrentYear, 10)
                AddMonth("November", CurrentYear, 11)
                AddMonth("December", CurrentYear, 12)
            Case "February"
                AddMonth("February", CurrentYear, 1)
                AddMonth("March", CurrentYear, 2)
                AddMonth("April", CurrentYear, 3)
                AddMonth("May", CurrentYear, 4)
                AddMonth("June", CurrentYear, 5)
                AddMonth("July", CurrentYear, 6)
                AddMonth("August", CurrentYear, 7)
                AddMonth("September", CurrentYear, 8)
                AddMonth("October", CurrentYear, 9)
                AddMonth("November", CurrentYear, 10)
                AddMonth("December", CurrentYear, 11)
                AddMonth("January", NextYear, 12)
            Case "March"
                AddMonth("March", CurrentYear, 1)
                AddMonth("April", CurrentYear, 2)
                AddMonth("May", CurrentYear, 3)
                AddMonth("June", CurrentYear, 4)
                AddMonth("July", CurrentYear, 5)
                AddMonth("August", CurrentYear, 6)
                AddMonth("September", CurrentYear, 7)
                AddMonth("October", CurrentYear, 8)
                AddMonth("November", CurrentYear, 9)
                AddMonth("December", CurrentYear, 10)
                AddMonth("January", NextYear, 11)
                AddMonth("February", NextYear, 12)
            Case "April"
                AddMonth("April", CurrentYear, 1)
                AddMonth("May", CurrentYear, 2)
                AddMonth("June", CurrentYear, 3)
                AddMonth("July", CurrentYear, 4)
                AddMonth("August", CurrentYear, 5)
                AddMonth("September", CurrentYear, 6)
                AddMonth("October", CurrentYear, 7)
                AddMonth("November", CurrentYear, 8)
                AddMonth("December", CurrentYear, 9)
                AddMonth("January", NextYear, 10)
                AddMonth("February", NextYear, 11)
                AddMonth("March", NextYear, 12)
            Case "May"
                AddMonth("May", CurrentYear, 1)
                AddMonth("June", CurrentYear, 2)
                AddMonth("July", CurrentYear, 3)
                AddMonth("August", CurrentYear, 4)
                AddMonth("September", CurrentYear, 5)
                AddMonth("October", CurrentYear, 6)
                AddMonth("November", CurrentYear, 7)
                AddMonth("December", CurrentYear, 8)
                AddMonth("January", NextYear, 9)
                AddMonth("February", NextYear, 10)
                AddMonth("March", NextYear, 11)
                AddMonth("April", NextYear, 12)
            Case "June"
                AddMonth("June", CurrentYear, 1)
                AddMonth("July", CurrentYear, 2)
                AddMonth("August", CurrentYear, 3)
                AddMonth("September", CurrentYear, 4)
                AddMonth("October", CurrentYear, 5)
                AddMonth("November", CurrentYear, 6)
                AddMonth("December", CurrentYear, 7)
                AddMonth("January", NextYear, 8)
                AddMonth("February", NextYear, 9)
                AddMonth("March", NextYear, 10)
                AddMonth("April", NextYear, 11)
                AddMonth("May", NextYear, 12)
            Case "July"
                AddMonth("July", CurrentYear, 1)
                AddMonth("August", CurrentYear, 2)
                AddMonth("September", CurrentYear, 3)
                AddMonth("October", CurrentYear, 4)
                AddMonth("November", CurrentYear, 5)
                AddMonth("December", CurrentYear, 6)
                AddMonth("January", NextYear, 7)
                AddMonth("February", NextYear, 8)
                AddMonth("March", NextYear, 9)
                AddMonth("April", NextYear, 10)
                AddMonth("May", NextYear, 11)
                AddMonth("June", NextYear, 12)
            Case "August"
                AddMonth("August", CurrentYear, 1)
                AddMonth("September", CurrentYear, 2)
                AddMonth("October", CurrentYear, 3)
                AddMonth("November", CurrentYear, 4)
                AddMonth("December", CurrentYear, 5)
                AddMonth("January", NextYear, 6)
                AddMonth("February", NextYear, 7)
                AddMonth("March", NextYear, 8)
                AddMonth("April", NextYear, 9)
                AddMonth("May", NextYear, 10)
                AddMonth("June", NextYear, 11)
                AddMonth("July", NextYear, 12)
            Case "September"
                AddMonth("September", CurrentYear, 1)
                AddMonth("October", CurrentYear, 2)
                AddMonth("November", CurrentYear, 3)
                AddMonth("December", CurrentYear, 4)
                AddMonth("January", NextYear, 5)
                AddMonth("February", NextYear, 6)
                AddMonth("March", NextYear, 7)
                AddMonth("April", NextYear, 8)
                AddMonth("May", NextYear, 9)
                AddMonth("June", NextYear, 10)
                AddMonth("July", NextYear, 11)
                AddMonth("August", NextYear, 12)
            Case "October"
                AddMonth("October", CurrentYear, 1)
                AddMonth("November", CurrentYear, 2)
                AddMonth("December", CurrentYear, 3)
                AddMonth("January", NextYear, 4)
                AddMonth("February", NextYear, 5)
                AddMonth("March", NextYear, 6)
                AddMonth("April", NextYear, 7)
                AddMonth("May", NextYear, 8)
                AddMonth("June", NextYear, 9)
                AddMonth("July", NextYear, 10)
                AddMonth("August", NextYear, 11)
                AddMonth("September", NextYear, 12)
            Case "November"
                AddMonth("November", CurrentYear, 1)
                AddMonth("December", CurrentYear, 2)
                AddMonth("January", NextYear, 3)
                AddMonth("February", NextYear, 4)
                AddMonth("March", NextYear, 5)
                AddMonth("April", NextYear, 6)
                AddMonth("May", NextYear, 7)
                AddMonth("June", NextYear, 8)
                AddMonth("July", NextYear, 9)
                AddMonth("August", NextYear, 10)
                AddMonth("September", NextYear, 11)
                AddMonth("October", NextYear, 12)
            Case "December"
                AddMonth("December", CurrentYear, 1)
                AddMonth("January", NextYear, 2)
                AddMonth("February", NextYear, 3)
                AddMonth("March", NextYear, 4)
                AddMonth("April", NextYear, 5)
                AddMonth("May", NextYear, 6)
                AddMonth("June", NextYear, 7)
                AddMonth("July", NextYear, 8)
                AddMonth("August", NextYear, 9)
                AddMonth("September", NextYear, 10)
                AddMonth("October", NextYear, 11)
                AddMonth("November", NextYear, 12)
        End Select

    End Sub

    Private Sub AddMonth(ByVal FMonth As String, ByVal FYear As String, ByVal MNo As Integer)

        Dim cmd As New InfiniCommand("ADDMONTH")
        With cmd
            .AddParameter("@MonthNumber", MNo)
            .AddParameter("@FMonth", FMonth)
            .AddParameter("@FYear", FYear)
        End With
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd)

    End Sub

#Region "For VAT"

    Public Function VATBetweenTranscation(ByVal Date1 As String, ByVal Date2 As String, ByVal Criteria As String) As Integer

        Dim _tempvar As Integer
        Dim cmd As New InfiniCommand("VATBETWEENTRANSACTION")
        With cmd
            .AddParameter("@Criteria", Criteria)
        End With
        _Rd = ClassGlobal.GetDataReader(cmd)
        While _Rd.Read()
            _tempvar = CInt(_Rd("Nor"))
        End While
        _Rd.Close()
        Return _tempvar

    End Function

    Public Function VATBeforeTranscation(ByVal Date1 As String, ByVal Criteria As String) As Integer

        Dim _tempvar As Integer
        Dim cmd1 As New InfiniCommand("VATBEFORETRANSACTION")
        With cmd1
            .AddParameter("@Criteria", Criteria)
        End With
        _Rd = ClassGlobal.GetDataReader(cmd1)
        While _Rd.Read()
            _tempvar = CInt(_Rd("Nor"))
        End While
        _Rd.Close()
        Return _tempvar

    End Function

    Public Function VATQUERY1(ByVal _tempQry As String) As Double

        Dim _tempVar As Double
        Dim cmd2 As New InfiniCommand("VATQUERY1")
        cmd2.AddParameter("@SQLStr", _tempQry)
        _Rd = ClassGlobal.GetDataReader(cmd2)
        While _Rd.Read
            _tempVar = _tempVar + CDbl(_Rd(0))
        End While
        _Rd.Close()
        Return _tempVar

    End Function

    Public Function VATQUERY2(ByVal _tempQry As String) As Double

        Dim _tempVar As Double
        Dim cmd3 As New InfiniCommand("VATQUERY2")
        cmd3.AddParameter("@SQLStr", _tempQry)
        _Rd = ClassGlobal.GetDataReader(cmd3)
        While _Rd.Read
            _tempVar = _tempVar + CDbl(_Rd(0))
        End While
        _Rd.Close()
        Return _tempVar

    End Function

    Public Function VATDetail(ByVal _strTemp As String) As IDataReader

        Dim cmd0 As New InfiniCommand("VATDETAILS")
        cmd0.AddParameter("@SQLStr", _strTemp)
        _Rd = ClassGlobal.GetDataReader(cmd0)
        Return _Rd

    End Function

    Public Sub ReconcileTrans(ByVal _QryStr As String)

        Dim _dSet As DataSet
        Dim _tp As String
        Dim _invNo As Integer

        Dim cmd As New InfiniCommand("RECONCILETRANS")
        cmd.AddParameter("@SQLStr", _QryStr)
        _dSet = ClassGlobal.GetDataSet(cmd)

        If Not _dSet.Tables(0).Rows.Count < 1 Then

            Dim TempTbl As DataTable = _dSet.Tables(0)
            Dim TempRow As DataRow = TempTbl.Rows(0)

            For Each TempRow In TempTbl.Rows

                _tp = CStr(TempRow.Item("invtp"))
                _invNo = CInt(TempRow.Item("ParentID"))

                Dim cmd1 As New InfiniCommand("UPDATERECONCILETRANS")
                With cmd1
                    .AddParameter("@InvNo", _invNo)
                    .AddParameter("@InvTP", _tp)
                End With
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(cmd1)

            Next

        End If

    End Sub

#End Region

End Class
