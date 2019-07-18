Option Strict On
Imports System
Imports System.Data
Imports InfiniExpress.DAL
Imports Microsoft.VisualBasic

Public Class ClassUtility
    Dim strPath As String
    Dim _cmd As InfiniCommand
    Dim _dts As DataSet
    Dim ClsGlobal As New ClassGlobal()
    Dim _ds As New DataSet()
    Public Shared cmdCollection As New Collection()
#Region "Import/Export Section"
    Public Sub Dbpath_Pro(ByVal Str As String)
        AccessHelper.DbPath = Str
    End Sub

    Public Sub DeleteAllTable_Pro()
        Dim _cmd As New InfiniCommand("DELETECUSTOMER_PRO")
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)
        Dim _cmd1 As New InfiniCommand("DELETESUPPLIER_PRO")
        Dim Id1 As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
        'Dim _cmd3 As New InfiniCommand("DELETEBANK")
        'Dim Id3 As Integer = ClassGlobal.ExecuteNonQuery(_cmd3)
        Dim _cmd6 As New InfiniCommand("DELETETRANSACTION")
        Dim Id6 As Integer = ClassGlobal.ExecuteNonQuery(_cmd6)
        Dim _cmd7 As New InfiniCommand("DELETELEDGER")
        Dim Id7 As Integer = ClassGlobal.ExecuteNonQuery(_cmd7)
        Dim _cmd8 As New InfiniCommand("DELETEOUTSTANDING")
        Dim Id8 As Integer = ClassGlobal.ExecuteNonQuery(_cmd8)
        Dim _cmd9 As New InfiniCommand("DELETEINVOICE")
        Dim Id9 As Integer = ClassGlobal.ExecuteNonQuery(_cmd9)
        Dim _cmd10 As New InfiniCommand("DELETEINVOICE_DET")
        Dim Id10 As Integer = ClassGlobal.ExecuteNonQuery(_cmd10)
        Dim _cmd11 As New InfiniCommand("DELETESALESOD")
        Dim Id11 As Integer = ClassGlobal.ExecuteNonQuery(_cmd11)
        Dim _cmd12 As New InfiniCommand("DELETESALESORDER")
        Dim Id12 As Integer = ClassGlobal.ExecuteNonQuery(_cmd12)
        Dim _cmd13 As New InfiniCommand("DELETEPURACHASEOD")
        Dim Id13 As Integer = ClassGlobal.ExecuteNonQuery(_cmd13)
        Dim _cmd14 As New InfiniCommand("DELETEPURCHASEORDER")
        Dim Id14 As Integer = ClassGlobal.ExecuteNonQuery(_cmd14)
        Dim _cmd15 As New InfiniCommand("DELETESUBSCRIPTION")
        Dim Id15 As Integer = ClassGlobal.ExecuteNonQuery(_cmd15)
        Dim _cmd16 As New InfiniCommand("DELETESUBSCRIPTION_DET")
        Dim Id16 As Integer = ClassGlobal.ExecuteNonQuery(_cmd16)
        Dim _cmd17 As New InfiniCommand("DELETERECURRING")
        Dim Id17 As Integer = ClassGlobal.ExecuteNonQuery(_cmd17)
        Dim _cmd18 As New InfiniCommand("DELETEPRODUCT")
        Dim Id18 As Integer = ClassGlobal.ExecuteNonQuery(_cmd18)
        Dim _cmd20 As New InfiniCommand("DELETEPRODUCT_BOM")
        Dim Id20 As Integer = ClassGlobal.ExecuteNonQuery(_cmd20)
        Dim _cmd21 As New InfiniCommand("DELETEPROD_SALES")
        Dim Id21 As Integer = ClassGlobal.ExecuteNonQuery(_cmd21)
        Dim _cmd22 As New InfiniCommand("DELETEPROD_LINK")
        Dim Id22 As Integer = ClassGlobal.ExecuteNonQuery(_cmd22)
        Dim _cmd23 As New InfiniCommand("DELETEPRODUCT_DISCOUNT")
        Dim Id23 As Integer = ClassGlobal.ExecuteNonQuery(_cmd23)
        Dim _cmd24 As New InfiniCommand("DELETEPROD_ACTIVITY")
        Dim Id24 As Integer = ClassGlobal.ExecuteNonQuery(_cmd24)
        Dim _cmd25 As New InfiniCommand("DELETEPREPAYMENT")
        Dim Id25 As Integer = ClassGlobal.ExecuteNonQuery(_cmd25)
        Dim _cmd26 As New InfiniCommand("DELETECART_ITEMS")
        Dim Id26 As Integer = ClassGlobal.ExecuteNonQuery(_cmd26)
        Dim _cmd27 As New InfiniCommand("DELETECART_ORDER")
        Dim Id27 As Integer = ClassGlobal.ExecuteNonQuery(_cmd27)
        Dim _cmd28 As New InfiniCommand("DELETEACCURAL")
        Dim Id28 As Integer = ClassGlobal.ExecuteNonQuery(_cmd28)
        Dim _cmd29 As New InfiniCommand("DELETECUST_PROD_PRICE")
        Dim Id29 As Integer = ClassGlobal.ExecuteNonQuery(_cmd29)
        Dim _cmd30 As New InfiniCommand("DELETEQUESTION")
        Dim Id30 As Integer = ClassGlobal.ExecuteNonQuery(_cmd30)
        Dim _cmd31 As New InfiniCommand("DELETEQANSWER")
        Dim Id31 As Integer = ClassGlobal.ExecuteNonQuery(_cmd31)
        Dim _cmd32 As New InfiniCommand("DELETEQTABLE")
        Dim Id32 As Integer = ClassGlobal.ExecuteNonQuery(_cmd32)
        Dim _cmd33 As New InfiniCommand("DELETECATEGORY")
        Dim Id33 As Integer = ClassGlobal.ExecuteNonQuery(_cmd33)
        Dim _cmd34 As New InfiniCommand("DELETENRECORD")
        Dim Id34 As Integer = ClassGlobal.ExecuteNonQuery(_cmd34)


        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub DeleteAllTable_Exp()
        Dim _cmd As New InfiniCommand("EDELETECUSTOMER")
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)
        Dim _cmd1 As New InfiniCommand("EDELETEVENDOR")
        Dim Id1 As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
        Dim _cmd6 As New InfiniCommand("EDELETETRANSACTION")
        Dim Id6 As Integer = ClassGlobal.ExecuteNonQuery(_cmd6)
        Dim _cmd7 As New InfiniCommand("EDELETELEDGER")
        Dim Id7 As Integer = ClassGlobal.ExecuteNonQuery(_cmd7)
        Dim _cmd8 As New InfiniCommand("EDELETEOUTSTANDING")
        Dim Id8 As Integer = ClassGlobal.ExecuteNonQuery(_cmd8)
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub AllCustomer()
        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTALLCUSTOMER")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("INSERTALLCUSTOMER")
            With _cmd1

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("ac", CStr(_dts.Tables(0).Rows(_i).Item(1)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(2)) = False Then .AddParameter("name", CStr(_dts.Tables(0).Rows(_i).Item(2)))

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(3)) = True Then
                    .AddParameter("street", "")
                Else
                    .AddParameter("street", CStr(_dts.Tables(0).Rows(_i).Item(3)))
                End If
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(4)) = True Then
                    .AddParameter("town", "")
                Else
                    .AddParameter("town", CStr(_dts.Tables(0).Rows(_i).Item(4)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(5)) = True Then
                    .AddParameter("country", "")
                Else
                    .AddParameter("country", CStr(_dts.Tables(0).Rows(_i).Item(5)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(6)) = True Then
                    .AddParameter("postcode", "")
                Else
                    .AddParameter("postcode", CStr(_dts.Tables(0).Rows(_i).Item(6)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(7)) = True Then
                    .AddParameter("cont_name", "")
                Else
                    .AddParameter("cont_name", CStr(_dts.Tables(0).Rows(_i).Item(7)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(8)) = False Then .AddParameter("vemail", CStr(_dts.Tables(0).Rows(_i).Item(8)))

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(9)) = True Then
                    .AddParameter("telephone", "")
                Else
                    .AddParameter("telephone", CStr(_dts.Tables(0).Rows(_i).Item(9)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(10)) = True Then
                    .AddParameter("fax", "")
                Else
                    .AddParameter("fax", CStr(_dts.Tables(0).Rows(_i).Item(10)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(11)) = True Then
                    .AddParameter("vat", "")
                Else
                    .AddParameter("vat", CStr(_dts.Tables(0).Rows(_i).Item(11)))
                End If
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(12)) = False Then .AddParameter("credit_limit", CDbl(_dts.Tables(0).Rows(_i).Item(12)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(13)) = False Then .AddParameter("std_tax_code", CStr(_dts.Tables(0).Rows(_i).Item(13)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(14)) = False Then .AddParameter("nc", CStr(_dts.Tables(0).Rows(_i).Item(14)))

                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next

        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub EAllCustomer(ByVal _strPath As String)

        Dim _rCount As Integer, _i As Integer
        Dim _ds As New DataSet
        Dim _myRow As DataRow

        'Get Dataset From XML File
        _ds.ReadXml(_strPath)

        For Each _myRow In _ds.Tables(0).Rows
            Dim _cmd1 As New InfiniCommand("EINSERTALLCUSTOMER")
            With _cmd1

                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("CustomerID", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("CustomerName", CStr(_myRow.Item(2)))

                If IsDBNull(_myRow.Item(3)) = True Then
                    .AddParameter("street", "")
                Else
                    .AddParameter("street", CStr(_myRow.Item(3)))
                End If

                If IsDBNull(_myRow.Item(4)) = True Then
                    .AddParameter("town", "")
                Else
                    .AddParameter("town", CStr(_myRow.Item(4)))
                End If

                If IsDBNull(_myRow.Item(5)) = True Then
                    .AddParameter("country", "")
                Else
                    .AddParameter("country", CStr(_myRow.Item(5)))
                End If

                If IsDBNull(_myRow.Item(6)) = True Then
                    .AddParameter("postalcode", "")
                Else
                    .AddParameter("postalcode", CStr(_myRow.Item(6)))
                End If

                If IsDBNull(_myRow.Item(7)) = True Then
                    .AddParameter("contactname", "")
                Else
                    .AddParameter("contactname", CStr(_myRow.Item(7)))
                End If

                If IsDBNull(_myRow.Item(8)) = False Then .AddParameter("vemail", CStr(_myRow.Item(8)))

                If IsDBNull(_myRow.Item(9)) = True Then
                    .AddParameter("telephone", "")
                Else
                    .AddParameter("telephone", CStr(_myRow.Item(9)))
                End If

                If IsDBNull(_myRow.Item(10)) = True Then
                    .AddParameter("fax", "")
                Else
                    .AddParameter("fax", CStr(_myRow.Item(10)))
                End If

                If IsDBNull(_myRow.Item(11)) = True Then
                    .AddParameter("vat", "")
                Else
                    .AddParameter("vat", CStr(_myRow.Item(11)))
                End If
                If IsDBNull(_myRow.Item(12)) = False Then .AddParameter("credit_limit", CDbl(_myRow.Item(12)))
                .AddParameter("Syn", CStr(_myRow.Item(16)))
                .AddParameter("trandatetime", CStr(_myRow.Item(17)))

                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub AllVendor()

        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTALLVENDOR")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("INSERTALLVENDOR")
            With _cmd1

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("ac", CStr(_dts.Tables(0).Rows(_i).Item(1)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(2)) = False Then .AddParameter("name", CStr(_dts.Tables(0).Rows(_i).Item(2)))

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(3)) = True Then
                    .AddParameter("street", "")
                Else
                    .AddParameter("street", CStr(_dts.Tables(0).Rows(_i).Item(3)))
                End If
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(4)) = True Then
                    .AddParameter("town", "")
                Else
                    .AddParameter("town", CStr(_dts.Tables(0).Rows(_i).Item(4)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(5)) = True Then
                    .AddParameter("country", "")
                Else
                    .AddParameter("country", CStr(_dts.Tables(0).Rows(_i).Item(5)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(6)) = True Then
                    .AddParameter("postcode", "")
                Else
                    .AddParameter("postcode", CStr(_dts.Tables(0).Rows(_i).Item(6)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(7)) = True Then
                    .AddParameter("cont_name", "")
                Else
                    .AddParameter("cont_name", CStr(_dts.Tables(0).Rows(_i).Item(7)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(9)) = True Then
                    .AddParameter("telephone", "")
                Else
                    .AddParameter("telephone", CStr(_dts.Tables(0).Rows(_i).Item(9)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(10)) = True Then
                    .AddParameter("fax", "")
                Else
                    .AddParameter("fax", CStr(_dts.Tables(0).Rows(_i).Item(10)))
                End If

                If IsDBNull(_dts.Tables(0).Rows(_i).Item(11)) = True Then
                    .AddParameter("vat", "")
                Else
                    .AddParameter("vat", CStr(_dts.Tables(0).Rows(_i).Item(11)))
                End If
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(12)) = False Then .AddParameter("credit_limit", CDbl(_dts.Tables(0).Rows(_i).Item(12)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(13)) = False Then .AddParameter("std_tax_code", CStr(_dts.Tables(0).Rows(_i).Item(13)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(14)) = False Then .AddParameter("nc", CStr(_dts.Tables(0).Rows(_i).Item(14)))

                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub EAllVendor(ByVal _strPath As String)

        Dim _i As Integer
        Dim _ds As New DataSet
        Dim _myRow As DataRow

        _ds.ReadXml(_strPath)

        For Each _myRow In _ds.Tables(1).Rows
            Dim _cmd1 As New InfiniCommand("EINSERTALLVENDOR")
            With _cmd1

                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("VendorID", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("VendorName", CStr(_myRow.Item(2)))

                If IsDBNull(_myRow.Item(3)) = True Then
                    .AddParameter("street", "")
                Else
                    .AddParameter("street", CStr(_myRow.Item(3)))
                End If
                If IsDBNull(_myRow.Item(4)) = True Then
                    .AddParameter("town", "")
                Else
                    .AddParameter("town", CStr(_myRow.Item(4)))
                End If

                If IsDBNull(_myRow.Item(5)) = True Then
                    .AddParameter("country", "")
                Else
                    .AddParameter("country", CStr(_myRow.Item(5)))

                End If

                If IsDBNull(_myRow.Item(6)) = True Then
                    .AddParameter("postalcode", "")
                Else
                    .AddParameter("postalcode", CStr(_myRow.Item(6)))
                End If

                If IsDBNull(_myRow.Item(7)) = True Then
                    .AddParameter("contactname", "")
                Else
                    .AddParameter("contactname", CStr(_myRow.Item(7)))
                End If

                If IsDBNull(_myRow.Item(9)) = True Then
                    .AddParameter("telephone", "")
                Else
                    .AddParameter("telephone", CStr(_myRow.Item(9)))
                End If

                If IsDBNull(_myRow.Item(10)) = True Then
                    .AddParameter("fax", "")
                Else
                    .AddParameter("fax", CStr(_myRow.Item(10)))
                End If

                If IsDBNull(_myRow.Item(11)) = True Then
                    .AddParameter("vat", "")
                Else
                    .AddParameter("vat", CStr(_myRow.Item(11)))
                End If
                If IsDBNull(_myRow.Item(12)) = False Then .AddParameter("creditlimit", CDbl(_myRow.Item(12)))
                If IsDBNull(_myRow.Item(8)) = False Then .AddParameter("email", CStr(_myRow.Item(8)))
                .AddParameter("Syn", CStr(_myRow.Item(16)))
                .AddParameter("trandatetime", CStr(_myRow.Item(17)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub Company()
        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTCOMPANY")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("INSERTCOMPANY")
            With _cmd1
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("comp_name", CStr(_dts.Tables(0).Rows(_i).Item(1)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(2)) = False Then .AddParameter("street1", CStr(_dts.Tables(0).Rows(_i).Item(2)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(3)) = False Then .AddParameter("twon", CStr(_dts.Tables(0).Rows(_i).Item(3)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(4)) = False Then .AddParameter("p_code", CStr(_dts.Tables(0).Rows(_i).Item(4)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(5)) = False Then .AddParameter("country", CStr(_dts.Tables(0).Rows(_i).Item(5)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(6)) = False Then .AddParameter("telephone", CStr(_dts.Tables(0).Rows(_i).Item(6)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(7)) = False Then .AddParameter("fax", CStr(_dts.Tables(0).Rows(_i).Item(7)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(8)) = False Then .AddParameter("vat_no", CStr(_dts.Tables(0).Rows(_i).Item(9)))
                'If IsDBNull(_dts.Tables(0).Rows(_i).Item(9)) = False Then .AddParameter("vlogo", CStr(_dts.Tables(0).Rows(_i).Item(11)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub ECompany(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _myRow As DataRow

        _ds.ReadXml(_strPath)
        For Each _myRow In _ds.Tables(2).Rows
            Dim _cmd1 As New InfiniCommand("EINSERTCOMPANY")
            With _cmd1
                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("comp_name", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("street1", CStr(_myRow.Item(2)))
                If IsDBNull(_myRow.Item(3)) = False Then .AddParameter("twon", CStr(_myRow.Item(3)))
                If IsDBNull(_myRow.Item(4)) = False Then .AddParameter("p_code", CStr(_myRow.Item(4)))
                If IsDBNull(_myRow.Item(5)) = False Then .AddParameter("country", CStr(_myRow.Item(5)))
                If IsDBNull(_myRow.Item(6)) = False Then .AddParameter("telephone", CStr(_myRow.Item(6)))
                If IsDBNull(_myRow.Item(7)) = False Then .AddParameter("fax", CStr(_myRow.Item(7)))
                If IsDBNull(_myRow.Item(9)) = False Then .AddParameter("vat_no", CStr(_myRow.Item(9)))
                If IsDBNull(_myRow.Item(8)) = False Then .AddParameter("Email", CStr(_myRow.Item(8)))
                If IsDBNull(_myRow.Item(10)) = False Then .AddParameter("logo", CObj(_myRow.Item(10)))
                .AddParameter("syn", CStr(_myRow.Item(11)))
                .AddParameter("trandatetime", CStr(_myRow.Item(12)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub Bank()
        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTBANK")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("INSERTBANK")
            With _cmd1
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(0)) = False Then .AddParameter("ac", CStr(_dts.Tables(0).Rows(_i).Item(0)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(0)) = False Then .AddParameter("nc", CStr(_dts.Tables(0).Rows(_i).Item(0)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("name", Left(CStr(_dts.Tables(0).Rows(_i).Item(1)), 250))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(2)) = False Then .AddParameter("street", Left(CStr(_dts.Tables(0).Rows(_i).Item(2)), 250))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(3)) = False Then .AddParameter("town", CStr(_dts.Tables(0).Rows(_i).Item(3)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(4)) = False Then .AddParameter("country", CStr(_dts.Tables(0).Rows(_i).Item(4)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(5)) = False Then .AddParameter("post_code", Left(CStr(_dts.Tables(0).Rows(_i).Item(5)), 20))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(6)) = False Then .AddParameter("acct_name", Left(CStr(_dts.Tables(0).Rows(_i).Item(6)), 250))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(7)) = False Then .AddParameter("acct_number", CStr(_dts.Tables(0).Rows(_i).Item(7)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(8)) = False Then .AddParameter("sort_code", CStr(_dts.Tables(0).Rows(_i).Item(8)))
                .AddParameter("Ac_Type", "B")
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub EBank(ByVal _strPath As String)

        Dim _ds As New DataSet
        Dim _myRow As DataRow

        _ds.ReadXml(_strPath)
        For Each _myRow In _ds.Tables(3).Rows
            Dim _cmd1 As New InfiniCommand("EINSERTBANK")
            With _cmd1
                If IsDBNull(_myRow.Item(0)) = False Then .AddParameter("BankID", CStr(_myRow.Item(0)))
                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("BankName", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("Street", CStr(_myRow.Item(2)))
                If IsDBNull(_myRow.Item(3)) = False Then .AddParameter("Town", CStr(_myRow.Item(3)))
                If IsDBNull(_myRow.Item(4)) = False Then .AddParameter("Country", CStr(_myRow.Item(4)))
                If IsDBNull(_myRow.Item(5)) = False Then .AddParameter("PostalCode", CStr(_myRow.Item(5)))
                If IsDBNull(_myRow.Item(6)) = False Then .AddParameter("AccountName", CStr(_myRow.Item(6)))
                If IsDBNull(_myRow.Item(7)) = False Then .AddParameter("AccountNumber", CStr(_myRow.Item(7)))
                If IsDBNull(_myRow.Item(8)) = False Then .AddParameter("SortCode", CStr(_myRow.Item(8)))
                .AddParameter("syn", CStr(_myRow.Item(9)))
                .AddParameter("trandatetime", CStr(_myRow.Item(10)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub FinancialYear()
        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTF_YEAR")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("INSERTF_YEAR")
            With _cmd1
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(0)) = False Then .AddParameter("myno", CInt(_dts.Tables(0).Rows(_i).Item(0)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("f_month", CStr(_dts.Tables(0).Rows(_i).Item(1)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(2)) = False Then .AddParameter("f_year", CStr(_dts.Tables(0).Rows(_i).Item(2)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub EFinancialYear(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _myRow As DataRow

        _ds.ReadXml(_strPath)
        For Each _myRow In _ds.Tables(4).Rows
            Dim _cmd1 As New InfiniCommand("EINSERTF_YEAR")
            With _cmd1
                If IsDBNull(_myRow.Item(0)) = False Then .AddParameter("myno", CInt(_myRow.Item(0)))
                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("f_month", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("f_year", CStr(_myRow.Item(2)))
                .AddParameter("syn", CStr(_myRow.Item(3)))
                .AddParameter("trandatetime", CStr(_myRow.Item(4)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub UpdateTaxIDs()
        Dim _rCount As Integer, _i As Integer
        Dim _cmd As New InfiniCommand("SELECTTAXID")

        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        For _i = 0 To _rCount - 1
            Dim _cmd1 As New InfiniCommand("UPDATETAXID")
            With _cmd1
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(0)) = False Then .AddParameter("code", CStr(_dts.Tables(0).Rows(_i).Item(1)))
                If IsDBNull(_dts.Tables(0).Rows(_i).Item(1)) = False Then .AddParameter("rate", CStr(_dts.Tables(0).Rows(_i).Item(2)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub EUpdateTaxIDs(ByVal _strpath As String)
        Dim _ds As New DataSet
        Dim _myRow As DataRow

        _ds.ReadXml(_strpath)
        For Each _myRow In _ds.Tables(5).Rows
            Dim _cmd1 As New InfiniCommand("EUPDATETAXID")
            With _cmd1
                If IsDBNull(_myRow.Item(1)) = False Then .AddParameter("TaxID", CStr(_myRow.Item(1)))
                If IsDBNull(_myRow.Item(2)) = False Then .AddParameter("TaxRate", CStr(_myRow.Item(2)))
                Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd1)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub Transaction()
        Dim _rCount As Integer, _i As Integer = 1
        Dim _dRow As DataRow
        Dim _ds As DataSet
        Dim _strChk As String
        Dim _cmd As New InfiniCommand("SELECTTRANSACTION")
        _dts = ClassGlobal.GetDataSet(_cmd)
        _rCount = _dts.Tables(0).Rows.Count

        '_rCount = _rCount + 1
        'For _i = 1 To _rCount
        For Each _dRow In _dts.Tables(0).Rows
            Dim _cmd1 As New InfiniCommand("SELECTRECORD")
            _cmd1.AddParameter("parentid", _dRow.Item("Parentid"))

            _ds = ClassGlobal.GetDataSet(_cmd1)

            If _ds.Tables(0).Rows.Count <> 0 Then

                'Insert Into Transaction
                Dim _cmd2 As New InfiniCommand("INSERTTRANSACTION_PRO")
                With _cmd2

                    .AddParameter("record_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))
                    .AddParameter("activity_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))
                    .AddParameter("tran_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))

                    Dim aas As String = CStr(_ds.Tables(0).Rows(0).Item(1))

                    If CStr(_ds.Tables(0).Rows(0).Item(1)) = "SR" Or CStr(_ds.Tables(0).Rows(0).Item(1)) = "PP" Or CStr(_ds.Tables(0).Rows(0).Item(1)) = "SA" Or CStr(_ds.Tables(0).Rows(0).Item(1)) = "PA" Then
                        .AddParameter("ac", CStr(_ds.Tables(0).Rows(0).Item(21))) '19
                        .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(17))) '15
                    Else
                        .AddParameter("ac", CStr(_ds.Tables(0).Rows(0).Item(17))) '15
                        _strChk = CStr(_ds.Tables(0).Rows(0).Item(21)) '19
                        If Len(_strChk) > 5 Then
                            .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(17))) '15
                        Else
                            .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(21))) '17
                        End If
                    End If

                    .AddParameter("tc", CStr(_ds.Tables(0).Rows(0).Item(2)))
                    .AddParameter("details", CStr(_ds.Tables(0).Rows(0).Item(4)))
                    .AddParameter("net", CDbl(_ds.Tables(0).Rows(0).Item(5)))
                    .AddParameter("vat", CDbl(_ds.Tables(0).Rows(0).Item(6)))
                    .AddParameter("tp", CStr(_ds.Tables(0).Rows(0).Item(1)))
                    .AddParameter("v", CStr(_ds.Tables(0).Rows(0).Item(9)))
                    .AddParameter("transfor", CStr(_ds.Tables(0).Rows(0).Item(20))) '18
                    .AddParameter("w", CStr(_ds.Tables(0).Rows(0).Item(8)))
                    .AddParameter("markw", CStr(_ds.Tables(0).Rows(0).Item(7)))
                    .AddParameter("date", CStr(_ds.Tables(0).Rows(0).Item(3)))
                    .AddParameter("ref", "")
                    Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
                End With
            End If
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        'Dim _rCount As Integer, _i As Integer = 1
        'Dim _ds As DataSet
        'Dim _strChk As String
        'Dim _cmd As New InfiniCommand("SELECTTRANSACTION")
        '_dts = ClassGlobal.GetDataSet(_cmd)
        '_rCount = _dts.Tables(0).Rows.Count

        '_rCount = _rCount + 1
        'For _i = 1 To _rCount
        '    Dim _cmd1 As New InfiniCommand("SELECTRECORD")
        '    _cmd1.AddParameter("parentid", _i)

        '    _ds = ClassGlobal.GetDataSet(_cmd1)
        '    Dim f As Integer = _ds.Tables(0).Rows.Count
        '    If _ds.Tables(0).Rows.Count <> 0 Then

        '        'Insert Into Transaction
        '        Dim _cmd2 As New InfiniCommand("INSERTTRANSACTION_PRO")
        '        With _cmd2

        '            .AddParameter("record_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))
        '            .AddParameter("activity_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))
        '            .AddParameter("tran_no", CDbl(_ds.Tables(0).Rows(0).Item(0)))

        '            Dim aas As String = CStr(_ds.Tables(0).Rows(0).Item(1))

        '            If CStr(_ds.Tables(0).Rows(0).Item(1)) = "SR" Or CStr(_ds.Tables(0).Rows(0).Item(1)) = "PP" Then
        '                .AddParameter("ac", CStr(_ds.Tables(0).Rows(0).Item(19)))
        '                .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(15)))
        '            Else
        '                .AddParameter("ac", CStr(_ds.Tables(0).Rows(0).Item(15)))
        '                _strChk = CStr(_ds.Tables(0).Rows(0).Item(19))
        '                If Len(_strChk) > 5 Then
        '                    .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(15)))
        '                Else
        '                    .AddParameter("nc", CStr(_ds.Tables(0).Rows(0).Item(19)))
        '                End If
        '            End If

        '            .AddParameter("tc", CStr(_ds.Tables(0).Rows(0).Item(2)))
        '            .AddParameter("details", CStr(_ds.Tables(0).Rows(0).Item(4)))
        '            .AddParameter("net", CDbl(_ds.Tables(0).Rows(0).Item(5)))
        '            .AddParameter("vat", CDbl(_ds.Tables(0).Rows(0).Item(6)))
        '            .AddParameter("tp", CStr(_ds.Tables(0).Rows(0).Item(1)))
        '            .AddParameter("v", CStr(_ds.Tables(0).Rows(0).Item(9)))
        '            .AddParameter("transfor", CStr(_ds.Tables(0).Rows(0).Item(18)))
        '            .AddParameter("w", CStr(_ds.Tables(0).Rows(0).Item(8)))
        '            .AddParameter("markw", CStr(_ds.Tables(0).Rows(0).Item(7)))
        '            .AddParameter("date", CStr(_ds.Tables(0).Rows(0).Item(3)))
        '            .AddParameter("ref", "")
        '            Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
        '        End With
        '    End If
        'Next _i
        'If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Public Sub ETransaction(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _myRow As DataRow
        Dim _strChk As String

        _ds.ReadXml(_strPath)
        For Each _myRow In _ds.Tables(7).Rows

            'Insert Into Transaction
            Dim _cmd2 As New InfiniCommand("EINSERTTRANSACTION")
            With _cmd2
                .AddParameter("Parentid", CDbl(_myRow.Item(0)))
                .AddParameter("invtp", CStr(_myRow.Item(1)))
                .AddParameter("taxid", CStr(_myRow.Item(2)))
                .AddParameter("invdate", CStr(_myRow.Item(3)))
                .AddParameter("invdetails", CStr(_myRow.Item(4)))
                .AddParameter("invnet", CDbl(_myRow.Item(5)))
                .AddParameter("invvat", CDbl(_myRow.Item(6)))
                .AddParameter("refund", CStr(_myRow.Item(7)))
                .AddParameter("mark", CStr(_myRow.Item(8)))
                .AddParameter("v", CStr(_myRow.Item(9)))
                .AddParameter("trandate", CDate(_myRow.Item(10)))
                .AddParameter("syn", CStr(_myRow.Item(11)))
                .AddParameter("trandatetime", CStr(_myRow.Item(12)))
                Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub Ledger()
        'Dim _rCount As Integer, _i As Integer = 1
        'Dim _ds As New DataSet
        'Dim _strChk As String
        'Dim _dr As DataRow
        'Dim _cmd As New InfiniCommand("SELECTLEDGER")

        '_dts = ClassGlobal.GetDataSet(_cmd)
        '_rCount = _dts.Tables(0).Rows.Count

        '_rCount = _rCount + 1
        'For _i = 1 To _rCount
        '    Dim _cmd1 As New InfiniCommand("SELECTLEDGERRECORD")
        '    _cmd1.AddParameter("parentid", _i)

        '    _ds = ClassGlobal.GetDataSet(_cmd1)

        '    If _ds.Tables(0).Rows.Count <> 0 Then

        '        For Each _dr In _ds.Tables(0).Rows
        '            'Insert Into Ledger
        '            Dim _cmd2 As New InfiniCommand("INSERTLEDGER_PRO")
        '            With _cmd2

        '                Dim r As Object
        '                .AddParameter("record_no", CDbl(_dr.Item(13)))
        '                .AddParameter("activity_no", CDbl(_dr.Item(13)))
        '                .AddParameter("tran_no", CDbl(_dr.Item(13)))
        '                .AddParameter("nc", CStr(_dr.Item(15)))
        '                .AddParameter("debit", CDbl(_dr.Item(16)))
        '                .AddParameter("credit", CDbl(_dr.Item(17)))
        '                .AddParameter("tran_type", CStr(_dr.Item(18)))
        '                .AddParameter("refn", CStr(_dr.Item(19)))

        '                If Len(_dr.Item(20)) = 0 Then
        '                    .AddParameter("pay_status", " ")
        '                Else
        '                    .AddParameter("pay_status", CStr(_dr.Item(20)))
        '                End If

        '                .AddParameter("myauto", CDbl(_dr.Item(21)))
        '                .AddParameter("details", " ")
        '                Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
        '            End With
        '        Next
        '    End If
        'Next _i
        'If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        'NEW
        Dim _rCount As Integer, _i As Integer = 1
        Dim _ds As New DataSet
        Dim _dRow As DataRow
        Dim _strChk As String
        Dim _dr As DataRow
        Dim _cmd As New InfiniCommand("SELECTLEDGER")

        _dts = ClassGlobal.GetDataSet(_cmd)

        For Each _dRow In _dts.Tables(0).Rows
            Dim _cmd1 As New InfiniCommand("SELECTLEDGERRECORD")
            _cmd1.AddParameter("parentid", _dRow.Item("parentid"))

            _ds = ClassGlobal.GetDataSet(_cmd1)

            'If _ds.Tables(0).Rows.Count <> 0 Then

            For Each _dr In _ds.Tables(0).Rows
                'Insert Into Ledger
                Dim _cmd2 As New InfiniCommand("INSERTLEDGER_PRO")
                With _cmd2

                    Dim r As Object
                    .AddParameter("record_no", CDbl(_dr.Item(15)))
                    .AddParameter("activity_no", CDbl(_dr.Item(15)))
                    .AddParameter("tran_no", CDbl(_dr.Item(15)))
                    .AddParameter("nc", CStr(_dr.Item(17)))
                    .AddParameter("debit", CDbl(_dr.Item(18)))
                    .AddParameter("credit", CDbl(_dr.Item(19)))
                    .AddParameter("tran_type", CStr(_dr.Item(20)))
                    .AddParameter("refn", CStr(_dr.Item(21)))

                    If Len(_dr.Item(22)) = 0 Then
                        .AddParameter("pay_status", " ")
                    Else
                        .AddParameter("pay_status", CStr(_dr.Item(22)))
                    End If

                    .AddParameter("myauto", CDbl(_dr.Item(23)))
                    .AddParameter("details", " ")
                    Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
                End With
            Next
            'End If
        Next
        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub
    End Sub

    Public Sub ELedger(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _dr As DataRow

        _ds.ReadXml(_strPath)
        For Each _dr In _ds.Tables(8).Rows
            'Insert Into Ledger
            Dim _cmd2 As New InfiniCommand("EINSERTLEDGER")
            With _cmd2
                .AddParameter("ParentID", CDbl(_dr.Item(0)))
                .AddParameter("ChildID", CDbl(_dr.Item(1)))
                .AddParameter("TransID", CStr(_dr.Item(2)))
                .AddParameter("Debit", CDbl(_dr.Item(3)))
                .AddParameter("Credit", CDbl(_dr.Item(4)))
                .AddParameter("TranType", CStr(_dr.Item(5)))
                .AddParameter("TranRef", CStr(_dr.Item(6)))
                .AddParameter("PayStatus", CStr(_dr.Item(7)))
                .AddParameter("MyAutoNo", CDbl(_dr.Item(8)))
                .AddParameter("syn", CStr(_dr.Item(9)))
                .AddParameter("trandatetime", CStr(_dr.Item(10)))
                Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub ECREDITCARDINFO(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _dr As DataRow

        _ds.ReadXml(_strPath)
        For Each _dr In _ds.Tables(8).Rows
            'Insert Into Ledger
            Dim _cmd2 As New InfiniCommand("EINSERTCCINFO")
            With _cmd2
                .AddParameter("Inv_no", CDbl(_dr.Item(0)))
                .AddParameter("Inv_date", CDbl(_dr.Item(1)))
                .AddParameter("Debit", CStr(_dr.Item(2)))
                .AddParameter("credit", CDbl(_dr.Item(3)))
                .AddParameter("ShippingComp", CDbl(_dr.Item(4)))
                .AddParameter("TrackingNumber", CStr(_dr.Item(5)))
                .AddParameter("DeliveryAddress", CStr(_dr.Item(6)))
                .AddParameter("CardType", CStr(_dr.Item(7)))
                .AddParameter("CardHolderName", CDbl(_dr.Item(8)))
                .AddParameter("CardNumber", CStr(_dr.Item(9)))
                .AddParameter("ExpairyDate", CStr(_dr.Item(10)))
                Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub
    End Sub

    Public Sub Outstanding()
        Dim _rCount As Integer, _i As Integer = 1
        Dim _ds As New DataSet
        Dim _dr As DataRow
        Dim _strChk As String

        Dim _cmd As New InfiniCommand("SELECTOUTSTANDING")
        _ds = ClassGlobal.GetDataSet(_cmd)

        If _ds.Tables(0).Rows.Count > 0 Then

            For Each _dr In _ds.Tables(0).Rows
                'Insert Into Outstanding
                Dim _cmd2 As New InfiniCommand("INSERTOUTSTANDING_PRO")
                With _cmd2
                    .AddParameter("myauto", CDbl(_dr.Item(0)))
                    .AddParameter("r_no", CDbl(_dr.Item(1)))
                    .AddParameter("a_no", CDbl(_dr.Item(1)))
                    .AddParameter("t_no", CDbl(_dr.Item(1)))
                    .AddParameter("nc", CStr(_dr.Item(3)))
                    .AddParameter("tran_type", CStr(_dr.Item(6)))
                    .AddParameter("debit", CDbl(_dr.Item(4)))
                    .AddParameter("credit", CDbl(_dr.Item(5)))
                    .AddParameter("refn", CStr(_dr.Item(8)))
                    .AddParameter("details", CStr(_dr.Item(7)))

                    If Len(_dr.Item(4)) > 0 Then
                        .AddParameter("net", CDbl(_dr.Item(4)))
                    Else
                        .AddParameter("net", CDbl(_dr.Item(5)))
                    End If

                    .AddParameter("odate", CStr(_dr.Item(9)))
                    .AddParameter("[from]", CDbl(_dr.Item(10)))
                    .AddParameter("[to]", CDbl(_dr.Item(11)))
                    .AddParameter("ft", CStr(_dr.Item(12)))

                    Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
                End With
            Next
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub
        End If

    End Sub

    Public Sub EOutstanding(ByVal _strPath As String)
        Dim _ds As New DataSet
        Dim _dr As DataRow

        _ds.ReadXml(_strPath)
        For Each _dr In _ds.Tables(6).Rows

            'Insert Into Outstanding
            Dim _cmd2 As New InfiniCommand("EINSERTOUTSTANDING")
            With _cmd2
                .AddParameter("myauto", CDbl(_dr.Item(0)))
                .AddParameter("Parentid", CDbl(_dr.Item(1)))
                .AddParameter("Childid", CDbl(_dr.Item(2)))
                .AddParameter("transif", CStr(_dr.Item(3)))
                .AddParameter("Debit", CStr(_dr.Item(4)))
                .AddParameter("Credit", CStr(_dr.Item(5)))
                .AddParameter("trantype", CStr(_dr.Item(6)))
                .AddParameter("Details", CStr(_dr.Item(7)))
                .AddParameter("tranref", CStr(_dr.Item(8)))
                .AddParameter("osdate", CStr(_dr.Item(9)))
                .AddParameter("osfrom", CDbl(_dr.Item(10)))
                .AddParameter("osto", CDbl(_dr.Item(11)))
                .AddParameter("osft", CStr(_dr.Item(12)))
                .AddParameter("syn", CStr(_dr.Item(13)))
                .AddParameter("trandatetime", CStr(_dr.Item(14)))
                Dim id As Integer = ClassGlobal.ExecuteNonQuery(_cmd2)
            End With
        Next
        If ClsGlobal.ExecuteCommand = False Then Exit Sub

    End Sub

    Public Sub N_Record()
        Dim i As Integer, _tkey As Integer, _inc As Integer, _inct As Integer, _tkeyM As Integer, _tkeyY As Integer, _counter As Integer
        Dim _dr As IDataReader, _dr1 As IDataReader, _dr2 As IDataReader
        Dim tolNum As Integer, _i As Integer
        Dim hTable As New Hashtable, _hmtable As New Hashtable
        Dim _htable As New Hashtable
        Dim _mStr As String, _invdate As String
        Dim _yStr As String
        Dim _dStr As String
        Dim cmdcollection As New Collection

        'Adding LedgerID In Collection Table
        Dim _cmd As New InfiniCommand("SELECTNRECORD")

        _dr = ClsGlobal.GetDataReader(_cmd)
        _dr.Read()

        _invdate = CStr(_dr.Item("invdate"))
        _dr = Nothing
        Dim r As Integer

        r = CInt(Len(_invdate)) '_invdate.Length)

        Dim rr As Integer = InStr(1, _invdate, "/", CompareMethod.Text)

        Dim u As String = Right(_invdate, r - rr)

        Dim q As Integer = Len(u)

        _yStr = Right(u, q - 3)

        'Get Months & years from Acc/Pro
        _dr2 = GetProMonth()

        _inct = 1
        Do While _dr2.Read = True
            _htable.Add(_inct, CStr(_dr2.Item("f_year")))
            _hmtable.Add(_inct, CStr(_dr2.Item("f_month")))
            _inct = _inct + 1
        Loop
        _dr2 = Nothing

        'Add LedgerID Into Table
        hTable.Add(1, 10000)
        hTable.Add(2, 71100)
        hTable.Add(3, 70100)
        hTable.Add(4, 70200)
        hTable.Add(5, 49999)
        'For Vendor
        hTable.Add(6, 20000)
        hTable.Add(7, 71101)
        hTable.Add(8, 71000)

        'LedgerId
        _tkey = hTable.Keys.Count()
        'year
        _tkeyY = _htable.Keys.Count()
        'Year


        For _inc = 1 To _tkey

            For _i = 1 To _tkeyY
                Dim _cmd101 As New InfiniCommand("INSERTN_RECORD")

                With _cmd101
                    .AddParameter("nc", CStr(hTable.Item(_inc)))
                    .AddParameter("year", CStr(_htable.Item(_i)))
                    .AddParameter("month", CStr(_hmtable.Item(_i)))
                End With
                Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd101)
            Next
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub
        Next
        cmdcollection = Nothing
        'Inserting entries in N_Record Table
        Dim _cmd123 As New InfiniCommand("SELECTNRECORD")
        _dr = ClsGlobal.GetDataReader(_cmd)

        Do While _dr.Read = True
            UpdateNrecord(CStr(_dr.Item("transid")), CStr(_dr.Item("invtp")), CStr(_dr.Item("invdate")), CDbl(_dr.Item("Invnet")), CDbl(_dr.Item("invvat")), CStr(_dr.Item("trantype")))
        Loop

    End Sub
    Private Function GetProMonth() As IDataReader
        Dim _cmd1 As New InfiniCommand("GETPROMONTH")
        Return ClassGlobal.GetDataReader_Pro(_cmd1)
    End Function
    Private Sub UpdateNrecord(ByVal i As Integer)

        Dim _cmd As New InfiniCommand("UPDATENRECORD")
        _cmd.AddParameter("myauto", CInt(i))
        Dim Id As Integer = ClassGlobal.ExecuteNonQuery(_cmd)

        If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

    End Sub

    Private Function RecordCount() As IDataReader
        Dim _cmd As New InfiniCommand("NRECORD")
        Return ClassGlobal.GetDataReader_Pro(_cmd)
    End Function

    Private Sub updateNrecord(ByVal _transid As String, ByVal _invtp As String, ByVal _invdate As String, ByVal _invnet As Double, ByVal _invvat As Double, ByVal _invtype As String)
        Dim _mStr As String, _yStr As String
        Dim _dStr As String
        Dim r As Integer, _mm As String

        r = CInt(Len(_invdate)) '_invdate.Length)

        Dim rr As Integer = InStr(1, _invdate, "/", CompareMethod.Text)

        Dim u As String = Right(_invdate, r - rr)

        Dim q As Integer = Len(u)

        _yStr = Right(u, q - 3)
        _mm = Left(u, q - 5)
        Dim ui As Integer = CInt(_mm)
        _mStr = getmonth(ui)


        'Month In Words
        '_dStr = Format(CDate(_invdate), "dd/MM/yyyy")
        '_mStr = Format(CDate(_dStr), "MMMM")

        ''Year In Number
        '_dStr = Format(CDate(_invdate), "yyyy/dd/MM")
        '_yStr = CStr(Format(CDate(_dStr), "yyyy"))

        'Update data in N_record table
        updateN_Record(_mStr, _yStr, _transid, _invtp, CDbl(Format(_invnet, "0.00")), CDbl(Format(_invvat, "0.00")), _invtype)

    End Sub
#End Region

    Public Sub updateN_Record(ByVal _mstr As String, ByVal _ystr As String, ByVal _transid As String, ByVal _invtp As String, ByVal _invnet As Double, ByVal _invvat As Double, ByVal _invtype As String)
        Dim CmdCollection As New Collection
        Dim _dr As IDataReader
        Dim _amt As Double = 0.0

        'Process for "SI"
        'Select record from N_record table
        If _invtp = "SI" And _transid = "10000" And _invtype = "U" Then
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = CDbl(_amt - _invnet)

            'Update for ledgerID 10000
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr(_transid))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 71100
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = CDbl(_amt - _invvat)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            'Update for ledgerID 70100
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = CDbl(_amt + (_invnet + _invvat))

            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("70100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        ElseIf _invtp = "SC" And _transid = "10000" And _invtype = "U" Then

            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = CDbl(_amt + _invnet)
            'Update for ledgerID 10000
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr(_transid))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 71100
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = _amt + CDbl(_invvat)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            'Update for ledgerID 70100
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("70100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        End If

        '***************************************************************
        'Process for "PI"
        'Select record from N_record table
        If _invtp = "PI" And _transid = "20000" And _invtype = "U" Then
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet)

            'Update for ledgerID 20000
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr(_transid))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 71000
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invnet + _invvat)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            'Update for ledgerID 71101
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71101")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invvat)

            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("71101"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        ElseIf _invtp = "PC" And _transid = "20000" And _invtype = "U" Then

            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invnet)
            'Update for ledgerID 20000
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr(_transid))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 71000
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet + _invvat)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            'Update for ledgerID 71101
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71101")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invvat)

            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("71101"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        End If

        '***************************************************************
        ' Process for "SR"
        '********************************************************************
        If _invtp = "SR" And _transid = "70200" And _invtype = "B" Then
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet)

            'Update for ledgerID 70200
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr("70200"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 70100
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invnet)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("70100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub
        End If


        '***************************************************************
        ' Process for "PP"
        '********************************************************************
        If _invtp = "PP" And _transid = "70200" And _invtype = "B" Then
            _dr = fatchrecord(_mstr, _ystr, _transid)
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invnet + _invvat)

            'Update for ledgerID 70200
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr("70200"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)

            'Update for ledgerID 71000
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet + _invvat)

            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub
        End If

        '***************************************************************

        'Process for Vendor refund
        '*******************************************************************
        _amt = 0.0
        If _invtp = "PC" And _transid = "20000" And _invtype = "T" Then
            _dr = fatchrecord(_mstr, _ystr, "20000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invnet)

            'Update for ledgerID 70100
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr("20000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '2)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71101")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(_invvat)

            'Update for ledgerID 71100
            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71101"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            '3)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet + _invvat)

            'Update for ledgerID 10000
            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("71000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)

            '3)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd3 As New InfiniCommand("UPNRECORD")
            With _cmd3
                .AddParameter("nc", CStr("71000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id3 As Integer = ClsGlobal.ExecuteNonQuery(_cmd3)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '4)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "49999")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd4 As New InfiniCommand("UPNRECORD")
            With _cmd4
                .AddParameter("nc", CStr("49999"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id4 As Integer = ClsGlobal.ExecuteNonQuery(_cmd4)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '5)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "49999")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd5 As New InfiniCommand("UPNRECORD")
            With _cmd5
                .AddParameter("nc", CStr("49999"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id5 As Integer = ClsGlobal.ExecuteNonQuery(_cmd5)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '6)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70200")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd6 As New InfiniCommand("UPNRECORD")
            With _cmd6
                .AddParameter("nc", CStr("70200"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id6 As Integer = ClsGlobal.ExecuteNonQuery(_cmd6)

            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        End If

        '***************************************************************
        'Process for Customer refund
        '*******************************************************************
        _amt = 0.0
        If _invtp = "SC" And _transid = "10000" And _invtype = "T" Then
            _dr = fatchrecord(_mstr, _ystr, "70100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd As New InfiniCommand("UPNRECORD")
            With _cmd
                .AddParameter("nc", CStr("70100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id As Integer = ClsGlobal.ExecuteNonQuery(_cmd)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '2)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "71100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invvat)

            'Update for ledgerID 71100
            Dim _cmd1 As New InfiniCommand("UPNRECORD")
            With _cmd1
                .AddParameter("nc", CStr("71100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id1 As Integer = ClsGlobal.ExecuteNonQuery(_cmd1)

            '3)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "10000")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(_invnet)

            'Update for ledgerID 10000
            Dim _cmd2 As New InfiniCommand("UPNRECORD")
            With _cmd2
                .AddParameter("nc", CStr("10000"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id2 As Integer = ClsGlobal.ExecuteNonQuery(_cmd2)

            '3)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70100")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd3 As New InfiniCommand("UPNRECORD")
            With _cmd3
                .AddParameter("nc", CStr("70100"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id3 As Integer = ClsGlobal.ExecuteNonQuery(_cmd3)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '4)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "49999")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd4 As New InfiniCommand("UPNRECORD")
            With _cmd4
                .AddParameter("nc", CStr("49999"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id4 As Integer = ClsGlobal.ExecuteNonQuery(_cmd4)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '5)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "49999")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt + CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd5 As New InfiniCommand("UPNRECORD")
            With _cmd5
                .AddParameter("nc", CStr("49999"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id5 As Integer = ClsGlobal.ExecuteNonQuery(_cmd5)
            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

            '6)
            _amt = 0.0
            _dr = fatchrecord(_mstr, _ystr, "70200")
            _dr.Read()
            _amt = CDbl(Format(_dr.Item("actual"), "0.00")) '_amt = CDbl(_dr.Item("actual"))
            _dr = Nothing

            _amt = _amt - CDbl(CDbl(_invnet) + CDbl(_invvat))

            'Update for ledgerID 70100
            Dim _cmd6 As New InfiniCommand("UPNRECORD")
            With _cmd6
                .AddParameter("nc", CStr("70200"))
                .AddParameter("year", CStr(_ystr))
                .AddParameter("month", CStr(_mstr))
                .AddParameter("actual", CDbl(_amt))
            End With
            Dim Id6 As Integer = ClsGlobal.ExecuteNonQuery(_cmd6)

            If ClsGlobal.ExecuteCommandFor_Pro = False Then Exit Sub

        End If

    End Sub
    Private Function fatchrecord(ByVal _mstr As String, ByVal _ystr As String, ByVal _transid As String) As IDataReader
        Dim _dr As IDataReader
        Dim _cmd As New InfiniCommand("FATCHRECORD")

        With _cmd
            .AddParameter("nc", CStr(_transid))
            .AddParameter("month", CStr(_mstr))
            .AddParameter("year", CStr(_ystr))
        End With

        Return ClassGlobal.GetDataReader_Pro(_cmd)
    End Function

    Private Function getmonth(ByVal _int As Integer) As String
        If _int = 1 Then
            Return "January"
        ElseIf _int = 2 Then
            Return "February"
        ElseIf _int = 3 Then
            Return "March"
        ElseIf _int = 4 Then
            Return "April"
        ElseIf _int = 5 Then
            Return "May"
        ElseIf _int = 6 Then
            Return "June"
        ElseIf _int = 7 Then
            Return "July"
        ElseIf _int = 8 Then
            Return "August"
        ElseIf _int = 9 Then
            Return "September"
        ElseIf _int = 10 Then
            Return "October"
        ElseIf _int = 11 Then
            Return "November"
        ElseIf _int = 12 Then
            Return "December"



        End If

    End Function

#Region "BackUp Section"

    Public Function getXml4AllTable(ByVal _strPath As String, ByVal _cmdname As String, ByVal _filename As String) As DataSet

        _cmd = New InfiniCommand(_cmdname)
        _ds = ClassGlobal.GetDataSet(_cmd)
        strPath = _strPath & _filename
        _ds.WriteXml(strPath, XmlWriteMode.WriteSchema)

    End Function

    Public Function WriteXML4BackUp(ByVal _strpath As String) As Integer 'DataSet
        Dim cmdcollection As New Collection

        Dim cmd0 As New InfiniCommand("SELECTALLCUSTOMER")
        cmdcollection.Add(cmd0)

        Dim cmd1 As New InfiniCommand("SELECTALLVENDOR")
        cmdcollection.Add(cmd1)

        Dim cmd2 As New InfiniCommand("SELECTCOMPANY")
        cmdcollection.Add(cmd2)

        Dim cmd3 As New InfiniCommand("SELECTBANK")
        cmdcollection.Add(cmd3)

        Dim cmd4 As New InfiniCommand("SELECTF_YEAR")
        cmdcollection.Add(cmd4)

        Dim cmd5 As New InfiniCommand("SELECTTAXID")
        cmdcollection.Add(cmd5)

        Dim cmd6 As New InfiniCommand("SELECTOUTSTANDING")
        cmdcollection.Add(cmd6)

        Dim cmd7 As New InfiniCommand("RSELECTTRANSACTION")
        cmdcollection.Add(cmd7)

        Dim cmd8 As New InfiniCommand("RSELECTLEDGER")
        cmdcollection.Add(cmd8)

        Dim cmd9 As New InfiniCommand("GETCCINFO")
        cmdcollection.Add(cmd9)

        _dts = AccessHelper.ExeComman4WriteXML(cmdcollection)

        If _dts.Tables(0).Rows.Count = 0 And _dts.Tables(1).Rows.Count = 0 Then
            MsgBox("Cannot take empty database backup.", MsgBoxStyle.Information, "Infini Express")
            Return 1
            Exit Function
        Else
            _strpath = _strpath & "ExpressBackUp.XML"
            _dts.WriteXml(_strpath, XmlWriteMode.DiffGram)
            cmdcollection = Nothing
            Return 0
        End If

    End Function
#End Region
End Class
