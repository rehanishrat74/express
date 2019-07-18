'Option Strict On
#Region "Import Libraries"
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
'Imports System.IO.Path
Imports Microsoft.VisualBasic
Imports Path = System.IO.Path
#End Region

Public Class AccessHelper

    Private Shared _cn As IDbConnection
    Private Shared _cmd As IDbCommand
    Private Shared _strConnectionString As String
    Private Shared _InTransaction As IDbTransaction
    Private Shared AccPath As String

    Public Shared Property DbPath() As String
        Get
            Return AccPath
        End Get
        Set(ByVal Value As String)
            AccPath = Value
        End Set
    End Property

    Public Shared Function GetConnectionString() As String

        Dim AppPath As String, ConnectionString As String

        'Dim mstrPath As String
        'mstrPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModule(0).FullyQualifiedName)
        'mstrPath = Application.ExecutablePath()

        AppPath = AppDomain.CurrentDomain.BaseDirectory()

        If InStr(AppPath, "bin") > 0 Then
            AppPath = AppPath.Replace("bin\", "DataBase\ExpressDB.mdb") 'For Developers
        Else
            AppPath = AppPath & "DataBase\ExpressDB.mdb" 'For User
        End If
        'MsgBox(AppPath)
        'AppPath = Path.GetFullPath("ExpressDB.mdb")

        'If InStr(AppPath, "bin") > 0 Then
        '    AppPath = AppPath.Replace("bin", "DataBase") 'For Developers
        'Else
        '    AppPath = Path.GetFullPath("DataBase\ExpressDB.mdb")  'For Users
        'End If

        ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & AppPath & "; Jet OLEDB:Database Password=n7e9s0t1le23"

        Return ConnectionString

    End Function

    Public Shared Function GetConnectionString_Pro() As String
        Dim ConnectionString As String
        ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & DbPath & "; Jet OLEDB:Database Password=n7e9s0t1le23"
        Return ConnectionString

    End Function

#Region " Infini Accounts Express overloads Functions"

    Public Overloads Shared Function ExecuteReader(ByRef ObjCmd As InfiniCommand) As IDataReader

        Dim dr As IDataReader

        PrepareCommand(ObjCmd)

        dr = _cmd.ExecuteReader()

        Return dr

    End Function
    Public Overloads Shared Function ExecuteReader_Pro(ByRef ObjCmd As InfiniCommand) As IDataReader

        Dim dr As IDataReader

        PrepareCommand_Pro(ObjCmd)

        dr = _cmd.ExecuteReader()

        Return dr

    End Function


    Public Overloads Shared Function ExecuteDataset(ByRef ObjCmd As InfiniCommand) As DataSet

        Dim ds As New DataSet
        PrepareCommand(ObjCmd)
        'create the DataAdapter & DataSet
        Dim da As OleDbDataAdapter = New OleDbDataAdapter(CType(_cmd, OleDbCommand))
        'fill the DataSet using default values for DataTable names, etc.
        da.Fill(ds)
        'return the dataset
        Return ds

    End Function

    Public Overloads Shared Function ExecuteNonQuery(ByRef ObjCmd As InfiniCommand) As Integer

        Dim returnValue As Integer = 1
        'PrepareCommand(ObjCmd)
        '_cmd.CommandTimeout = 300
        ''finally, execute the command.
        'returnValue = _cmd.ExecuteNonQuery()
        '_cn.Close()
        Return returnValue

    End Function

    Public Overloads Shared Function ExecuteCollection(ByRef _cmdCollection As Collection) As Boolean

        Dim _Strsql As String
        Dim i As Integer = 0
        Dim tempcmd As New InfiniCommand("Nothing")

        _strConnectionString = GetConnectionString()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
                _InTransaction = _cn.BeginTransaction(IsolationLevel.ReadCommitted)
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        For Each tempcmd In _cmdCollection

            i += 1
            tempcmd = CType(_cmdCollection(i), InfiniCommand)

            If tempcmd.CommandName = "UPDATECOMPANYLOGO" Then
                LogoUpdate(tempcmd)
            Else

                Dim SPName As String
                SPName = tempcmd.CommandName
                SPName = SPName.Substring(0, 3)
                If SPName = "Tbl" Then
                    _Strsql = AccessQueries.GetSynchronizeQueryString(tempcmd)
                Else
                    _Strsql = AccessQueries.GetQueryString(tempcmd)
                End If
                'set the command text SQL statement
                _cmd.CommandText = _Strsql
                'associate the connection with the command
                _cmd.Connection = _cn

                _cmd.Transaction = _InTransaction

            End If
            _cmd.CommandTimeout = 300
            'finally, execute the command.
            Try
                _cmd.ExecuteNonQuery()
            Catch e As Exception
                _InTransaction.Rollback()
                Throw New Exception(" Transaction rolled back because of the following error. " & e.Message)
                Return False
            End Try
        Next
        _cmdCollection = New Collection
        _InTransaction.Commit()
        _cn.Close()
        Return True

    End Function

    Public Overloads Shared Function ExecuteSynchronizeCollect(ByRef cmdCollect As Collection) As Boolean

        Dim _Strsql As String
        Dim i As Integer = 0
        Dim tcmd As New InfiniCommand("Nothing")

        _strConnectionString = GetConnectionString()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
                _InTransaction = _cn.BeginTransaction(IsolationLevel.ReadCommitted)
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        For Each tcmd In cmdCollect

            i += 1
            tcmd = CType(cmdCollect(i), InfiniCommand)

            _Strsql = AccessQueries.GetSynchronizeQueryString(tcmd)

            'set the command text SQL statement
            _cmd.CommandText = _Strsql
            'associate the connection with the command
            _cmd.Connection = _cn

            _cmd.Transaction = _InTransaction

            _cmd.CommandTimeout = 300
            'finally, execute the command.
            Try
                _cmd.ExecuteNonQuery()
            Catch e As Exception
                _InTransaction.Rollback()
                Throw New Exception(" Transaction rolled back because of the following error. " & e.Message)
                Return False
            End Try
        Next
        cmdCollect = New Collection
        _InTransaction.Commit()
        _cn.Close()
        Return True

    End Function

#Region "Private Methods "

    Private Shared Sub PrepareCommand(ByRef ObjCmd As InfiniCommand)

        Dim _Strsql As String

        _strConnectionString = GetConnectionString()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        _Strsql = AccessQueries.GetQueryString(ObjCmd)

        'associate the connection with the command
        _cmd.Connection = _cn

        'set the command text SQL statement
        _cmd.CommandText = _Strsql

    End Sub

    Private Shared Sub PrepareCommand_Pro(ByRef ObjCmd As InfiniCommand)

        Dim _Strsql As String

        _strConnectionString = GetConnectionString_Pro()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        _Strsql = AccessQueries.GetQueryString(ObjCmd)

        'associate the connection with the command
        _cmd.Connection = _cn

        'set the command text SQL statement
        _cmd.CommandText = _Strsql

    End Sub

    Private Shared Sub LogoUpdate(ByRef ObjCmd As InfiniCommand)

        Dim param As New OleDbParameter
        With param
            .Value = CObj(ObjCmd.Parameter(0))
            .DbType = DbType.Binary
        End With
        _cmd.Connection = _cn
        With _cmd
            .CommandText = "Update TblCompany Set Logo=(?)"
            .CommandType = CommandType.Text
            .Parameters.Add(param)
        End With

    End Sub

#End Region

#Region "Synchronization"

    Public Overloads Shared Function ExecuteSynchronizeCollection(ByRef _cmdCollection As Collection) As DataSet

        Dim _Strsql As String, I As Integer = 0
        Dim tds As New DataSet
        Dim tcmd As New InfiniCommand("Nothing")

        _strConnectionString = GetConnectionString()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
                _InTransaction = _cn.BeginTransaction(IsolationLevel.ReadCommitted)
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        For Each tcmd In _cmdCollection

            I += 1
            tcmd = CType(_cmdCollection(I), InfiniCommand)

            _Strsql = AccessQueries.GetUnSynchronizeDataQueryString(tcmd)

            'associate the connection with the command
            _cmd.Connection = _cn

            'set the command text SQL statement
            _cmd.CommandText = _Strsql

            _cmd.Transaction = _InTransaction

            Try

                Dim Da As OleDbDataAdapter = New OleDbDataAdapter(CType(_cmd, OleDbCommand))
                Da.Fill(tds, tcmd.CommandName)

            Catch e As Exception
                _InTransaction.Rollback()
                Throw New Exception(" Transaction rolled back because of the following error. " & e.Message)
            End Try
        Next
        _cmdCollection = New Collection
        _InTransaction.Commit()

        Return tds

    End Function

#End Region

#End Region

    Public Overloads Shared Function ExecuteCollectionFor_Pro(ByRef _cmdCollection As Collection) As Boolean

        Dim _Strsql As String
        Dim i As Integer = 0
        Dim tempcmd As New InfiniCommand("Nothing")

        _strConnectionString = GetConnectionString_Pro()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand()
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
                _InTransaction = _cn.BeginTransaction(IsolationLevel.ReadCommitted)
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        For Each tempcmd In _cmdCollection

            i += 1
            tempcmd = CType(_cmdCollection(i), InfiniCommand)
            _Strsql = AccessQueries.GetQueryString(tempcmd)
            _cmd.CommandText = _Strsql
            _cmd.Connection = _cn
            _cmd.Transaction = _InTransaction

            Try
                _cmd.ExecuteNonQuery()
            Catch e As Exception
                _InTransaction.Rollback()
                Throw New Exception(" Transaction rolled back because of the following error. " & e.Message)
                Return False
            End Try
        Next
        _cmdCollection = New Collection()
        _InTransaction.Commit()
        _cn.Close()
        Return True

    End Function

    Public Shared Sub PrepareConnectionFor_pro(ByRef ObjCmd As InfiniCommand, ByVal _str As String)

        Dim _Strsql As String

        _cn = New OleDbConnection(_str)
        _cmd = New OleDbCommand()
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        _Strsql = AccessQueries.GetQueryString(ObjCmd)

        'associate the connection with the command
        _cmd.Connection = _cn

        'set the command text SQL statement
        _cmd.CommandText = _Strsql

    End Sub

#Region "BackUp Files In The Form Of Xml Sechma"

    Public Shared Function ExeComman4WriteXML(ByVal _cmdCollection As Collection) As DataSet
        Dim _Strsql As String, I As Integer = 0
        Dim tds As New DataSet()
        Dim tcmd As New InfiniCommand("Nothing")

        _strConnectionString = GetConnectionString()
        _cn = New OleDbConnection(_strConnectionString)
        _cmd = New OleDbCommand()
        Try
            If (_cn.State = ConnectionState.Closed) Then
                _cn.Open()
                _InTransaction = _cn.BeginTransaction(IsolationLevel.ReadCommitted)
            End If
        Catch sqle As OleDbException
            Throw New Exception(" : " & sqle.Message)
        Catch e As Exception
            Throw New Exception(" : " & e.Message)
        End Try

        For Each tcmd In _cmdCollection

            I += 1
            tcmd = CType(_cmdCollection(I), InfiniCommand)

            _Strsql = AccessQueries.GetQueryString(tcmd)

            'associate the connection with the command
            _cmd.Connection = _cn

            'set the command text SQL statement
            _cmd.CommandText = _Strsql

            _cmd.Transaction = _InTransaction

            Try

                Dim Da As OleDbDataAdapter = New OleDbDataAdapter(CType(_cmd, OleDbCommand))
                Da.Fill(tds, tcmd.CommandName)

            Catch e As Exception
                _InTransaction.Rollback()
                Throw New Exception(" Transaction rolled back because of the following error. " & e.Message)
            End Try
        Next
        _cmdCollection = New Collection()
        _InTransaction.Commit()

        Return tds

    End Function
#End Region
End Class