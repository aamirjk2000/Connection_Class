Imports System.Text
Imports System.Data.SqlClient

Public Module General

    Public MyMessages As New ArrayList
    Dim MyDefault_Values As Default_Values

    Public Enum DBConnection
        AppliedDB = 1
        AppliedUser = 2
        AppliedDelete = 3
        AppliedDocuments = 4
        AppliedMigration = 5
        AppliedDuplicate = 6
    End Enum
    Public Sub ShowMessages()
        Dim _ShowMessageForm As New TextVisualizer(MyMessages)
        _ShowMessageForm.ShowDialog()
    End Sub
    Public Sub ShowMessages(_Messages As ArrayList)
        Dim _ShowMessageForm As New TextVisualizer(_Messages)
        _ShowMessageForm.ShowDialog()
    End Sub
    Public Function Get_Table(_TableName As String, _Defaults As Default_Values) As DataTable

        'Get Table from Applied BMS Database.

        MyMessages.Add("Start Get_Table " & Now.ToLongDateString)
        MyMessages.Add("Server     | " & _Defaults.DBServer)
        MyMessages.Add("Login      | " & _Defaults.DBLogin)
        MyMessages.Add("PW Hash    | " & _Defaults.DBPWHash)
        MyMessages.Add("Table Name | " & _TableName)

        'Dim _TableName = ConnectionClass.Get_Table_Name(_TableID)
        Dim _Connection As SqlClient.SqlConnection = AppliedConnection(DBConnection.AppliedDB, MyDefault_Values)

        Dim _Command As New SqlClient.SqlCommand("SELECT * FROM " & _TableName & ";", _Connection)
        Dim _Adapter As New SqlClient.SqlDataAdapter(_Command)
        Dim _Dataset As New DataSet
        Dim _DataTable As New DataTable

        MyMessages.Add("Connection String" & _Connection.ConnectionString)
        MyMessages.Add("Connection Status" & _Connection.State.ToString)
        MyMessages.Add("Command Text" & _Command.CommandText)

        Try
            _Adapter.Fill(_Dataset, _TableName(3))
            _DataTable = _Dataset.Tables(0)

            MyMessages.Add("Table..  | " & _DataTable.TableName)
            MyMessages.Add("Records  | " & _DataTable.Rows.Count.ToString)

        Catch ex As Exception
            MyMessages.Add("******* ERROR *********")
            MyMessages.Add(ex.Message)
            ShowMessages(MyMessages)
        End Try

        ShowMessages(MyMessages)

        Return _DataTable

    End Function
    Public Function Get_TableView(_TableID As Integer, _Defaults As Default_Values)
        Dim _Datatable As DataTable = Get_Table(_TableID, _Defaults)
        Dim _TableView As New DataView
        _TableView.Table = _Datatable

        Return _TableView

    End Function
    Public ReadOnly Property MySQLConnection(_Default_Values As Default_Values) As String
        Get
            Return "Server=" & _Default_Values.DBServer & ";User ID=" & _Default_Values.DBLogin & "; Password=" & GetPassword(_Default_Values.DBPWHash, _Default_Values.DBPWWrapper)
        End Get
    End Property
    Public ReadOnly Property MySQLConnection_User(_Default_Values As Default_Values) As String
        Get
            Return "Server=" & _Default_Values.DBServer & ";User ID=" & _Default_Values.DBUser & "; Password=" & GetPassword(_Default_Values.DBPWHash, _Default_Values.DBPWWrapper)
        End Get
    End Property
    Public Function GetSetupConnection(_Defaults As Default_Values) As SqlClient.SqlConnection
        Dim _SetupConnection_String As String
        Dim _SetupConnection As New SqlClient.SqlConnection

        _SetupConnection_String = "Data Source=" & _Defaults.DBServer & ";Initial Catalog=" & _Defaults.DBDatabase & ";User ID=" & _Defaults.DBLogin & ";Password="
        _SetupConnection = New SqlClient.SqlConnection(_SetupConnection_String)
        _SetupConnection.Open()

        Return _SetupConnection


    End Function
    Public Sub Show_Connection_Record(_DataRow As DataRow)

        Dim ShowForm As New frmConnection_Record
        ShowForm.MyDataRow = _DataRow
        ShowForm.Show()

    End Sub

    Public Function SQLInsert(ByVal _Columns As DataColumnCollection) As String

        Dim _CommandString As StringBuilder = New StringBuilder()
        Dim _TableName As String = _Columns(0).Table.TableName
        Dim _LastColumn As String = _Columns(_Columns.Count - 1).ColumnName

        _CommandString.Append("INSERT INTO ")
        _CommandString.Append(_TableName)
        _CommandString.Append(" ( ")

        For Each _Column As DataColumn In _Columns
            Dim _ColumnName As String = _Column.ColumnName

            _CommandString.Append(String.Concat("[", _Column.ColumnName, "]"))

            If _ColumnName <> _LastColumn Then
                _CommandString.Append(",")
            Else
                _CommandString.Append(") ")
            End If
        Next

        _CommandString.Remove(_CommandString.ToString().Trim().Length - 1, 1)
        _CommandString.Append(") VALUES (")

        For Each _Column As DataColumn In _Columns
            Dim _ColumnName As String = _Column.ColumnName
            _CommandString.Append(String.Concat("@", _Column.ColumnName.Replace(" ", "")))

            If _ColumnName <> _LastColumn Then
                _CommandString.Append(",")
            Else
                _CommandString.Append(") ")
            End If
        Next

        Return _CommandString.ToString

    End Function
    Public Function SQLUpdate(ByVal _Columns As DataColumnCollection, _PrimaryKeyName As String) As String

        Dim _CommandString As StringBuilder = New StringBuilder()
        Dim _TableName As String = _Columns(0).Table.TableName
        Dim _LastColumn As String = _Columns(_Columns.Count - 1).ColumnName

        _CommandString.Append(String.Concat("UPDATE ", _TableName, " SET "))

        For Each _Column As DataColumn In _Columns
            Dim _ColumnName As String = _Column.ColumnName
            Dim _SkipColumns() As String = {_PrimaryKeyName, "Created"}

            _CommandString.Append(String.Concat("[", _Column.ColumnName, "]"))
            _CommandString.Append("=")
            _CommandString.Append(String.Concat("@", _Column.ColumnName.Replace(" ", "")))

            If _ColumnName <> _LastColumn Then
                _CommandString.Append(",")
            Else
                _CommandString.Append(String.Concat(" WHERE ", _PrimaryKeyName, "= @", _PrimaryKeyName, ";"))
            End If

        Next

        Return _CommandString.ToString()
    End Function
    Public Function SQLInsert(_DataRow As DataRow, _Connection As SqlConnection) As SqlCommand
        Dim _Command As New SqlCommand(SQLInsert(_DataRow.Table.Columns), _Connection)
        Dim _ParmaterName As String

        For Each _Column As DataColumn In _DataRow.Table.Columns
            If _Column Is Nothing Then
                Continue For
            End If

            _ParmaterName = String.Concat("@" & _Column.ColumnName.Replace(" ", ""))
            _Command.Parameters.AddWithValue(_ParmaterName, _DataRow(_Column.ColumnName))
        Next

        Return _Command
    End Function
    Public Function SQLUpdate(ByVal _DataRow As DataRow, _PrimaryKeyName As String, _Connection As SqlConnection) As SqlCommand

        Dim _Command As New SqlCommand(SQLUpdate(_DataRow.Table.Columns, _PrimaryKeyName), _Connection)
        Dim _ParmaterName As String

        For Each _Column As DataColumn In _DataRow.Table.Columns
            If _Column Is Nothing Then
                Continue For
            End If

            _ParmaterName = String.Concat("@" & _Column.ColumnName.Replace(" ", ""))
            _Command.Parameters.AddWithValue(_ParmaterName, _DataRow(_Column.ColumnName))
        Next

        Return _Command
    End Function

    Public Function ExcelToTable(_FileName As String, _Sheet As String) As DataTable
        Dim _Connection As OleDb.OleDbConnection
        Dim _DataSet As DataSet
        Dim _Command As OleDb.OleDbDataAdapter
        Dim _Table As DataTable

        ' Note : to executive this code, download the driver from below link.
        ' https://www.microsoft.com/en-pk/download/details.aspx?id=13255

        'Get Connection string from the following web site.
        'https://social.msdn.microsoft.com/Forums/vstudio/en-US/5f8297aa-4f45-4669-8adb-e7d7ac0e6f61/could-not-find-installable-isam-in-mycommandfilldtset-line?forum=vbgeneral
        _Connection = New System.Data.OleDb.OleDbConnection _
        ("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & _FileName & "'; Extended Properties=""Excel 12.0 Xml;HDR=YES""")
        _Command = New OleDb.OleDbDataAdapter("Select * from " & _Sheet, _Connection)
        _Command.TableMappings.Add("Table", "TestTable")
        _DataSet = New DataSet
        _Command.Fill(_DataSet, _Sheet)
        _Connection.Close()
        _Table = _DataSet.Tables(0)

        _Connection.Dispose()
        _Command.Dispose()
        _DataSet.Dispose()

        Return _Table
    End Function

End Module

