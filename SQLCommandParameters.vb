Imports System.Data.SqlClient
Imports System.Text

Public Class SQLCommandParameters
    ' This Class provides paramters for SQL Command
    Public Sub New()
        MyMessage = "Start...." & Now.ToLongDateString
    End Sub

    Public Property RecordID As Integer
    Public Property PKeyColumn As String
    Public Property SQLRow As DataRow
    Public Property SQLTable As DataTable
    Public Property SQLTableName As String
    Public Property SQLConnection As SqlConnection
    Public Property MyMessage As String = ""
    Public Property Skip_PK As Boolean = True
    Public Property ShowLog As Boolean = False

    Public ReadOnly Property InsertCommand As SqlCommand
        Get
            Return SQLInsertCommand()
        End Get
    End Property

    Public ReadOnly Property UpdateCommand As SqlCommand
        Get
            Return SQLUpdateCommand()
        End Get
    End Property

    Public ReadOnly Property DeleteCommand As SqlCommand
        Get
            Return SQLDeleteCommand()
        End Get
    End Property


    Private Function SQLInsertCommand() As SqlCommand

        Dim _SysMessages As New ArrayList
        _SysMessages.Add("Start...Insert Command. " & Now.ToLongDateString)

        Dim _TotalColumns As Integer = SQLTable.Columns.Count
        Dim _SQLCommand As SqlCommand = SQLConnection.CreateCommand         ' Create Command for Connection
        Dim _Parameters(_TotalColumns) As SqlParameter             ' Array SQL Parameters 
        Dim _CommandBuild As New StringBuilder


        'Dim _EmployeeTable As String = GetTableName(Tables.HREmployees)     ' Get Name for skip title column of employee

        _CommandBuild.Append("INSERT INTO ")
        _CommandBuild.Append(SQLTableName)
        _CommandBuild.Append(" (")


        '-------------------------------------------------------------------- SKIP COLUMN END
        For Each _Column As DataColumn In SQLTable.Columns


            'If Skip_PK Then                                                    ' Include primary Key column 
            'If _Column.ColumnName.ToUpper = PKeyColumn.ToUpper Then        ' Skip ID Column
            'Continue For
            'End If
            'End If
            _CommandBuild.Append("[")
            _CommandBuild.Append(_Column.ColumnName)
            _CommandBuild.Append("]")

            If _Column.ColumnName.Equals(SQLTable.Columns(_TotalColumns - 1).ColumnName) Then             ' Last Column action.
                _CommandBuild.Append(") VALUES (")
            Else
                _CommandBuild.Append(", ")
            End If
        Next

        For Each _Column As DataColumn In SQLTable.Columns

            _CommandBuild.Append("@")
            _CommandBuild.Append(_Column.ColumnName)

            If _Column.ColumnName.Equals(SQLTable.Columns(_TotalColumns - 1).ColumnName) Then
                _CommandBuild.Append(");")
            Else
                _CommandBuild.Append(", ")
            End If
        Next

        Dim _Index As Integer = 0

        For Each _Column As DataColumn In SQLTable.Columns

            Dim _ColumnParameter As String = "@" & _Column.ColumnName
            Dim _ColumnValue As Object = SQLRow(_Column.ColumnName)

            _Parameters(_Index) = _SQLCommand.Parameters.AddWithValue(_ColumnParameter, _ColumnValue)

            _SysMessages.Add("Parameter " & _ColumnParameter.ToString)
            _SysMessages.Add("Value " & _Parameters(_Index).Value)
            _SysMessages.Add("---------------------------------------------")

            _Index += 1
        Next

        _SysMessages.Add(_CommandBuild.ToString)
        _SQLCommand.CommandText = _CommandBuild.ToString
        _SysMessages.Add("END .....Insert Command.")
        _SysMessages.Add("")

        If ShowLog Then
            ShowMessages(_SysMessages)
        End If

        Return _SQLCommand
    End Function

    Private Function SQLUpdateCommand() As SqlCommand

        Dim _SysMessages As New ArrayList
        _SysMessages.Add("Start...Update Command. " & Now.ToLongDateString)

        Dim _TotalColumns As Integer = SQLTable.Columns.Count
        Dim _SQLCommand As SqlCommand = SQLConnection.CreateCommand         ' Create Command for Connection
        Dim _Parameters(_TotalColumns) As SqlParameter             ' Array SQL Parameters 
        Dim _CommandBuild As New StringBuilder


        'Dim _EmployeeTable As String = GetTableName(Tables.HREmployees)     ' Get Name for skip title column of employee

        _CommandBuild.Append("UPDATE ")
        _CommandBuild.Append(SQLTableName)
        _CommandBuild.Append(" SET ")

        '-------------------------------------------------------------------- SKIP COLUMN END
        For Each _Column As DataColumn In SQLTable.Columns

            If Skip_PK Then                                                    ' Include primary Key column 
                If _Column.ColumnName.ToUpper = PKeyColumn.ToUpper Then        ' Skip ID Column
                    Continue For
                End If
            End If

            _CommandBuild.Append("[")
            _CommandBuild.Append(_Column.ColumnName)
            _CommandBuild.Append("] = @")
            _CommandBuild.Append(_Column.ColumnName)

            If _Column.ColumnName.Equals(SQLTable.Columns(_TotalColumns - 1).ColumnName) Then             ' Last Column action.
                _CommandBuild.Append("")
            Else
                _CommandBuild.Append(", ")
            End If
        Next

        _CommandBuild.Append(" WHERE " & PKeyColumn & "=@" & PKeyColumn)

        Dim _Index As Integer = 0

        For Each _Column As DataColumn In SQLTable.Columns

            Dim _ColumnParameter As String = "@" & _Column.ColumnName
            Dim _ColumnValue As Object = SQLRow(_Column.ColumnName)

            _Parameters(_Index) = _SQLCommand.Parameters.AddWithValue(_ColumnParameter, _ColumnValue)

            _SysMessages.Add("Parameter " & _ColumnParameter.ToString)
            _SysMessages.Add("Value " & _Parameters(_Index).Value)
            _SysMessages.Add("---------------------------------------------")

            _Index += 1

        Next

        _SysMessages.Add(_CommandBuild.ToString)
        _SQLCommand.CommandText = _CommandBuild.ToString
        _SysMessages.Add("END .....Update Command.")
        _SysMessages.Add("")

        If ShowLog Then
            ShowMessages(_SysMessages)
        End If

        Return _SQLCommand
    End Function

    Private Function SQLDeleteCommand() As SqlCommand

        Dim _SysMessages As New ArrayList
        _SysMessages.Add("Start...Delete Command. " & Now.ToLongDateString)

        Dim _TotalColumns As Integer = SQLTable.Columns.Count
        Dim _SQLCommand As SqlCommand = SQLConnection.CreateCommand         ' Create Command for Connection
        Dim _Parameters(_TotalColumns) As SqlParameter             ' Array SQL Parameters 
        Dim _CommandBuild As New StringBuilder


        'Dim _EmployeeTable As String = GetTableName(Tables.HREmployees)     ' Get Name for skip title column of employee

        _CommandBuild.Append("DELETE FROM ")
        _CommandBuild.Append(SQLTableName)
        _CommandBuild.Append(" WHERE " & PKeyColumn & "=@" & PKeyColumn)

        Dim _ColumnParameter As String = "@" & PKeyColumn
        Dim _ColumnValue As Object = SQLRow(PKeyColumn)

        _Parameters(0) = _SQLCommand.Parameters.AddWithValue(_ColumnParameter, _ColumnValue)


        _SysMessages.Add("Parameter " & _ColumnParameter.ToString)
        _SysMessages.Add("Value " & _Parameters(0).Value)
        _SysMessages.Add("---------------------------------------------")


        _SysMessages.Add(_CommandBuild.ToString)
        _SQLCommand.CommandText = _CommandBuild.ToString
        _SysMessages.Add("END .....Delete Command.")
        _SysMessages.Add("")

        If ShowLog Then
            ShowMessages(_SysMessages)
        End If

        Return _SQLCommand
    End Function

    Public Function Execute_SQLCommand(_SQLCommand As SqlCommand) As Integer
        Return _SQLCommand.ExecuteNonQuery()
    End Function


End Class
