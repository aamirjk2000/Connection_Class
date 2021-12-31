Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.SQLite

Public Module SQLite


    Public Function ExcelConnection(_ExcelFileName As String) As OleDbConnection

        Dim ExcelVersion As String = "Excel 12.0;HDR=YES;IMEX=1;"

        Dim oledbConn As OleDbConnection
        oledbConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _ExcelFileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';")
        Return oledbConn


    End Function

    Public Function SQLiteConnection(_FilePath As String) As SQLiteConnection

        Dim sqlite_conn As SQLiteConnection

        sqlite_conn = New SQLiteConnection("Data Source=" + _FilePath + " Version = 3; New = True; Compress = True; ")

        Try

        Catch ex As Exception

        End Try

        Try

            sqlite_conn.Open()

        Catch ex As Exception
            MsgBox("DataBase Connection is not being established \r" + ex.Message, "ERROR")

        End Try

        If sqlite_conn.State = ConnectionState.Open Then
            Return sqlite_conn
        Else
            Return New SQLiteConnection
        End If



    End Function



    Public Function SQLiteInsert(_DataRow As DataRow, _Connection As SQLiteConnection) As SQLiteCommand

        'Dim _Connection As Data.sql

        Dim _Columns As DataColumnCollection = _DataRow.Table.Columns                       ' Assign Columns to create SQLite Command.
        Dim _Command As New SQLiteCommand(General.SQLInsert(_Columns), _Connection)
        Dim _ParmaterName As String

        For Each _Column As DataColumn In _Columns
            If _Column Is Nothing Then
                Continue For
            End If

            _ParmaterName = String.Concat("@" & _Column.ColumnName.Replace(" ", ""))
            _Command.Parameters.AddWithValue(_ParmaterName, _DataRow(_Column.ColumnName))
        Next

        Return _Command
    End Function

    Public Function SQLiteUpdate(ByVal _DataRow As DataRow, _PrimaryKeyName As String, _Connection As SQLiteConnection) As SQLiteCommand

        Dim _Command As New SQLiteCommand(General.SQLUpdate(_DataRow.Table.Columns, _PrimaryKeyName), _Connection)
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

    Public Function SQLiteDelete(ByVal _DataRow As DataRow, _Connection As SQLiteConnection) As SQLiteCommand

        Dim _Command As New SQLiteCommand("", _Connection)
        Dim _DataTableName As String = _DataRow.Table().TableName
        Dim _ID As Long = Convert.ToInt64(_DataRow("ID"))

        _Command.Parameters.AddWithValue("@ID", _ID)
        _Command.CommandText = "DELETE FROM " + _DataTableName + " WHERE ID=@ID"

        Return _Command

    End Function


End Module
