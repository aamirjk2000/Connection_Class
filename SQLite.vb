Imports System.Data.SQLite

Public Module SQLite

    Public Function SQLiteInsert(_DataRow As DataRow, _Connection As SQLiteConnection) As SQLiteCommand

        'Dim _Connection As Data.sql

        Dim _Columns As DataColumnCollection = _DataRow.Table.Columns                       ' Assign Columns to create SQLite Command.
        Dim _Command As New SqliteCommand(General.SQLInsert(_Columns), _Connection)
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


End Module
