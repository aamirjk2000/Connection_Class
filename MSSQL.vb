Imports System
Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Collections

Public Module MSSQL

    Public Function Get_DataTable(_TableName As String, _Connection As SqlConnection) As DataTable
        Dim _ResultTable As New DataTable
        Dim _CommandText As String = "SELECT * FROM " & _TableName & ";"
        Dim _Command As New SqlCommand(_CommandText, _Connection)
        Dim _Adapter As New SqlDataAdapter(_Command)
        Dim _Dataset As New DataSet

        _Adapter.Fill(_Dataset, _TableName)
        _ResultTable = _Dataset.Tables(0)

        Return _ResultTable
    End Function
    Public Function Seek(_ID As Integer, _TableName As String, _Connection As SqlConnection) As DataRow
        'Supply a Datarow search by Record ID .....   26-03-2019
        Dim _ResultRow As DataRow
        Dim _CommandString As String
        Dim _Command As SqlCommand
        Dim _Adapter As New SqlDataAdapter
        Dim _Dataset As New DataSet

        _CommandString = "SELECT * FROM " & _TableName & " WHERE ID=" & _ID & ";"
        _Command = _Connection.CreateCommand
        _Adapter.Fill(_Dataset)
        _ResultRow = _Dataset.Tables(0).Rows(0)

        Return _ResultRow

    End Function
    Public Function Seek(_Code As String, _TableName As String, _Connection As SqlConnection) As DataRow
        'Supply a Datarow search by Record Code .....   26-03-2019
        Dim _ResultRow As DataRow
        Dim _CommandString As String
        Dim _Command As SqlCommand
        Dim _Adapter As New SqlDataAdapter
        Dim _Dataset As New DataSet

        _CommandString = "SELECT * FROM " & _TableName & " WHERE Code=" & _Code & ";"
        _Command = _Connection.CreateCommand
        _Adapter.FillSchema(_Dataset, SchemaType.Mapped)
        _Adapter.Fill(_Dataset)
        _ResultRow = _Dataset.Tables(0).Rows(0)

        Return _ResultRow

    End Function
    Public Function SeekTitle(_Title As String, _TableName As String, _Connection As SqlConnection) As String
        'Supply a Datarow search by Record Column Exact Title .....   26-03-2019

        Dim _ResultTitle As String = ""
        Dim _CommandString As String = "SELECT * FROM " & _TableName & " WHERE Title ='" & _Title & "';"
        Dim _Command As New SqlCommand(_CommandString, _Connection)
        Dim _Reader As SqlDataReader = _Command.ExecuteReader
        _ResultTitle = _Reader.Item("Title")

        Return _ResultTitle

    End Function
    Public Function SeekTitles(_Titles As String, _TableName As String, _Connection As SqlConnection) As String
        'Supply a Datarow search by Record Column Exact Title .....   26-03-2019

        Dim _ResultTitle As String = ""
        Dim _CommandString As String = "SELECT * FROM " & _TableName & " WHERE Title Like '%" & _Titles & "%';"
        Dim _Command As New SqlCommand(_CommandString, _Connection)
        Dim _Reader As SqlDataReader = _Command.ExecuteReader
        _ResultTitle = _Reader.Item("Title")

        Return _ResultTitle

    End Function
    Public Function Get_Title(_ID As Integer, Seek_Column As String, Get_Column As String, _TableName As String, _Connection As SqlConnection) As String
        'Supply a Datarow search by Record ID .....   26-03-2019
        Dim _ResultTitle As String
        Dim _CommandString As String
        Dim _Command As SqlCommand
        Dim _Adapter As New SqlDataAdapter
        Dim _Dataset As New DataSet
        Dim _DataRow As DataRow

        Try
            _CommandString = "SELECT * FROM " & _TableName & " WHERE " & Seek_Column & "=" & _ID & ";"
            _Command = _Connection.CreateCommand
            _Command.CommandText = _CommandString
            _Adapter = New SqlDataAdapter(_Command)
            _Adapter.Fill(_Dataset)
            _DataRow = _Dataset.Tables(0).Rows(0)
            _ResultTitle = _DataRow(Get_Column)
        Catch ex As Exception
            _ResultTitle = "ERROR " & ex.Message
        End Try

        Return _ResultTitle

    End Function

    Public Function Get_Procedure(_Report As Report_Parameters) As DataTable
        ' 27-08-2019.
        ' Supplie Data Table get from SQL Procedure  Updated Verion of (_ProcedureName As String, _SQLParameters As Dictionary(Of String, String), _Connection As SqlConnection)

        Dim _ResultTable As New DataTable
        Dim _Messages As New ArrayList

        _Messages.Add(_Report.SQLCommand.CommandType.ToString)
        _Messages.Add(_Report.SQLCommand.CommandText)
        _Messages.Add(_Report.SQLConnection.ConnectionString)

        _Report.SQLCommand.CommandText = _Report.SQLProcedure
        _Report.SQLCommand.CommandType = CommandType.StoredProcedure

        _Report.SQLAdapter = New SqlDataAdapter(_Report.SQLCommand)
        _Report.SQLDataSet = New DataSet
        _Report.SQLAdapter.Fill(_Report.SQLDataSet)                  ' Fill DataSet.
        _ResultTable = _Report.SQLDataSet.Tables(0)           ' Get DataTable from Dataset.

        If _Report.ShowMessages Then
            MsgBox(_ResultTable.Rows.Count & " Records Found.")
        End If

        If _ResultTable Is Nothing Then
            _Messages.Add("EXEC | ConnectionClassDLL.Get_Procedure(ReportParameter)")
            _Messages.Add("DataTable is Nothing in the result of SQL Proceidre execution.")

            _ResultTable = New DataTable
        End If

        Return _ResultTable

    End Function

    Public Function Get_Procedure(_SQLCommand As SqlCommand) As DataTable
        ' 27-08-2019.  02-Sep-=2019
        ' Supplie Data Table get from SQL Procedure  Updated Verion of (_ProcedureName As String, _SQLParameters As Dictionary(Of String, String), _Connection As SqlConnection)

        _SQLCommand.CommandType = CommandType.StoredProcedure

        Dim _ResultTable As DataTable
        Dim _Messages As New ArrayList
        Dim _SQLAdapter As New SqlDataAdapter(_SQLCommand)
        Dim _SQLDataSet As New DataSet

        _Messages.Add(_SQLCommand.CommandType.ToString)
        _Messages.Add(_SQLCommand.CommandText)
        _Messages.Add(_SQLCommand.Connection.ConnectionString)

        _SQLAdapter.Fill(_SQLDataSet)                  ' Fill DataSet.

        If _SQLDataSet.Tables(0) IsNot Nothing Then
            _ResultTable = _SQLDataSet.Tables(0)           ' Get DataTable from Dataset.
        Else
            _Messages.Add("EXEC | ConnectionClassDLL.Get_Procedure(ReportParameter)")
            _Messages.Add("DataTable is Nothing in the result of SQL Proceidre execution.")
            _ResultTable = New DataTable
        End If

        Return _ResultTable

        'Completed ... 02-Sep-19.

    End Function

End Module
