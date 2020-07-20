Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
'
' Create Setup
' This class will Create a Physical File in Filder as specifiy by User for Setup Database.
' that file will be use in Getconnections Class to getting Database Connection for Applied BMS.
' Date ....... 10-Mar-2017 by Muhammad Aamir Jahangir
' Date ....... 23-Jul-2018 updated some codes and fixed error on connection string.
' Date ....... 22-Feb-2019 updated with new throry of class...... Major changes made.

Public Class CreateSetup

    Public Property HasError As Boolean
    Public Property SysMessage As New ArrayList

    Dim Default_Value_Class As Default_Values

    Dim Instance_Not_Exist As String = "doesn't exist!"                 ' SQL Instance does not exist message contain this vlaue
    Dim Instance_Exist As String = "State:"                             ' SQL Instance not exist message contain this vlaue

    Dim ProcessLevel As Integer = 0                                     ' Process bar value for processing
    Dim ProcessLevelMax As Integer = 0                                  ' Process bar value for processing MAx Value

    Public Sub New()
        'DB_Connection = DefaultValues.SQLConnection

        SysMessage.Clear()                     ' Clear logs
        SysMessage.Add("Start at " & Date.Now.ToLongDateString & " " & Date.Now.ToLongTimeString)
        Default_Value_Class = New Default_Values

        HasError = False

        SysMessage.Add("Load parameters from default value class.")
    End Sub

    Public Sub New(_Default_Values As Default_Values)
        'DB_Connection = DefaultValues.SQLConnection
        SysMessage.Clear()                     ' Clear logs
        SysMessage.Add("Start at " & Date.Now.ToLongDateString & " " & Date.Now.ToLongTimeString)
        HasError = False
        Default_Value_Class = _Default_Values
        SysMessage.Add("Load parameters from user define values.")
    End Sub

    '=============================================================== SQL Local DB Commands
    Public Shared Function InstanceCreate(DB_Instance) As String

        Dim OutPutResult As String
        Dim SetupDBProcess As New Process
        SetupDBProcess.StartInfo.FileName = "SQLLocalDB"
        SetupDBProcess.StartInfo.Arguments = "create " & DB_Instance
        SetupDBProcess.StartInfo.CreateNoWindow = True
        SetupDBProcess.StartInfo.RedirectStandardOutput = True
        SetupDBProcess.StartInfo.UseShellExecute = False
        SetupDBProcess.Start()
        OutPutResult = SetupDBProcess.StandardOutput.ReadToEnd
        Return OutPutResult

    End Function
    Public Shared Function InstanceDelete(DB_Instance) As String

        Dim OutPutResult As String
        Dim SetupDBProcess As New Process
        SetupDBProcess.StartInfo.FileName = "SQLLocalDB"
        SetupDBProcess.StartInfo.Arguments = "delete " & DB_Instance
        SetupDBProcess.StartInfo.CreateNoWindow = True
        SetupDBProcess.StartInfo.RedirectStandardOutput = True
        SetupDBProcess.StartInfo.UseShellExecute = False
        SetupDBProcess.Start()
        OutPutResult = SetupDBProcess.StandardOutput.ReadToEnd
        Return OutPutResult

    End Function
    Public Shared Function InstanceInfo(DB_Instance) As String

        Dim OutPutResult As String
        Dim SetupDBProcess As New Process
        SetupDBProcess.StartInfo.FileName = "SQLLocalDB"
        SetupDBProcess.StartInfo.Arguments = "info " & DB_Instance
        SetupDBProcess.StartInfo.CreateNoWindow = True
        SetupDBProcess.StartInfo.RedirectStandardOutput = True
        SetupDBProcess.StartInfo.RedirectStandardError = True
        SetupDBProcess.StartInfo.UseShellExecute = False
        SetupDBProcess.Start()
        OutPutResult = SetupDBProcess.StandardOutput.ReadToEnd
        Return OutPutResult

    End Function
    Public Shared Function InstanceStart(DB_Instance) As String

        Dim OutPutResult As String
        Dim SetupDBProcess As New Process
        SetupDBProcess.StartInfo.FileName = "SQLLocalDB"
        SetupDBProcess.StartInfo.Arguments = "Start " & DB_Instance
        SetupDBProcess.StartInfo.CreateNoWindow = True
        SetupDBProcess.StartInfo.RedirectStandardOutput = True
        SetupDBProcess.StartInfo.UseShellExecute = False
        SetupDBProcess.Start()
        OutPutResult = SetupDBProcess.StandardOutput.ReadToEnd
        Return OutPutResult

    End Function
    Public Shared Function InstanceStop(DB_Instance) As String

        Dim OutPutResult As String
        Dim SetupDBProcess As New Process
        SetupDBProcess.StartInfo.FileName = "SQLLocalDB"
        SetupDBProcess.StartInfo.Arguments = "Stop " & DB_Instance
        SetupDBProcess.StartInfo.CreateNoWindow = True
        SetupDBProcess.StartInfo.RedirectStandardOutput = True
        SetupDBProcess.StartInfo.UseShellExecute = False
        SetupDBProcess.Start()
        OutPutResult = SetupDBProcess.StandardOutput.ReadToEnd
        Return OutPutResult

    End Function
    '---------------------------------------------------------------------------------------

    Public Sub CreateSetupDB(_Default_Values As Default_Values)

        Default_Value_Class = _Default_Values

        Dim BarClass As New MyProgress
        BarClass.MyProgressBar.Maximum = 5
        BarClass.MyProgressBar.Value = 0

        SysMessage.Add("=================================================Default Values")
        'SysMessage.Add("Connection String   |" & SQLConnection(Default_Value_Class))
        SysMessage.Add("Database Server     |" & Default_Value_Class.DBServer)
        SysMessage.Add("Database Instance   |" & Default_Value_Class.SQLInstance)
        SysMessage.Add("DataBase Engine     |" & Default_Value_Class.DBEngine)
        SysMessage.Add("Database File Path  |" & Default_Value_Class.DBPath)
        SysMessage.Add("DataBase File       |" & Default_Value_Class.DBSetupFile)
        SysMessage.Add("DataBase Log File   |" & Default_Value_Class.DBSetupLog)
        SysMessage.Add("DataBase Name       |" & Default_Value_Class.DBDatabase)
        SysMessage.Add("DataBase Schema     |" & Default_Value_Class.DBSchema)
        SysMessage.Add("DataBase Conn.Table |" & Default_Value_Class.DBSetupTable)
        SysMessage.Add("DataBase Login ID   |" & Default_Value_Class.DBLogin)
        SysMessage.Add("DataBase User       |" & Default_Value_Class.DBUser)
        SysMessage.Add("DataBase PW Hash    |" & Default_Value_Class.DBPWHash)
        SysMessage.Add("DataBase PW Wrapper |" & Default_Value_Class.DBPWWrapper)
        SysMessage.Add("================================================================")

        BarClass.MyLabel.Text = "Values have been initialized."
        BarClass.Show()
        BarClass.Refresh()

        If HasError = False Then
            MsgBox("Start....Create Setup DB   " & Default_Value_Class.SQLInstance)
            BarClass.MyProgressBar.Value = 1
            BarClass.Refresh()
            CreateInstance()                    ' Create Database Instanse

        End If

        If HasError = False Then
            MsgBox("Start.   CreateDBFile   " & Default_Value_Class.SQLInstance)
            CreateDBFile()                      ' Create Databsse in Local SQL Database Service
            BarClass.MyProgressBar.Value = 2
            BarClass.Refresh()

        End If

        If HasError = False Then
            MsgBox("Start.   CreateLogin " & Default_Value_Class.DBUser)
            CreateLogin()                       ' Create Login ID
            BarClass.MyProgressBar.Value = 3
            BarClass.Refresh()
        End If

        If HasError = False Then
            MsgBox("Start.   CreateUser")
            CreateUser()                        ' Create Database user 
            BarClass.MyProgressBar.Value = 4
            BarClass.Refresh()
        End If

        If HasError = False Then
            MsgBox("Start.   Table")
            CreateTable()                       ' Create a Table for Applied System Database Records.
            BarClass.MyProgressBar.Value = 5
            BarClass.Refresh()
        End If



        If HasError Then
            SysMessage.Add("Procedure got error.")
        Else
            SysMessage.Add("")
            SysMessage.Add("====================================")
            SysMessage.Add("Setup Database successfully created.")
            SysMessage.Add("====================================")
        End If

        ShowMessages(SysMessage)               ' Show Messages 
    End Sub
    Public Sub CreateInstance()

        SysMessage.Add("========================================================")
        SysMessage.Add("Start CreateInstance()" & Now.ToString)

        Dim _Result As String
        Dim _Message As String = ""

        _Result = InstanceInfo(Default_Value_Class.SQLInstance)

        Try
            If _Result.Contains(Instance_Not_Exist) Then
                InstanceCreate(Default_Value_Class.SQLInstance)
                _Message = Default_Value_Class.SQLInstance & " | Created."
            ElseIf _Result.Contains(Instance_Exist) Then
                _Message = Default_Value_Class.SQLInstance & " already exist."
            End If

        Catch ex As Exception
            HasError = True
            SysMessage.Add("ERROR : " & ex.Message)
        End Try

        SysMessage.Add(_Message)
        SysMessage.Add("End CreateInstance()" & Now.ToString)
        SysMessage.Add("")

    End Sub
    Public Sub CreateDBFile()

        'Create Local SQL Database in Databsse Service

        SysMessage.Add("======================================================== ")
        SysMessage.Add("Start CreateDBFile()" & Now.ToString)


        Dim _Connection As New SqlConnection(First_Setup_Connection)
        Dim _Command As SqlCommand = _Connection.CreateCommand

        SysMessage.Add("CONNECTION | " & _Connection.ConnectionString)
        Try
            _Connection.Open()
        Catch ex As Exception
            HasError = True
            SysMessage.Add(ex.Message)
            Exit Sub
        End Try

        Dim _DBFile As String = Path.Combine(Default_Value_Class.DBPath, Default_Value_Class.DBSetupFile)
        Dim _DBLog As String = Path.ChangeExtension(_DBFile, "ldf")
        Dim TempFile As String = Now.ToString("dd-mm-yyyy") & "_" _
                               & Path.GetFileNameWithoutExtension(Path.GetTempFileName)

        SysMessage.Add("_DBFile | " & _DBFile)
        SysMessage.Add("_DBLog | " & _DBLog)

        Try
            If Not Directory.Exists(Default_Value_Class.DBPath) Then
                Directory.CreateDirectory(Default_Value_Class.DBPath)
                SysMessage.Add("Created Folder :" & Default_Value_Class.DBPath)
            End If

            If File.Exists(_DBFile) Then
                SysMessage.Add("Exist:" & _DBFile)
                SysMessage.Add("Target:" & Path.GetFileNameWithoutExtension(TempFile) & ".mdf")
                My.Computer.FileSystem.RenameFile(_DBFile, TempFile & ".mdf")
            End If

            If File.Exists(_DBLog) Then
                SysMessage.Add("Exist:" & _DBLog)
                SysMessage.Add("Target:" & Path.GetFileNameWithoutExtension(TempFile) & ".ldf")
                My.Computer.FileSystem.RenameFile(_DBLog, TempFile & ".ldf")
            End If

        Catch ex As Exception
            HasError = True
            SysMessage.Add("Error | " & ex.Message)

        End Try

        If Not HasError Then
            _Command.CommandText = "CREATE DATABASE [" & Default_Value_Class.DBDatabase & "] " _
                             & "ON PRIMARY ( NAME=" & Default_Value_Class.DBDatabase & "_DB, " _
                             & "FILENAME='" & _DBFile & "') " _
                             & "LOG ON ( NAME=" & Default_Value_Class.DBDatabase & "_LOG, " _
                             & "FILENAME='" & _DBLog & "')" _
                             & ";"

            SysMessage.Add(_Command.CommandText)
            Try
                _Command.ExecuteNonQuery()
                SysMessage.Add("Database [" & Default_Value_Class.DBDatabase & "] has been created")
                HasError = False

            Catch ex As Exception
                HasError = True
                SysMessage.Add(_Command.CommandText)
                SysMessage.Add(ex.Message)
                SysMessage.Add("HasError = True")
            End Try

            _Connection = Nothing
            _Command = Nothing


        End If
    End Sub
    Public Function CreateLogin() As String
        SysMessage.Add(" ")
        SysMessage.Add("========================================================")
        SysMessage.Add("Start CreateLogin() " & Now.ToString)

        Dim _Connection As New SqlConnection(First_Setup_Connection)
        Dim _Command As New SqlCommand

        SysMessage.Add("Connection String | " & First_Setup_Connection)

        Try
            _Connection.Open()
            SysMessage.Add(MySqlConnection(Default_Value_Class))
            SysMessage.Add("Connection Opened")
        Catch ex As Exception
            HasError = True
            SysMessage.Add(ex.Message)
            SysMessage.Add("HasError=True")
            Return ex.Message
        End Try

        If HasError = False Then

            Try
                '---------------------
                _Command = _Connection.CreateCommand
                _Command.CommandText = "CREATE LOGIN " & Default_Value_Class.DBLogin & " " _
                                 & "WITH PASSWORD='" & DecryptPassword(Default_Value_Class.DBPWHash, Default_Value_Class.DBPWWrapper) & "' "

                SysMessage.Add("CONNECTION " & _Command.CommandText)
                _Command.ExecuteNonQuery()
                SysMessage.Add(Default_Value_Class.DBLogin & " Login Created.")

                '---------------------
                _Command.CommandText = "ALTER SERVER ROLE [bulkadmin] ADD MEMBER " & Default_Value_Class.DBLogin & ";"
                SysMessage.Add(_Command.CommandText)
                _Command.ExecuteNonQuery()
                SysMessage.Add(Default_Value_Class.DBLogin & " Grant to bulkAdmin")

                '---------------------
                _Command.CommandText = "ALTER SERVER ROLE [sysadmin] ADD MEMBER " & Default_Value_Class.DBLogin & ";"
                SysMessage.Add(_Command.CommandText)
                _Command.ExecuteNonQuery()
                SysMessage.Add(Default_Value_Class.DBLogin & " Grant to sysAdmin")

                '---------------------
                Dim _DBUser As String = Default_Value_Class.DBUser

            Catch ex As Exception
                HasError = True
                SysMessage.Add("CONNECTION [ERROR]" & _Command.CommandText)
                SysMessage.Add(ex.Message)
                SysMessage.Add("HasError = True")
                Return ex.Message
            End Try
        End If

        _Connection = Nothing
        _Command = Nothing


        SysMessage.Add(IIf(HasError, "", "No Error found."))
        SysMessage.Add("End CreateLogin() " & Now.ToString)
        SysMessage.Add("")
        Return ""

    End Function
    Public Function CreateUser() As String

        SysMessage.Add(" ")
        SysMessage.Add("========================================================")
        SysMessage.Add("Start CreateUser() " & Now.ToString)

        Dim _Connection As New SqlConnection(MySQLConnection_User(Default_Value_Class) & " ;Initial Catalog=" & Default_Value_Class.DBDatabase & ";")
        Dim _Command As New SqlCommand


        SysMessage.Add("Connection [User] |" & _Connection.ConnectionString)

        Try
            _Connection.Open()
            _Command = _Connection.CreateCommand
            _Command.CommandText = "CREATE USER " & Default_Value_Class.DBUser & " FROM LOGIN " & Default_Value_Class.DBLogin
            SysMessage.Add("Command    |" & _Command.CommandText)
            SysMessage.Add("ExecuteNonQuery() Result | " & _Command.ExecuteNonQuery().ToString)            ' Execute SQL Command.
            SysMessage.Add(Default_Value_Class.DBUser & " User Created for Database " & Default_Value_Class.DBDatabase)

            '---------------------
            _Command.CommandText = "ALTER ROLE [db_owner] ADD MEMBER [" & Default_Value_Class.DBUser & "];"
            SysMessage.Add(_Command.CommandText)
            _Command.ExecuteNonQuery()
            SysMessage.Add(Default_Value_Class.DBUser & " role granted [db_owner].")

        Catch ex As Exception
            HasError = True
            SysMessage.Add(ex.Message)
            SysMessage.Add("HasError = True")
            Return ex.Message
        End Try

        _Connection = Nothing
        _Command = Nothing

        SysMessage.Add("End CreateUser() " & Now.ToString)
        SysMessage.Add("")
        Return IIf(HasError, "CreateUser() has Error", "No Error Found.")

    End Function
    Public Function CreateTable() As String

        SysMessage.Add("======================================================== ")
        SysMessage.Add("Start CreateTable() " & Now.ToString)

        Dim _Connection As New SqlConnection(MySQLConnection_User(Default_Value_Class) & " ;Initial Catalog=" & Default_Value_Class.DBDatabase & ";")
        Dim _Command As New SqlCommand

        SysMessage.Add("CONNECTION | " & _Connection.ConnectionString)

        Try
            _Connection.Open()
            SysMessage.Add(MySQLConnection(Default_Value_Class))
            SysMessage.Add("Connection Opened")
        Catch ex As Exception
            HasError = True
            SysMessage.Add(ex.Message)
            SysMessage.Add("HasError=True")
            Return ex.Message
        End Try

        If HasError = False Then
            _Command = _Connection.CreateCommand
            _Command.CommandText = "CREATE SCHEMA " & Default_Value_Class.DBSchema & " AUTHORIZATION " & Default_Value_Class.DBUser & ";"

            Try
                _Command.ExecuteNonQuery()
                SysMessage.Add(Default_Value_Class.DBSchema & ": Schema Created")
            Catch ex As Exception
                HasError = True
                SysMessage.Add(_Command.CommandText)
                SysMessage.Add(ex.Message)
                SysMessage.Add("HasError = True")
                Return ex.Message
            End Try
        End If

        Dim _Table As String = Default_Value_Class.DBSchema & "." & Default_Value_Class.DBSetupTable

        If HasError = False Then
            _Command = _Connection.CreateCommand
            _Command.CommandText = " CREATE TABLE " & _Table & " 
	        (
	        [ID]                    [int]			PRIMARY KEY,
	        [Code]					[char](20)		NOT NULL,
	        [Title]					[nchar](60)		NOT NULL,
	        [Provider]				[nvarchar](max) NULL,
	        [Driver]				[nvarchar](max) NULL,
	        [Data Source]			[nvarchar](max) NULL,
	        [Server]				[nvarchar](max) NULL,
	        [Address]				[nvarchar](max) NULL,
	        [Addr]					[nvarchar](max) NULL,
	        [NetWork Address]		[nvarchar](max) NULL,
	        [Database]				[nvarchar](max) NULL,
	        [Initial Catalog]       [nvarchar](max) NULL,
	        [Integrated Security]	[nvarchar](max) NULL,
	        [Trusted_Connection]	[nvarchar](max) NULL,
	        [DataSchema]			[nvarchar](max) NULL,
	        [Timeout]				[nvarchar](max) NULL,
            [Connect Timeout]	    [nvarchar](max) NULL,
	        [Connection Timeout]	[nvarchar](max) NULL,
	        [Connection Lifetime]   [nvarchar](max) NULL,
            [Load Balance Timeout]  [nvarchar](max) NULL,
	        [Context Connection]    [nvarchar](max) NULL,
	        [Current Language]      [nvarchar](max) NULL,
	        [Language]              [nvarchar](max) NULL,
	        [User ID]				[nvarchar](max) NULL,
	        [Password]				[nvarchar](max) NULL,
	        [PWD]					[nvarchar](max) NULL,
	        [Encrypt]				[nvarchar](max) NULL,
	        [AttachDBFilename]		[nvarchar](max) NULL,
	        [Extended Properties]	[nvarchar](max) NULL,
	        [Initial File Name]		[nvarchar](max) NULL,
	        [User Instance]			[nvarchar](max) NULL,
	        [Failover Partner]		[nvarchar](max) NULL,
	        [MultiSubnet Failover]	[nvarchar](max) NULL,
	        [Persist Security Info] [nvarchar](max) NULL,
	        [Authentication]        [nvarchar](max) NULL,
	        [Workstation ID]		[nvarchar](max) NULL,
            [App]                   [nvarchar](max) NULL,
            [Application Name]      [nvarchar](max) NULL,
            [ApplicationIntent]     [nvarchar](max) NULL,
            [Column Encryption Setting] [nvarchar](max) NULL,
	        )"

            Try
                _Command.ExecuteNonQuery()
                SysMessage.Add("[Applied].[AppliedSetupDBTable] TableCreated.")
            Catch ex As Exception
                HasError = True
                SysMessage.Add(ex.Message)
                SysMessage.Add("HasError = True")
                Return ex.Message
            End Try
        End If

        Return IIf(HasError, "CreateTable() has error", "CreateTable()....OK ")
    End Function

    Friend ReadOnly Property First_Setup_Connection() As String
        Get
            Return "Data Source=" & Default_Value_Class.DBServer & ";Initial Catalog=Master;Integrated Security=SSPI;"
        End Get
    End Property

    Private Function MyTextMessage() As String
        Dim _Text As String = ""
        Dim _ArrayList As Array = SysMessage.ToArray

        For Each _Line As String In _ArrayList
            _Text += _Line + Environment.NewLine
        Next
        Return _Text
    End Function

End Class

Public Class Utilities

    ' This Class Create / Delete Local DB Instance through CMD.exe

    Property UtilityMessage As String

    Public Sub New()
        UtilityMessage = "New"

    End Sub
    Shared Sub Run_Command(command As String, arguments As String, permanent As Boolean)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(permanent = True, "/K", "/C") + " " + command + " " + arguments
        pi.FileName = "cmd.exe"
        p.StartInfo = pi
        p.Start()
    End Sub
    Public Shared Function GetServerList() As String
        Dim Instances As New ArrayList
        Dim Server As String = String.Empty

        Dim instance As SqlDataSourceEnumerator = SqlDataSourceEnumerator.Instance
        Dim table As System.Data.DataTable = instance.GetDataSources()

        For Each row As System.Data.DataRow In table.Rows
            Server = String.Empty
            Server = row("ServerName")
            If row("InstanceName").ToString.Length > 0 Then
                Server = Server & "\" & row("InstanceName")
            End If
            MsgBox(Server & " : Server Instance")
            Instances.Add(Server)
        Next
        Return Instances.ToString
    End Function

End Class
