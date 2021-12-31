Imports System.Data.SqlClient
Imports System.Data.SQLite

Public Module DataConnection

    Private Property MyDefault_Values As New Default_Values
    Private Property MySysMessage As New ArrayList
    Private Property MyRowID As Integer = 0

    'AppliedConnection is main Propert which will provides the Data connection to Apploied BMS System

    Public ReadOnly Property AppliedConnection(_RowID As Integer, _Default_Values As Default_Values) As SqlConnection
        Get
            MyDefault_Values = _Default_Values

            MyMessages.Add("Start ..... Applied Connection at " & Now.ToLongDateString)
            Dim _Row As DataRow = Get_Setup_Row(_RowID, MyDefault_Values)
            Dim _AppliedConnection As New SqlConnection(GetAppliedConnectionString(_Row))
            Dim _IsError As Boolean = False

            Try
                _AppliedConnection.Open()
                MyMessages.Add(_AppliedConnection.ConnectionString)
            Catch ex As Exception
                MyMessages.Add("********* ERROR **********")
                MyMessages.Add(ex.Message)
                MyMessages.Add(ex.Source)
                MyMessages.Add(ex.InnerException)
                MyMessages.Add(ex.Source)
                MyMessages.Add(ex.TargetSite)
                MyMessages.Add(ex.StackTrace)
                ShowMessages(MyMessages)
            End Try

            Return _AppliedConnection
        End Get
    End Property
    Friend Function GetAppliedConnectionString(_Row) As String

        ' This Function convert Datarow into Connection String  


        Dim BuildString As New SqlConnectionStringBuilder
        Dim Value As New Object
        Dim _HasError As Boolean = False

        MyMessages.Add("")
        MyMessages.Add("Start GetDBConnectionString().....")

        '------------------------------------------------------------------------------ Data Source
        Try

            If Not IsDBNull(_Row("Data Source")) Then
                MyMessages.Add("Data Source | " & _Row("Data Source").ToString)
                BuildString.DataSource = _Row("Data Source")

            ElseIf Not IsDBNull(_Row("Server")) Then
                MyMessages.Add("Server | " & _Row("Server").ToString)
                BuildString.Add("Server", _Row("Server"))

            ElseIf Not IsDBNull(_Row("Address")) Then
                MyMessages.Add("Address | " & _Row("Address").ToString)
                BuildString.Add("Address", _Row("Address"))

            ElseIf Not IsDBNull(_Row("Addr")) Then
                MyMessages.Add("Addr | " & _Row("Addr").ToString)
                BuildString.Add("Addr", _Row("Addr"))

            ElseIf Not IsDBNull(_Row("Network Address")) Then
                MyMessages.Add("Network Address |" & _Row("Network Address").ToString)
                BuildString.Add("Network Address", _Row("Network Address"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Data Source / Server / Address / Addr / Network Address")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Data Base 
        Try

            If Not IsDBNull(_Row("Database")) Then
                MyMessages.Add("Database | " & _Row("Database").ToString)
                BuildString.Add("Database", _Row("Database"))

            ElseIf Not IsDBNull(_Row("Initial Catalog")) Then
                MyMessages.Add("Initial Catalog |" & _Row("Initial Catalog").ToString)
                BuildString.InitialCatalog = _Row("Initial Catalog")
            End If
        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Initial Catalog")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- User ID
        Try
            If Not IsDBNull(_Row("User ID")) Then
                MyMessages.Add("User ID | " & _Row("User ID").ToString)
                BuildString.UserID = _Row("User ID")

            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- User ID")
            MyMessages.Add(ex.Message)
        End Try
        '------------------------------------------------------------------------------- Password
        Try

            If Not IsDBNull(_Row("Password")) Then
                MyMessages.Add("Password | " & _Row("Password").ToString)

                Try
                    BuildString.Password = PW.GetPassword(_Row("Password"), MyDefault_Values.DBPWWrapper)
                Catch ex As Exception
                    _HasError = True
                    MyMessages.Add("Error ----------- Password")
                    MyMessages.Add(ex.Message)
                End Try

                If _HasError Then
                    BuildString.Password = ""
                End If

            ElseIf Not IsDBNull(_Row("PWD")) Then
                MyMessages.Add("PWD | " & _Row("PWD").ToString)

                Try
                    BuildString.Add("PWD", PW.GetPassword(_Row("PWD"), MyDefault_Values.DBPWWrapper))
                Catch ex As Exception
                    _HasError = True
                    MyMessages.Add("Error ----------- PWD")
                    MyMessages.Add(ex.Message)

                    BuildString.Password = ""
                End Try

                If _HasError Then
                    BuildString.Add("PWD", "")
                End If

            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Password / PWD")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Encrypt
        Try
            If Not IsDBNull(_Row("Encrypt")) Then
                MyMessages.Add("Encrypt | " & _Row("Encrypt").ToString)

                Try
                    BuildString.Encrypt = Convert.ToBoolean(_Row("Encrypt"))
                Catch ex As Exception
                    MyMessages.Add("Encrypt can not be convert")
                    MyMessages.Add(ex.Message)
                    BuildString.Encrypt = Nothing
                End Try
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Encrypt")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Authentication
        Try
            If Not IsDBNull(_Row("Authentication")) Then
                MyMessages.Add("Authentication | " & _Row("Authentication").ToString)
                Try
                    BuildString.Add("Authentication", Convert.ToBoolean(_Row("Authentication")))
                Catch ex As Exception
                    _HasError = True
                    MyMessages.Add("Authentication can not be convert")
                    MyMessages.Add(ex.Message)
                Finally
                End Try
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Authentication")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Connect Time
        Try
            If Not IsDBNull(_Row("Connect Timeout")) Then
                MyMessages.Add("Connect Timeout | " & _Row("Connect Timeout").ToString)
                BuildString.ConnectTimeout = Convert.ToInt32(_Row("Connect Timeout"))

            ElseIf Not IsDBNull(_Row("Timeout")) Then
                MyMessages.Add("Timeout | " & _Row("Timeout").ToString)
                BuildString.Add("TimeOut", Convert.ToInt32(_Row("TimeOut")))

            ElseIf Not IsDBNull(_Row("Connection Timeout")) Then
                MyMessages.Add("Connection Timeout | " & _Row("Connection Timeout").ToString)
                BuildString.Add("Connection Timeout", Convert.ToInt32(_Row("Connection Timeout")))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Connect Timeout")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Connection Lifetime
        Try
            If Not IsDBNull(_Row("Connection Lifetime")) Then
                MyMessages.Add("Connection Lifetime | " & _Row("Connection Lifetime").ToString)
                BuildString.Add("Connection Lifetime", Convert.ToBoolean(_Row("Connection Lifetime")))

            ElseIf Not IsDBNull(_Row("Load Balance Timeout")) Then
                MyMessages.Add("Load Balance Timeout | " & _Row("Load Balance Timeout").ToString)
                BuildString.Add("Load Balance Timeout", Convert.ToBoolean(_Row("Load Balance Timeout")))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Connection Lifetime")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Provider
        'Try


        'If Not IsDBNull(_Row("Provider")) Then
        'MyMessages.Add("Provider | " & _Row("Provider").ToString)
        'BuildString.Add("Provider", _Row("Provider"))
        'End If
        'Catch ex As Exception
        '_HasError = True
        'MyMessages.Add("Error ----------- Provider")
        'MyMessages.Add(ex.Message)
        'End Try
        '------------------------------------------------------------------------------- Driver

        'Try
        'If Not IsDBNull(_Row("Driver")) Then
        'MyMessages.Add("Driver | " & _Row("Driver").ToString)
        'BuildString.Add("Driver", _Row("Driver"))
        'End If
        'Catch ex As Exception
        '_HasError = True
        'MyMessages.Add("Error ----------- Driver")
        'MyMessages.Add(ex.Message)
        'End Try

        '------------------------------------------------------------------------------- Integrated Security
        Try
            If Not IsDBNull(_Row("Integrated Security")) Then
                MyMessages.Add("Integrated Security | " & _Row("Integrated Security").ToString)
                BuildString.IntegratedSecurity = Convert.ToBoolean(_Row("Integrated Security"))

            ElseIf Not IsDBNull(_Row("Trusted_Connection")) Then
                MyMessages.Add("Trusted_Connection | " & _Row("Trusted_Connection").ToString)
                BuildString.Add("Trusted_Connection", Convert.ToBoolean(_Row("Trusted_Connection")))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Integrated Security")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Data Schema
        'Try
        'If Not IsDBNull(_Row("DataSchema")) Then
        'MyMessages.Add("DataSchema | " & _Row("DataSchema").ToString)
        'BuildString.Add("DataSchema", _Row("DataSchema"))
        'End If

        'Catch ex As Exception

        'End Try

        '------------------------------------------------------------------------------- Context Connection
        Try
            If Not IsDBNull(_Row("Context Connection")) Then
                MyMessages.Add("Context Connection | " & _Row("Context Connection").ToString)
                BuildString.ContextConnection = Convert.ToBoolean(_Row("Context Connection"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Context Connection")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Current Language
        Try
            If Not IsDBNull(_Row("Current Language")) Then
                MyMessages.Add("Current Language | " & _Row("Current Language").ToString)
                BuildString.CurrentLanguage = _Row("Current Language")
            ElseIf Not IsDBNull(_Row("Language")) Then
                MyMessages.Add("Language | " & _Row("Language").ToString)
                BuildString.Add("Language", _Row("Language"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Current Language / Language")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Attached DB file Name
        Try
            If Not IsDBNull(_Row("AttachDBFileName")) Then
                MyMessages.Add("AttachDBFileName | " & _Row("AttachDBFileName").ToString)
                BuildString.AttachDBFilename = _Row("AttachDBFileName")

            ElseIf Not IsDBNull(_Row("Extended Properties")) Then
                MyMessages.Add("Extended Properties | " & _Row("Extended Properties").ToString)
                BuildString.Add("Extended Properties", _Row("Extended Properties"))

            ElseIf Not IsDBNull(_Row("Initial File Name")) Then
                MyMessages.Add("Initial File Name |" & _Row("Initial File Name").ToString)
                BuildString.Add("Initial File Name", _Row("Initial File Name"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- AttachDBFileName / Extended Properties / Initial File Name")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- User Instance
        Try
            If Not IsDBNull(_Row("User Instance")) Then
                MyMessages.Add("User Instance |" & _Row("User Instance").ToString)
                BuildString.UserInstance = Convert.ToBoolean(_Row("User Instance"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- User Instance")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Fail Over partner
        Try

            If Not IsDBNull(_Row("Failover Partner")) Then
                MyMessages.Add("Failover Partner | " & _Row("Failover Partner").ToString)
                BuildString.FailoverPartner = _Row("Failover Partner")
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Failover Partner")
            MyMessages.Add(ex.Message)
        End Try


        '------------------------------------------------------------------------------- MultiSubnet Failover 
        Try
            If Not IsDBNull(_Row("MultiSubnet Failover")) Then
                MyMessages.Add("MultiSubnet Failover | " & _Row("MultiSubnet Failover").ToString)
                BuildString.MultiSubnetFailover = Convert.ToBoolean(_Row("MultiSubnet Failover"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- MultiSubnet Failover")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Persist Security Infor
        Try
            If Not IsDBNull(_Row("Persist Security Info")) Then
                MyMessages.Add("Persist Security Info | " & _Row("Persist Security Info").ToString)
                BuildString.PersistSecurityInfo = Convert.ToBoolean(_Row("Persist Security Info"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Persist Security Info")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Persist Security Infor
        Try
            If Not IsDBNull(_Row("Workstation ID")) Then
                MyMessages.Add("Workstation ID | " & _Row("Workstation ID").ToString)
                BuildString.WorkstationID = _Row("WorkStation ID")
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Workstation ID")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Column Encryption Setting
        Try
            If Not IsDBNull(_Row("Column Encryption Setting")) Then
                MyMessages.Add("Column Encryption Setting | " & _Row("Column Encryption Setting").ToString)
                BuildString.Add("Column Encryption Setting", _Row("Column Encryption Setting"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Column Encryption Setting")
            MyMessages.Add("Has Error = " & _HasError.ToString)
            MyMessages.Add(ex.Message)

        End Try

        '------------------------------------------------------------------------------- Application
        Try
            If Not IsDBNull(_Row("Application Name")) Then
                MyMessages.Add("Application Name | " & _Row("Application Name").ToString)
                BuildString.ApplicationName = _Row("Application Name")
            ElseIf Not IsDBNull(_Row("App")) Then
                MyMessages.Add("App | " & _Row("App").ToString)
                BuildString.Add("App", _Row("App"))
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- Application Name")
            MyMessages.Add(ex.Message)
        End Try

        '------------------------------------------------------------------------------- Application Intent
        Try
            If Not IsDBNull(_Row("ApplicationIntent")) Then
                MyMessages.Add("ApplicationIntent | " & _Row("ApplicationIntent").ToString)
                BuildString.ApplicationIntent = _Row("ApplicationIntent")
            End If

        Catch ex As Exception
            _HasError = True
            MyMessages.Add("Error ----------- ApplicationIntent")
            MyMessages.Add("Has Error = " & _HasError.ToString)
            MyMessages.Add(ex.Message)
        End Try

        Dim _String As String

        If _HasError Then
            _String = "Connection String has error."
            MyMessages.Add("Error = True")
        Else
            _String = BuildString.ToString
            MyMessages.Add("Error = False")
        End If

        MyMessages.Add("----------------------------------------")
        MyMessages.Add(_String.Replace(BuildString.Password, "********"))
        MyMessages.Add("----------------------------------------")

        MyMessages.Add("End GetDBConnectionString......")


        If _HasError Then
            ShowMessages()
        End If

        Return _String

    End Function
    Public ReadOnly Property Get_Setup_Row(_RowID As Integer, _Default_Values As Default_Values) As DataRow
        Get
            ' This Property provide the row of Database Connection for Applied Database Connection String.

            MyDefault_Values = _Default_Values

            MySysMessage.Add("Start .... Get_Setup_Row " & Now.ToLongDateString)
            MySysMessage.Add("Server  | " & MyDefault_Values.DBServer)
            MySysMessage.Add("Login   | " & MyDefault_Values.DBLogin)
            MySysMessage.Add("PW Hash | " & MyDefault_Values.DBPWHash)
            MySysMessage.Add("Row ID  | " & _RowID)
            MySysMessage.Add("")
            MySysMessage.Add("My Server  | " & MyDefault_Values.DBServer)
            MySysMessage.Add("My Login   | " & MyDefault_Values.DBLogin)
            MySysMessage.Add("My PW Hash | " & MyDefault_Values.DBPWHash)
            MySysMessage.Add("My Row ID  | " & _RowID)

            Dim _SetupRow As DataRow                                                    ' Result    
            Dim _DataTable As DataTable = Get_Setup_Table()  ' Get Local SQL DB Table
            'Dim _DataTable As DataTable = Get_Setup_Table(MyDefault_Values)  ' Get Local SQL DB Table
            Dim _TableView As New DataView                                              ' Set Data Table View

            _TableView.Table = _DataTable
            _TableView.RowFilter = "ID=" & _RowID
            _SetupRow = _DataTable.NewRow

            ' Assign result of this function

            If _TableView.Count = 0 Then
                MsgBox("No Database Connection record found....", MsgBoxStyle.Information, "ALERT 1")

            ElseIf _TableView.Count = 1 Then
                _SetupRow = _TableView.Item(0).Row
            ElseIf _TableView.Count > 1 Then
                MsgBox(_TableView.Count.ToString & " Database Connection records found....", MsgBoxStyle.Information, "ALERT 2")

            End If

            Return _SetupRow

        End Get
    End Property

    Public ReadOnly Property Get_Setup_Table() As DataTable
        Get
            Dim _SQLConnection As SQLiteConnection      ' Establish Setup Database Connection
            Dim _SQLCommand As New SQLiteCommand
            Dim _Adapter As New SQLiteDataAdapter
            Dim _DataSet As New DataSet
            Dim _DataTable As New DataTable
            Dim _TableName As String = MyDefault_Values.SetupTableName                  ' Schema plus Table Name
            Dim _FilePath As String = "E:\AMCORP_ERP_APP\Setup_Database\AMCORP_ERP.db"


            MySysMessage.Add("Table | " & _TableName)

            Try
                _SQLCommand.CommandText = "SELECT * FROM " & _TableName & ";"
                _Adapter = New SQLiteDataAdapter(_SQLCommand.CommandText, SQLite.SQLiteConnection(_FilePath))
                _Adapter.FillSchema(_DataSet, SchemaType.Mapped, _TableName)
                _Adapter.Fill(_DataSet, MyDefault_Values.DBSetupTable)
                _DataTable = _DataSet.Tables(MyDefault_Values.DBSetupTable)
            Catch ex As SqlException
                MsgBox(MyDefault_Values.DBSetupTable & " TABLE NOT FOUND")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MySysMessage.Add("SQL Command    | " & _SQLCommand.CommandText)
            MySysMessage.Add("SQL Connection | " & _SQLConnection.State.ToString)
            MySysMessage.Add("Total Records  | " & _DataTable.Rows.Count.ToString)

            Return _DataTable               ' Return Setup Database records entries table.
        End Get
    End Property




    Public ReadOnly Property Get_Setup_Table(_Default_Values As Default_Values) As DataTable
        Get
            MyDefault_Values = _Default_Values

            MySysMessage.Clear()
            MySysMessage.Add("Start.....Get_Setup_Table " & Now.ToLongDateString)
            MySysMessage.Add("")
            MySysMessage.Add("Server  | " & MyDefault_Values.DBServer)
            MySysMessage.Add("LoginID | " & MyDefault_Values.DBLogin)
            MySysMessage.Add("PWHash  | " & MyDefault_Values.DBPWHash)

            Dim _SQLConnection As SqlConnection = MyDefault_Values.SetupConnection      ' Establish Setup Database Connection
            Dim _SQLCommand As New SqlCommand
            Dim _Adapter As New SqlDataAdapter
            Dim _DataSet As New DataSet
            Dim _DataTable As New DataTable
            Dim _TableName As String = MyDefault_Values.SetupTableName                  ' Schema plus Table Name

            MySysMessage.Add("Table | " & _TableName)

            Try
                _SQLCommand.CommandText = "SELECT * FROM " & _TableName & ";"
                _Adapter = New SqlDataAdapter(_SQLCommand.CommandText, _SQLConnection)
                _Adapter.FillSchema(_DataSet, SchemaType.Mapped, _TableName)
                _Adapter.Fill(_DataSet, MyDefault_Values.DBSetupTable)
                _DataTable = _DataSet.Tables(MyDefault_Values.DBSetupTable)
            Catch ex As SqlException
                MsgBox(MyDefault_Values.DBSetupTable & " TABLE NOT FOUND")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MySysMessage.Add("SQL Command    | " & _SQLCommand.CommandText)
            MySysMessage.Add("SQL Connection | " & _SQLConnection.State.ToString)
            MySysMessage.Add("Total Records  | " & _DataTable.Rows.Count.ToString)

            Return _DataTable               ' Return Setup Database records entries table.
        End Get
    End Property






End Module
