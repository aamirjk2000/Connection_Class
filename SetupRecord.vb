Imports System.Data.SqlClient
Imports System.Text

Public Class SetupRecord

    ' Date  .......... 27-Mar-2017 / 22-Mar-2019
    ' Author ......... Muhammad Aamir Jahangir

    Public Property MySysMessage As New ArrayList
    Public Property MySetupConnection As SqlConnection
    Public Property MyDataRecord As DataRecord
    Public Property MyDataTable As DataTable
    Public Property MyDataRow As DataRow
    Public Property MyDefaults As Default_Values
    Public Property HasError As Boolean
    Public Property DBTable As DataTable

    Sub New(_Defaults As Default_Values)
        MySysMessage.Add("Start ....DataRecordSave() " & Now.ToLongDateString)
        MySysMessage.Add("")
        MyDefaults = _Defaults

    End Sub

    '====================================================================== Get / Save / Delete Record / Get Table

    Public Function SaveRecord() As Boolean

        MySysMessage.Add("Start Save Record..." & Now.ToLongDateString)

        Dim _DataRow As DataRow = Get_Setup_Row(MyDataRecord.ID, MyDefaults)     ' Assign DataRow for Connection
        Dim _RecordNew As Boolean = IIf(_DataRow.IsNull("ID"), True, False)       ' Check this record is new and exist
        Dim _SetupConnection As SqlConnection
        Dim _SQLCommand As New SqlCommand
        Dim _TableName = MyDefaults.DBSchema & "." & MyDefaults.DBSetupTable    ' Get Table Name from DataRow 
        'Dim _TableName = _DataRow.Table.TableName                               ' Get Table Name from DataRow 

        MySysMessage.Add("Table Name " & _TableName)
        MySysMessage.Add(" ")

        _SetupConnection = MyDefaults.SetupConnection                           ' Setup Connection established.
        _SQLCommand = _SetupConnection.CreateCommand                            ' Create Command for Datarow save.

        MySysMessage.Add("Setup Connection Status is " & _SetupConnection.State.ToString)
        MySysMessage.Add("Connection String " & _SetupConnection.ConnectionString)

        MySysMessage.Add("Connection Data Row is " & IIf(_RecordNew, "New", "Exist"))

        If _RecordNew Then                                                      ' If DataRow is new
            _DataRow = _DataRow.Table.NewRow                                    ' Assign a new Datarow for Connection save.
            MySysMessage.Add("Connection Data Row is New Row")
        Else
            MySysMessage.Add("Connection Data Row is Exist")
        End If

        'File Values into Datarow from Datarow class.
        _DataRow("ID") = MyDataRecord.ID
        _DataRow("Code") = MyDataRecord.Code
        _DataRow("Title") = MyDataRecord.Title
        _DataRow("Provider") = MyDataRecord.Provider
        _DataRow("Driver") = MyDataRecord.Driver
        _DataRow("Data Source") = MyDataRecord.Data_Source
        _DataRow("Server") = MyDataRecord.Server
        _DataRow("Address") = MyDataRecord.Address
        _DataRow("Addr") = MyDataRecord.Addr
        _DataRow("Network Address") = MyDataRecord.NetWork_Address
        _DataRow("Database") = MyDataRecord.Database
        _DataRow("Initial Catalog") = MyDataRecord.Initial_Catalog
        _DataRow("Integrated Security") = MyDataRecord.Integrated_Security
        _DataRow("Trusted_Connection") = MyDataRecord.Trusted_Connection
        _DataRow("DataSchema") = MyDataRecord.DataSchema
        _DataRow("Timeout") = MyDataRecord.Timeout
        _DataRow("Connection Timeout") = MyDataRecord.Connection_Timeout
        _DataRow("Connection Lifetime") = MyDataRecord.Connection_Lifetime
        _DataRow("Context Connection") = MyDataRecord.Context_Connection
        _DataRow("Current Language") = MyDataRecord.Current_Language
        _DataRow("User ID") = MyDataRecord.User_ID
        _DataRow("Password") = MyDataRecord.Password
        _DataRow("PWD") = MyDataRecord.PWD
        _DataRow("Encrypt") = MyDataRecord.Encrypt
        _DataRow("AttachDBFilename") = MyDataRecord.AttachDBFilename
        _DataRow("User Instance") = MyDataRecord.User_Instance
        _DataRow("Failover Partner") = MyDataRecord.Failover_Partner
        _DataRow("Persist Security Info") = MyDataRecord.Persist_Security_Info
        _DataRow("Workstation ID") = MyDataRecord.Workstation_ID

        'Assign SQL Command Parameters and fill vlaues.
        _SQLCommand.Parameters.AddWithValue("@ID", _DataRow("ID"))
        _SQLCommand.Parameters.AddWithValue("@Code", _DataRow("Code"))
        _SQLCommand.Parameters.AddWithValue("@Title", _DataRow("Title"))
        _SQLCommand.Parameters.AddWithValue("@Provider", _DataRow("Provider"))
        _SQLCommand.Parameters.AddWithValue("@Driver", _DataRow("Driver"))
        _SQLCommand.Parameters.AddWithValue("@DataSource", _DataRow("Data Source"))
        _SQLCommand.Parameters.AddWithValue("@Server", _DataRow("Server"))
        _SQLCommand.Parameters.AddWithValue("@Address", _DataRow("Address"))
        _SQLCommand.Parameters.AddWithValue("@Addr", _DataRow("Addr"))
        _SQLCommand.Parameters.AddWithValue("@NetworkAddress", _DataRow("Network Address"))
        _SQLCommand.Parameters.AddWithValue("@Database", _DataRow("Database"))
        _SQLCommand.Parameters.AddWithValue("@InitialCatalog", _DataRow("Initial Catalog"))
        _SQLCommand.Parameters.AddWithValue("@IntegratedSecurity", _DataRow("Integrated Security"))
        _SQLCommand.Parameters.AddWithValue("@Trusted_Connection", _DataRow("Trusted_Connection"))
        _SQLCommand.Parameters.AddWithValue("@DataSchema", _DataRow("DataSchema"))
        _SQLCommand.Parameters.AddWithValue("@Timeout", _DataRow("Timeout"))
        _SQLCommand.Parameters.AddWithValue("@ConnectionTimeout", _DataRow("Connection Timeout"))
        _SQLCommand.Parameters.AddWithValue("@ConnectionLifetime", _DataRow("Connection Lifetime"))
        _SQLCommand.Parameters.AddWithValue("@ContextConnection", _DataRow("Context Connection"))
        _SQLCommand.Parameters.AddWithValue("@CurrentLanguage", _DataRow("Current Language"))
        _SQLCommand.Parameters.AddWithValue("@UserID", _DataRow("User ID"))
        _SQLCommand.Parameters.AddWithValue("@Password", _DataRow("Password"))
        _SQLCommand.Parameters.AddWithValue("@PWD", _DataRow("PWD"))
        _SQLCommand.Parameters.AddWithValue("@Encrypt", _DataRow("Encrypt"))
        _SQLCommand.Parameters.AddWithValue("@AttachDBFilename", _DataRow("AttachDBFilename"))
        _SQLCommand.Parameters.AddWithValue("@UserInstance", _DataRow("User Instance"))
        _SQLCommand.Parameters.AddWithValue("@FailoverPartner", _DataRow("Failover Partner"))
        _SQLCommand.Parameters.AddWithValue("@PersistSecurityInfo", _DataRow("Persist Security Info"))
        _SQLCommand.Parameters.AddWithValue("@WorkstationID", _DataRow("Workstation ID"))

        'Assign Command Insert if Datarow is new owthersie Assign Command Update 
        If _RecordNew Then
            MySysMessage.Add("SQL Query Insert.....")
            _SQLCommand.CommandText = "INSERT INTO " & _TableName _
                                & " ( " _
                                & "[ID], [Code], [Title], [Provider], [Driver], [Data Source], [Server], [Address], [Addr], " _
                                & "[Network Address], [Database], [Initial Catalog], [Integrated Security], [Trusted_Connection], " _
                                & "[DataSchema], [Timeout], [Connection Timeout], [Connection Lifetime], [Context Connection], " _
                                & "[Current Language], [User ID], [Password], [PWD], [Encrypt], [AttachDBFilename], [User Instance], " _
                                & "[Failover Partner], [Persist Security Info], [Workstation ID] " _
                                & ") VALUES " _
                                & "( " _
                                & "@ID, @Code, @Title, @Provider, @Driver, @DataSource, @Server, @Address, @Addr, " _
                                & "@NetworkAddress, @Database, @InitialCatalog, @IntegratedSecurity, @Trusted_Connection, " _
                                & "@DataSchema, @Timeout, @ConnectionTimeout, @ConnectionLifetime, @ContextConnection, " _
                                & "@CurrentLanguage, @UserID, @Password, @PWD, @Encrypt, @AttachDBFilename, @UserInstance, " _
                                & "@FailoverPartner, @PersistSecurityInfo, @WorkstationID " _
                                & ");"
        Else
            MySysMessage.Add("SQL Query Update.....")
            _SQLCommand.CommandText = "UPDATE " & _TableName _
                                & " SET " _
                                & "[Code]=@Code, " _
                                & "[Title]=@Title, " _
                                & "[Provider]=@Provider, " _
                                & "[Driver]=@Driver, " _
                                & "[Data Source]=@DataSource, " _
                                & "[Server]=@Server, " _
                                & "[Address]=@Address, " _
                                & "[Addr]=@Addr, " _
                                & "[Network Address]=@NetworkAddress, " _
                                & "[Database]=@Database, " _
                                & "[Initial Catalog]=@InitialCatalog, " _
                                & "[Integrated Security]=@IntegratedSecurity, " _
                                & "[Trusted_Connection]=@Trusted_Connection, " _
                                & "[DataSchema]=@DataSchema, " _
                                & "[Timeout]=@Timeout, " _
                                & "[Connection Timeout]=@ConnectionTimeout, " _
                                & "[Connection Lifetime]=@ConnectionLifetime, " _
                                & "[Context Connection]=@ContextConnection, " _
                                & "[Current Language]=@CurrentLanguage, " _
                                & "[User ID]=@UserID, " _
                                & "[Password]=@Password, " _
                                & "[PWD]=@PWD, " _
                                & "[Encrypt]=@Encrypt, " _
                                & "[AttachDBFilename]=@AttachDBFilename, " _
                                & "[User Instance]=@UserInstance, " _
                                & "[Failover Partner]=@FailoverPartner, " _
                                & "[Persist Security Info]=@PersistSecurityInfo, " _
                                & "[Workstation ID]=@WorkstationID " _
                                & "WHERE ID=@ID;"
        End If

        MySysMessage.Add("SQL Command Text " & _SQLCommand.CommandText)
        MySysMessage.Add("SQL Command Paramaters are " & _SQLCommand.Parameters.ToString)

        Try
            _SQLCommand.ExecuteNonQuery()                       ' Exeucutive SQl Command
            MySysMessage.Add("Execute SQL Command")
            HasError = True
        Catch ex As Exception
            MySysMessage.Add("Error.....Execute SQL Command")
            MySysMessage.Add(ex.Message)
            HasError = False
        End Try

        Dim _Result As Boolean = False

        If HasError Then
            MySysMessage.Add("Record Save.")
            _Result = True
        Else
            MySysMessage.Add("Record NOT Save.")
        End If
        MySysMessage.Add("End SaveRecordRow() Row ID=" & _DataRow("ID") & " | " & Now.ToString)
        MySysMessage.Add("")


        MsgBox("End Connection Record Saved.")
        ShowMessages(MySysMessage)

        Return _Result

    End Function

    Public Function DeleteRecord(_ID As Integer) As String
        MySysMessage.Add("DeleteRecord() Row ID=" & _ID & " | " & Now.ToString)

        Dim _DataRow As DataRow = Get_Setup_Row(_ID, MyDefaults)     ' Assign DataRow for Connection
        Dim _TotRows As Integer = 0

        If _DataRow Is Nothing Then
            MsgBox("Connection Data Row not found.")
            Return "No record found."
        End If

        Dim _SetupConnection As SqlConnection
        Dim _SQLCommand As New SqlCommand
        Dim _TableName = _DataRow.Table.TableName

        _SetupConnection = GetSetupConnection(MyDefaults)
        _SQLCommand = _SetupConnection.CreateCommand
        _SQLCommand.CommandText = "DELETE FROM " & _DataRow.Table.TableName & " WHERE ID=@ID"
        _SQLCommand.Parameters.AddWithValue("@ID", _ID)

        HasError = False

        Try
            _TotRows = _SQLCommand.ExecuteNonQuery()
            MySysMessage.Add("Record ID " & _ID & " Found....")
            MySysMessage.Add("Execute SQL Query. Total Rows effected " & _TotRows)
            MySysMessage.Add(_SQLCommand.CommandText)
        Catch ex As Exception
            MySysMessage.Add("Error.....")
            MySysMessage.Add(ex.Message)
            HasError = True
        End Try

        MySysMessage.Add("End Delete... " & Now.ToString)
        Return _TotRows & " Rows deleted of ID " & _ID

    End Function


End Class
