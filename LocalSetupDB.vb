Imports System.Data.SqlClient

Public Class LocalSetupDB

    ' Class for establish a Local Database Connection for Getting Database String for BMS System.
    Dim MyDefaults As Default_Values
    Dim LocalConnection As SqlConnection
    Dim MySysMessage As New ArrayList
    Dim MyErrorClass As New ErrorClass

    Dim DB_Instance As String
    Dim DB_Engine As String
    Dim DB_Path As String
    Dim DB_File As String
    Dim DB_Log As String
    Dim DB_Database As String
    Dim DB_Login As String
    Dim DB_PWHash As String

    Public Sub New()

        MySysMessage.Add("Start....." & Now)
        MySysMessage.Add("")

        DB_Engine = ""
        DB_Path = ""
        DB_File = ""
        DB_Log = ""
        DB_Database = ""
        DB_Login = ""
        DB_PWHash = ""

    End Sub

    Sub New(_Defaults As Default_Values)
        MyDefaults = _Defaults
    End Sub

    ' New Object created with empty values

    Public Function SetupConnection() As SqlConnection
        ' This will provide Local Database Connection (Setup DB Connection).

        MySysMessage = New ArrayList
        MySysMessage.Add("")
        MySysMessage.Add("Start.....LocalDBConnected()")
        MyErrorClass = New ErrorClass

        ' Generate Local DB String
        Dim _SQLConnection As New SqlConnection(MyDefaults.SetupConnectionString() & GetPassword(MyDefaults.DBPWHash) & ";")     'Establish a Local Connection.

        MySysMessage.Add(MyDefaults.SetupConnectionString() & DB_PWHash & ";")

        Dim InstanceMessage As String = Check_SQLInstance(DB_Instance)          'Get Instance of Local Database

        If InstanceMessage.Contains("failed".ToLower) Then                      ' Check instance is exist or not?
            MySysMessage.Add("SQL LocalDB instance not found.")
            MyErrorClass = New ErrorClass(ErrorClass.ErrorNumbers.SetupDB_Instance_NotFound)
            MyErrorClass.ErrorMessageExternal = InstanceMessage
            Return Nothing
        Else
            MySysMessage.Add("SQL LocalDB instance found.")
        End If

        If Not ErrorStatus Then                                     ' Execute is Error not Found.
            Try
                _SQLConnection.Open()                               ' Open a Local DB Connection
                MySysMessage.Add("Local Setup Connection Open")
            Catch ex As Exception
                MySysMessage.Add(ex.Message)
                MyErrorClass = New ErrorClass(ErrorClass.ErrorNumbers.LocalSetupConnection_Not_Open)
                MySysMessage.Add(MyErrorClass.GetErrorMessage())

            End Try

            If Not ErrorStatus Then
                If _SQLConnection.State = ConnectionState.Open Then
                    MySysMessage.Add("LocalDBConnected() | Local Setup Connection Sucessfully return")
                    Return _SQLConnection
                Else
                    MySysMessage.Add("LocalDBConnected() | Local Setup Connection return nothing")
                    MyErrorClass.ErrorNumber = ErrorClass.ErrorNumbers.LocalSetupConnection_Nothing
                    MySysMessage.Add(MyErrorClass.GetErrorMessage())
                    Return Nothing
                End If
            End If
        Else
            Return Nothing
        End If
        Return Nothing

    End Function             ' Provide SQL Connection of Local Setup Database

    Private Function GetPassword(_Password) As String
        Dim PWHash As New EncryptPW("Applied")
        Dim PWText As String = PWHash.DecryptData(_Password)
        Return PWText
    End Function                ' Convert Password Hash to Real Password
    Friend Function AddBackSlash(_Directory As String) As String
        If Right(_Directory, 1) = "\" Then
            Return _Directory
        Else
            Return _Directory & "\"
        End If
    End Function    ' Add a Backslech at the end of Path string  

    Friend Function Check_SQLInstance(_DBEngine As String) As String
        Dim _Result As String = "No Result"
        _Result = CreateSetup.InstanceInfo(_DBEngine)
        MySysMessage.Add("----------------------------------- Check_SQL Instance | " & _DBEngine)
        MySysMessage.Add(_Result)
        MySysMessage.Add("----------------------------------- Check_SQLInstance Result End")
        Return _Result
    End Function

    Public ReadOnly Property ErrorStatus As Boolean
        Get
            Return MyErrorClass.ErrorStatus
        End Get
    End Property

End Class




