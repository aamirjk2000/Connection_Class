Imports System.Data.SqlClient


Public Class Default_Values

    Public Property SQLInstance As String = ""
    Public Property DBEngine As String = ""
    Public Property DBDatabase As String = ""
    Public Property DBSchema As String = ""
    Public Property DBSetupTable As String = ""
    Public Property DBPath As String = ""
    Public Property DBLogin As String
    Public Property DBUser As String = ""
    Public Property DBPWHash As String = ""
    Public Property DBPWWrapper As String = ""
    Public Property DBSetupFile As String = ""
    Public Property DBSetupLog As String = ""
    Public Property ShowMessages As Boolean = True
    Public Property SysMessages As New ArrayList

    Public ReadOnly Property DBServer As String
        Get
            Return DBEngine & "\" & SQLInstance '& SQLIntance
        End Get
    End Property

    Public ReadOnly Property SetupConnectionString As String
        Get
            Return "Server=" & DBServer & ";Initial Catalog=" & DBDatabase & ";User ID=" & DBLogin & ";Password="
        End Get
    End Property

    Public ReadOnly Property SetupTableName As String
        Get
            Return DBSchema & "." & DBSetupTable
        End Get
    End Property

    Public Sub New()

        DBEngine = "(LocalDB)"
        SQLInstance = "SetupDB"
        DBDatabase = "AppliedSetupDB"
        DBSchema = "[Applied]"
        DBSetupTable = "[AppliedSetupDBTable]"
        DBPath = Environment.SpecialFolder.Personal
        DBLogin = "AppliedLogin"
        DBUser = "AppliedUser"
        DBPWHash = "A369Dvl/h7fvYxdQCD52YQUjfRVMo8O3"   ' Applied123!
        DBPWWrapper = "Applied"
        DBSetupFile = "AppliedSetupDB.mdf"
        DBSetupLog = "AppliedSetupDB.ldf"
        ShowMessages = True

    End Sub

    Public Function SetupConnection() As SqlConnection
        ' This will provide Local Database Connection (Setup DB Connection).

        SysMessages = New ArrayList
        SysMessages.Add("")
        SysMessages.Add("Start.....LocalDBConnected()")

        ' Generate Local DB String
        Dim _SQLConnection As New SqlConnection(SetupConnectionString() & PW.GetPassword(DBPWHash, DBPWWrapper) & ";")     'Establish a Local Connection.

        SysMessages.Add(SetupConnectionString() & DBPWHash & ";")

        Try
            _SQLConnection.Open()                               ' Open a Local DB Connection
            SysMessages.Add("Local Setup Connection Open")
        Catch ex As Exception
            SysMessages.Add("Setup Connection has error.")
            SysMessages.Add(ex.Message)
        End Try


        Return _SQLConnection

    End Function             ' Provide SQL Connection of Local Setup Database



    '======================================================================================

    Public Sub New(_Engine As String, _Instance As String, _Login As String, _PWHash As String)

        DBEngine = _Engine
        SQLInstance = _Instance
        DBDatabase = "AppliedSetupDB"
        DBSchema = "[Applied]"
        DBSetupTable = "[AppliedSetupDBTable]"
        DBPath = Environment.SpecialFolder.Personal
        DBLogin = _Login
        DBUser = "AppliedUser"
        DBPWHash = _PWHash
        DBPWWrapper = "Applied"
        DBSetupFile = "AppliedSetupDB.mdf"
        DBSetupLog = "AppliedSetupDB.ldf"
        ShowMessages = True

    End Sub

    Public Sub New(_Engine As String, _Instance As String, _Login As String, _PWHash As String, _FilePath As String, _DBFile As String)

        DBEngine = _Engine
        SQLInstance = _Instance
        DBDatabase = "AppliedSetupDB"
        DBSchema = "[Applied]"
        DBSetupTable = "[AppliedSetupDBTable]"
        DBPath = _FilePath
        DBLogin = _Login
        DBUser = "AppliedUser"
        DBPWHash = _PWHash
        DBPWWrapper = "Applied"
        DBSetupFile = _DBFile & ".mdf"
        DBSetupLog = _DBFile & ".ldf"
        ShowMessages = True

    End Sub


End Class
