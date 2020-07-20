Public Class ErrorClass

    Dim _ErrorNumber As Integer
    Dim _ErrorMessage As String
    Dim _ErrorGroup As String
    Dim _ErrorStatus As Boolean
    Dim _ErrorMessageExternal As String

    Public Sub New()
        ErrorNumber = -1
        ErrorMessage = ""
        ErrorGroup = ""
        ErrorStatus = False
        ErrorMessageExternal = ""
    End Sub

    Public Sub New(ErrorID As Integer)

        ErrorNumber = ErrorID
        ErrorMessage = GetErrorMessage(ErrorID)
        ErrorGroup = GetErrorGroup(ErrorID)
        ErrorStatus = True
        ErrorMessageExternal = ""

    End Sub

    Public Property ErrorNumber
        Set(value)
            _ErrorNumber = value
        End Set
        Get
            Return _ErrorNumber
        End Get
    End Property

    Public Property ErrorMessage
        Set(value)
            _ErrorMessage = value
        End Set
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public Property ErrorMessageExternal
        Set(value)
            _ErrorMessageExternal = value
        End Set
        Get
            Return _ErrorMessageExternal
        End Get
    End Property

    Public Property ErrorGroup
        Set(value)
            _ErrorGroup = value
        End Set
        Get
            Return _ErrorGroup
        End Get
    End Property

    Public Property ErrorStatus
        Set(value)
            _ErrorStatus = value
        End Set
        Get
            Return _ErrorStatus
        End Get
    End Property

    Enum ErrorNumbers
        DBFileNotFound = 100
        LogFileNotFound = 101

        LocalSetupConnection_Not_Open = 102
        LocalSetupConnection_Nothing = 103

        SetupDB_Instance_NotFound = 201
        SetupDB_Instance_NotRunning = 202
        LocalDB_Scheme_NotFound = 203
    End Enum


    Public Function GetErrorMessage(_Number As Integer) As String
        Dim _ErrorString As String = ""

        Select Case _Number
            Case ErrorNumbers.DBFileNotFound
                _ErrorString = "Local Database files does not exist."
            Case ErrorNumbers.LogFileNotFound
                _ErrorString = "Local Database LOG files does not exist."
            Case ErrorNumbers.LocalSetupConnection_Not_Open
                _ErrorString = "Local Setup Connection NOT Opened."
            Case ErrorNumbers.LocalSetupConnection_Nothing
                _ErrorString = "Local Setup Connection is nothing."
            Case ErrorNumbers.SetupDB_Instance_NotFound
                _ErrorString = "SQL Local DB Instance SetupDB not found or not created."
            Case ErrorNumbers.SetupDB_Instance_NotRunning
                _ErrorString = "SQL Local DB Instance 'SetupDB' is not running now."
            Case ErrorNumbers.LocalDB_Scheme_NotFound
                _ErrorString = "Local Database Scheme is not define correct."

        End Select

        Return _ErrorString
    End Function

    Private Function GetErrorGroup(_Number As Integer) As String
        Dim _ErrorString As String = ""

        Select Case _Number
            Case ErrorNumbers.DBFileNotFound
                _ErrorString = "DB_STRING"
            Case ErrorNumbers.LogFileNotFound
                _ErrorString = "DB_STRING"
            Case ErrorNumbers.LocalSetupConnection_Not_Open
                _ErrorString = "DB_STRING"
            Case ErrorNumbers.LocalSetupConnection_Nothing
                _ErrorString = "DB_STRING"
            Case ErrorNumbers.SetupDB_Instance_NotFound
                _ErrorString = "INSTANCE"
            Case ErrorNumbers.SetupDB_Instance_NotRunning
                _ErrorString = "INSTANCE"
            Case ErrorNumbers.LocalDB_Scheme_NotFound
                _ErrorString = "SCHEME"
        End Select

        Return _ErrorString
    End Function



    Public Function GetErrorMessage() As String
        Return ErrorNumber.ToString & " | " & GetErrorMessage(ErrorNumber)
    End Function


End Class
