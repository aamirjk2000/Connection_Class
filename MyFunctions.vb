Module MyFunctions

    Public Function IIF(_Logic As Boolean, Condition_true As Object, Condition_False As Object) As Object
        If _Logic Then
            Return Condition_true
        Else
            Return Condition_False
        End If
    End Function


    Public Function MaxID(_DataTable As DataTable) As Integer
        Dim _MaxID As Integer = 0
        _MaxID = _DataTable.Compute("MAX(ID)", "")
        Return _MaxID + 1
    End Function

End Module
