Public Class TextVisualizer


    Public Sub New(_ArrayList As ArrayList)
        InitializeComponent()

        txtVariableValue.Text = _ArrayList.ToString

        For Each _Line As String In _ArrayList
            txtVariableValue.Text += _Line + Environment.NewLine
        Next
    End Sub

    Public Property DisplayValue As ArrayList

    Private Sub TextVisualizer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnExit_Click_1(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()
    End Sub
End Class