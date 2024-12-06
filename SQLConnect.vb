Imports MySqlConnector

Module SQLConnect
    Public datacon As New MySqlConnection

    Dim server As String = "127.0.0.1"
    Dim dbasename As String = "" ' db schema
    Dim username As String = ""
    Dim passwd As String = ""

    Public Sub databaseConnect()
        Try
            datacon.ConnectionString = $"server={server};user id={username};password={passwd};database={dbasename};"

            datacon.Open()

        Catch ex As MySqlException

            MsgBox($"Database Connection Error: {ex.Message}", MsgBoxStyle.Critical)
        Catch ex As Exception

            MsgBox($"Unexpected Error: {ex.Message}", MsgBoxStyle.Critical)
        End Try
    End Sub
End Module
