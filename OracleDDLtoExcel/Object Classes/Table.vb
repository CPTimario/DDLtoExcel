Public Class Table
    Public Name As String
    Public Columns As New List(Of Column)

    Public Sub AddColumn(ByVal name As String, ByVal dataType As DataType)
        Columns.Add(New Column(name, dataType))
    End Sub
End Class
