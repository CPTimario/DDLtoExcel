Public Class Table
    Public Name As String
    Public Columns As New List(Of Column)

    Public Sub AddColumn(ByVal pName As String, ByVal pDataType As DataType)
        Columns.Add(New Column(pName, pDataType))
    End Sub
End Class
