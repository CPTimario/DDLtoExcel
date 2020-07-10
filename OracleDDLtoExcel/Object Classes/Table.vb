Public Class Table
    Public Name As String
    Public Columns As List(Of Column)

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal name As String, Optional ByVal columns As List(Of Column) = Nothing)
        Me.Name = name
        Me.Columns = IIf(columns Is Nothing, New List(Of Column), columns)
    End Sub

    '---------
    ' Methods
    '---------
    Public Sub AddColumn(ByVal name As String, ByVal dataType As DataType)
        Columns.Add(New Column(name, dataType))
    End Sub

    '-----------
    ' Functions
    '-----------
    Public Function Column(ByVal name As String) As Column
        Return Columns.Find(Function(col) col.Name = name)
    End Function
End Class
