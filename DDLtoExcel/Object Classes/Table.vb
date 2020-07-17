Public Class Table
    Public ParentSchema As Schema
    Public Name As String
    Public Columns As List(Of Column)
    Public Comment As String

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal parentSchema As Schema, ByVal name As String, ByVal columns As List(Of Column))
        Me.ParentSchema = parentSchema
        Me.Name = name
        Me.Columns = columns

        For Each col In Me.Columns
            col.ParentTable = Me
        Next
    End Sub

    '-----------
    ' Functions
    '-----------
    Public Function Column(ByVal name As String) As Column
        Return Columns.Find(Function(col) col.Name = name)
    End Function
End Class
