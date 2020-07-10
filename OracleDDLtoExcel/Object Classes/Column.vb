Public Class Column
    '------------
    ' Attributes
    '------------
    Public Name As String
    Public DataType As DataType
    Public DefaultValue As String
    Public Constraints As List(Of Constraint)
    Public Comment As String

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal name As String, ByVal dataType As DataType, Optional ByVal defaultValue As String = "", Optional ByVal constraints As List(Of Constraint) = Nothing)
        Me.Name = name
        Me.DataType = dataType
        Me.DefaultValue = defaultValue
        Me.Constraints = IIf(constraints Is Nothing, New List(Of Constraint), constraints)
    End Sub

    '---------
    ' Methods
    '---------
    Public Sub AddConstraint(ByVal name As String, ByVal type As Constraint._Type, Optional ByVal addlClauses As List(Of String) = Nothing)
        If Not Constraints.Exists(Function(ct) ct.Type = type) Then
            If type = Constraint._Type.FOREIGN_KEY Then
                Constraints.Add(New Constraint(name, type, addlClauses.Item(0), addlClauses.Item(1)))
            ElseIf type = Constraint._Type.CHECK Then
                Constraints.Add(New Constraint(name, type, addlClauses.Item(0)))
            Else
                Constraints.Add(New Constraint(name, type))
            End If
        End If
    End Sub
End Class
