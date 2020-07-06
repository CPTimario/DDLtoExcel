Public Class Column
    '------------
    ' Attributes
    '------------
    Public Name As String
    Public DataType As DataType
    Public DefaultValue As String
    Public Constraints As New List(Of Constraint)

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal name As String, ByVal dataType As DataType)
        Me.Name = name
        Me.DataType = dataType
    End Sub

    '---------
    ' Methods
    '---------
    Public Sub AddConstraint(ByVal type As ConstraintType, Optional ByVal addlClauses As String() = Nothing)
        If Not Constraints.Exists(Function(ct) ct.Type = type) Then
            If type = ConstraintType.ctFOREIGN Then
                Constraints.Add(New Constraint(type, addlClauses(0), addlClauses(1)))
            ElseIf type = ConstraintType.ctCHECK Then
                Constraints.Add(New Constraint(type, addlClauses(0)))
            Else
                Constraints.Add(New Constraint(type))
            End If
        End If
    End Sub
End Class
