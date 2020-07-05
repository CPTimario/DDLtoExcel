Public Class Column
    '------------
    ' Attributes
    '------------
    Public Name As String
    Public DataType As DataTypes
    Public DefaultValue As String
    Public Constraints As New List(Of Constraint)

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal pName As String, ByVal pDataType As DataTypes)
        Name = pName
        DataType = pDataType
    End Sub

    '---------
    ' Methods
    '---------
    Public Sub AddConstraint(ByVal pType As ConstraintType, Optional ByVal pAddlClauses As String() = Nothing)
        If Not Constraints.Exists(Function(ct) ct.Type = pType) Then
            If pType = ConstraintType.ctFOREIGN Then
                Constraints.Add(New Constraint(pType, pAddlClauses(0), pAddlClauses(1)))
            ElseIf pType = ConstraintType.ctCHECK Then
                Constraints.Add(New Constraint(pType, pAddlClauses(0)))
            Else
                Constraints.Add(New Constraint(pType))
            End If
        End If
    End Sub
End Class
