Public Class Column
    '------------
    ' Attributes
    '------------
    Public ParentTable As Table
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
    Public Sub AddConstraint(ByVal name As String, ByVal type As Constraint._Type)
        If type = Constraint._Type.NOT_NULL OrElse type = Constraint._Type.PRIMARY_KEY OrElse Constraint._Type.UNIQUE Then
            Constraints.Add(New Constraint(name, type))
        End If
    End Sub

    Public Sub AddConstraint(ByVal name As String, ByVal type As Constraint._Type, ByVal referenceColumn As Column)
        If type = Constraint._Type.FOREIGN_KEY Then
            Constraints.Add(New Constraint(name, type, referenceColumn))
        End If
    End Sub

    Public Sub AddConstraint(ByVal name As String, ByVal type As Constraint._Type, ByVal expression As String)
        If type = Constraint._Type.CHECK Then
            Constraints.Add(New Constraint(name, type, expression))
        End If
    End Sub
End Class
