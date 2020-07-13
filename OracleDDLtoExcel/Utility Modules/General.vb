Imports System.Runtime.CompilerServices
Imports OracleDDLtoExcel

Public Module General
    '-----------
    ' Variables
    '-----------
    Public CancelFlg As Boolean = True

    Public Structure StringPair
        Public StartString As String
        Public EndString As String

        Public Sub New(ByVal startString As String, ByVal endString As String)
            Me.StartString = startString
            Me.EndString = endString
        End Sub
    End Structure

    Public Structure DataType
        Public Enum _Type
            _CHAR
            _VARCHAR2
            _NCHAR
            _NVARCHAR2
            _LONG
            _NUMBER
            _DATE
        End Enum

        Public Type As _Type
        Public Arguments As String

        Public Sub New(ByVal type As _Type, Optional ByVal arguments As String = "")
            Me.Type = type
            Me.Arguments = arguments
        End Sub
    End Structure

    Public Structure Constraint
        Public Enum _Type
            NOT_NULL
            PRIMARY_KEY
            UNIQUE
            FOREIGN_KEY
            CHECK
        End Enum

        Public Name As String
        Public Type As _Type
        Public Expression As String
        Public Reference As KeyValuePair(Of String, String)

        Public Sub New(ByVal name As String, ByVal type As _Type)
            Me.Name = name
            Me.Type = type
        End Sub

        Public Sub New(ByVal name As String, ByVal type As _Type, ByVal refTable As String, ByVal refColumn As String)
            Me.Name = name
            Me.Type = type
            Me.Reference = New KeyValuePair(Of String, String)(refTable, refColumn)
        End Sub

        Public Sub New(ByVal name As String, ByVal type As _Type, ByVal expression As String)
            Me.Name = name
            Me.Type = type
            Me.Expression = expression
        End Sub
    End Structure

    '-----------
    ' Functions
    '-----------
    <Extension>
    Public Function EnumToString(ByVal value As [Enum]) As String
        Return value.ToString("F").Replace(Chr(95), Chr(32)).Trim
    End Function

    <Extension>
    Public Function GetKeywordRegex(ByVal keywordEnum As [Enum], Optional ByVal regexPrefix As String = "", Optional ByVal regexSuffix As String = "") As String
        Dim regex As String = String.Empty
        Dim keywordString As String = keywordEnum.EnumToString()

        For Each keyword As String In keywordString.Split(Chr(32))
            regex &= regexPrefix & "(" & keyword.ToUpper & "|" & keyword.ToLower & ")" & regexSuffix
        Next

        Return regex
    End Function

    Public Function GetElements(ByVal value As String) As List(Of String)
        Dim character As Char
        Dim element As String = String.Empty
        Dim parenthesis As New Stack(Of Char)
        Dim elements As New List(Of String)

        For charIndex As Integer = 0 To value.Length - 1
            character = value.Chars(charIndex)

            If character.Equals(Chr(40)) Then
                parenthesis.Push(character)
            ElseIf character.Equals(Chr(41)) Then
                parenthesis.Pop()
            End If

            If charIndex = value.Length - 1 Then
                element &= character.ToString
                elements.Add(element.Trim)
            ElseIf parenthesis.Count = 0 AndAlso character.Equals(Chr(44)) Then
                elements.Add(element.Trim)
                element = String.Empty
            Else
                element &= character.ToString
            End If
        Next

        Return elements
    End Function

    <Extension>
    Public Function Table(ByVal tables As List(Of Table), ByVal tableName As String) As Table
        Return tables.Find(Function(tbl) tbl.Name = tableName)
    End Function
End Module
