Imports System.Runtime.CompilerServices
Imports DDLtoExcel

Public Module General
    '-----------
    ' Variables
    '-----------
    Public CancelFlg As Boolean = True

    '------------
    ' Structures
    '------------
    Public Structure SchemaFormControls
        Property Label As Label
        Property Textbox As TextBox
        Property Open As Button
        Property Remove As Button

        Public Sub New(ByVal label As Label, ByVal textbox As TextBox, ByVal open As Button, ByVal remove As Button)
            Me.Label = label
            Me.Textbox = textbox
            Me.Open = open
            Me.Remove = remove
        End Sub
    End Structure

    Public Structure DataType
        Public Enum _Type
            _CHAR
            _VARCHAR2
            _NCHAR
            _NVARCHAR2
            _LONG_RAW
            _NUMBER
            _DATE
            _BLOB
            _ROWID
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
        Public ReferenceColumn As Column

        Public Sub New(ByVal name As String, ByVal type As _Type)
            Me.Name = name
            Me.Type = type
        End Sub

        Public Sub New(ByVal name As String, ByVal type As _Type, ByVal referenceColumn As Column)
            Me.Name = name
            Me.Type = type
            Me.ReferenceColumn = referenceColumn
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
    Public Function ToRegex(ByVal keywordEnum As [Enum], Optional ByVal regexPrefix As String = "", Optional ByVal regexSuffix As String = "") As String
        Dim regex As String
        Dim keywordString As String = keywordEnum.EnumToString()

        regex = regexPrefix
        regex &= keywordString.Replace(Chr(32), "\s+")
        regex &= regexSuffix

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
                element = element.Replace(Chr(34), String.Empty).Trim
                elements.Add(element)
            ElseIf parenthesis.Count = 0 AndAlso character.Equals(Chr(44)) Then
                element = element.Replace(Chr(34), String.Empty).Trim
                elements.Add(element)
                element = String.Empty
            Else
                element &= character.ToString
            End If
        Next

        Return elements
    End Function
End Module
