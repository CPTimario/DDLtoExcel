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
        Public Const ENUM_PREFIX As String = "dt"

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
        Public Const ENUM_PREFIX As String = "ct"

        Public Enum _Type
            _NOT_NULL
            _PRIMARY
            _UNIQUE
            _FOREIGN
            _CHECK
        End Enum

        Public Type As _Type
        Public Expression As String
        Public Reference As KeyValuePair(Of String, String)

        Public Sub New(ByVal type As _Type)
            Me.Type = type
        End Sub

        Public Sub New(ByVal type As _Type, ByVal refTable As String, ByVal refColumn As String)
            Me.Type = type
            Me.Reference = New KeyValuePair(Of String, String)(refTable, refColumn)
        End Sub

        Public Sub New(ByVal type As _Type, ByVal expression As String)
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
End Module
