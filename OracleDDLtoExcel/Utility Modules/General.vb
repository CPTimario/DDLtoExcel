Public Module General
    '-----------
    ' Variables
    '-----------
    Public CancelFlg As Boolean = False
    Public SqlComments As New List(Of SQlComment)(
        {
            New SQlComment("SINGLE", "--", vbLf),
            New SQlComment("MULTI", "/*", "*/")
        })
    Public DDLCommands As New Dictionary(Of DDLCommand, String) From {
        {DDLCommand.ddlCREATE, "CREATE "},
        {DDLCommand.ddlALTER, "ALTER "},
        {DDLCommand.ddlDROP, "DROP "},
        {DDLCommand.ddlCOMMENT_ON, "COMMENT ON "}
    }

    '------------
    ' Structures
    '------------
    Public Structure EnumPrefix
        Public Const CONSTRAINT_TYPE As String = "ct"
        Public Const DATA_TYPE As String = "dt"
        Public Const DDL_COMMAND As String = "ddl"
    End Structure

    Public Structure SQlComment
        Public Type As String
        Public StartString As String
        Public EndString As String

        Public Sub New(ByVal pType As String, ByVal pStartString As String, ByVal pEndString As String)
            Type = pType
            StartString = pStartString
            EndString = pEndString
        End Sub
    End Structure

    Public Structure Constraint
        Public Type As ConstraintType
        Public Expression As String
        Public Reference As KeyValuePair(Of String, String)

        Public Sub New(ByVal type As ConstraintType)
            Me.Type = type
        End Sub

        Public Sub New(ByVal type As ConstraintType, ByVal refTable As String, ByVal refColumn As String)
            Me.Type = type
            Me.Reference = New KeyValuePair(Of String, String)(refTable, refColumn)
        End Sub

        Public Sub New(ByVal type As ConstraintType, ByVal expression As String)
            Me.Type = type
            Me.Expression = expression
        End Sub
    End Structure

    '--------------
    ' Enumerations
    '--------------
    Public Enum ConstraintType
        ctNOTNULL
        ctPRIMARY
        ctUNIQUE
        ctFOREIGN
        ctCHECK
    End Enum

    Public Enum DataType
        dtCHAR
        dtVARCHAR2
        dtNCHAR
        dtNVARCHAR2
        dtLONG
        dtNUMBER
        dtDATE
    End Enum

    Public Enum DDLCommand
        ddlCREATE
        ddlALTER
        ddlDROP
        ddlCOMMENT_ON
    End Enum
End Module
