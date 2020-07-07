Public Module General
    '------------
    ' Constants
    '------------
    Public Const TABLE_SYNTAX As String = "(TABLE\s+)\""*\w+\""*\s*\(\s*(" & COLUMN_SYNTAX & ")*\,*\s*)+\)"
    Public Const COLUMN_SYNTAX As String = "\""*\w+\""*\s*\w+(\s*\(\s*\d+\s*\,*\s*\w*\s*\))*"
    '-----------
    ' Variables
    '-----------
    Public CancelFlg As Boolean = False
    Public Tables As List(Of Table)
    Public SqlComments As New List(Of StringPair)(
        {
            New StringPair("--", vbLf),
            New StringPair("/*", "*/")
        })
    Public Parenthesis As New StringPair("(", ")")
    Public DDLCommands As New Dictionary(Of DDLCommand, String) From {
        {DDLCommand.ddlCREATE, "CREATE "},
        {DDLCommand.ddlALTER, "ALTER "},
        {DDLCommand.ddlDROP, "DROP "},
        {DDLCommand.ddlCOMMENT_ON_COLUMN, "COMMENT ON COLUMN"}
    }
    Public DDLCreateDropObjects As New Dictionary(Of DDLCreateDropObject, String) From {
        {DDLCreateDropObject.crdrTABLE, "TABLE "},
        {DDLCreateDropObject.crdrVIEW, "VIEW "},
        {DDLCreateDropObject.crdrINDEX, "INDEX "}
    }
    Public DDLAlterObjects As New Dictionary(Of DDLAlterObject, String) From {
        {DDLAlterObject.altTABLE, "TABLE "},
        {DDLAlterObject.altVIEW, "VIEW "}
    }

    Public Structure StringPair
        Public StartString As String
        Public EndString As String

        Public Sub New(ByVal startString As String, ByVal endString As String)
            Me.StartString = startString
            Me.EndString = endString
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
        ddlCOMMENT_ON_COLUMN
    End Enum

    Public Enum DDLCreateDropObject
        crdrTABLE
        crdrVIEW
        crdrINDEX
    End Enum

    Public Enum DDLAlterObject
        altTABLE
        altVIEW
    End Enum
End Module
