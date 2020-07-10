Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Public Class SQL
    '-----------
    ' Variables
    '-----------
    Public Tables As List(Of Table)
    Private SQLCommandString As String
    Private SQLCommands As List(Of String)

    '-----------
    ' Constants
    '-----------
    Private SQL_COMMENTS As New List(Of StringPair)(
        {
            New StringPair("--", vbLf),
            New StringPair("/*", "*/")
        })

    '--------------
    ' Enumerations
    '--------------
    Private Enum DDLCommand
        CREATE_TABLE
        CREATE_VIEW
        CREATE_INDEX
        ALTER_TABLE
        ALTER_VIEW
        DROP_TABLE
        DROP_VIEW
        DROP_INDEX
        COMMENT_ON_COLUMN
    End Enum

    Private Enum TableGroup
        NONE
        SCHEMA_NAME
        TABLE_NAME
    End Enum

    Private Enum ColumnGroup
        NONE
        COLUMN_NAME
        DATA_TYPE
        ARGUMENTS
        AUTO_DEFAULT
        DEFAULT_VALUE
    End Enum

    Private Enum ColumnCommentGroup
        NONE
        TABLE_NAME
        COLUMN_NAME
        COMMENT
    End Enum

    Private Enum NotNullGroup
        NONE
        TABLE_NAME
        COLUMN_NAME
        CONSTRAINT_NAME
    End Enum

    Private Enum PrimaryKeyGroup
        NONE
        TABLE_NAME
        CONSTRAINT_NAME
        COLUMN_LIST
    End Enum

    Private Enum UniqueGroup
        NONE
        TABLE_NAME
        CONSTRAINT_NAME
        COLUMN_LIST
    End Enum

    Private Enum ForeignGroup
        NONE
        TABLE_NAME
        CONSTRAINT_NAME
        COLUMN_LIST
        REF_TABLE_NAME
        REF_COLUMN_LIST
    End Enum

    Private Enum CheckGroup
        NONE
        TABLE_NAME
        CONSTRAINT_NAME
        COLUMN_NAME
        CONDITION
    End Enum

    '-----------
    ' Structure
    '-----------
    Private Structure CreateTableRegex
        Const TABLE_NAME As String = "(?<=CREATE\sTABLE\s)(?:[\""\']?(\w+)[\""\']?\.)?[\""\']*(\w+)[\""\']?"
        Const COLUMN_LIST As String = "(?<=\()\s*(" & COLUMN & "[\s\,]+)+"
        Const COLUMN As String = "[\""\']?(\w+)[\""\']?\s+(\w+)\s*(\([\w\s\,]+\))?\s*((?:DEFAULT\s+([\w\'\""]+))|(?:AUTO INCREMENT))?"
        Const COLUMN_COMMENT As String = "[\""\']?(\w+)[\""\']?\.[\""\']?(\w+)[\""\']?\s+IS\s+[\""\']?([\w一-龠ぁ-ゔァ-ヴーａ-ｚＡ-Ｚ０-９々〆〤]+)[\""\']?"
    End Structure

    Private Structure AlterTableRegex
        Const COLUMN_LIST = "\(\s*((?:[\""\']?\w+[\""\']?\,*\s*)+)\s*\)"
        Const NOT_NULL = "(?<=ALTER\sTABLE\s)[\""\']?(\w+)[\""\']?\s+MODIFY\s*\([\""\']?(\w+)[\""\']?\s+CONSTRAINT\s+[\""\']?(\w+)[\""\']?\s+NOT\s+NULL\s+ENABLE\s*\)"
        Const PRIMARY_KEY = "(?<=ALTER\sTABLE\s)[\""\']?(\w+)[\""\']?\s+ADD\s+CONSTRAINT\s+[\""\']?(\w+)[\""\']?\s+PRIMARY\s+KEY\s*" & COLUMN_LIST & "\s*ENABLE"
        Const UNIQUE = "(?<=ALTER\sTABLE\s)[\""\']?(\w+)[\""\']?\s+ADD\s+CONSTRAINT\s+[\""\']?(\w+)[\""\']?\s+UNIQUE\s*" & COLUMN_LIST & "\s*ENABLE"
        Const CHECK = "(?<=ALTER\sTABLE\s)[\""\']?(\w+)[\""\']?\s+ADD\s+CONSTRAINT\s+[\""\']?(\w+)[\""\']?\s+CHECK\s*\(\s*[\""\']?(\w+)[\""\']?(.+)\)"
        Const FOREIGN_KEY = "(?<=ALTER\sTABLE\s)[\""\']?(\w+)[\""\']?\s+ADD\s+CONSTRAINT\s+[\""\']?(\w+)[\""\']?\s+FOREIGN\s+KEY\s*" & COLUMN_LIST & "\s*REFERENCES\s+[\""\']?(\w+)[\""\']?\s*" & COLUMN_LIST & "\s+ENABLE"
    End Structure

    '-------------
    ' Constructor
    '-------------
    Public Sub New(ByVal command As String)
        SQLCommandString = command
        SQLCommands = New List(Of String)
        Tables = New List(Of Table)
    End Sub

    '---------
    ' Methods
    '---------
    Public Sub RemoveComments()
        Dim commentCount As Integer = 0
        Dim progressValue As Integer = 0
        Dim subString As String

        For Each comment As StringPair In SQL_COMMENTS
            commentCount += SQLCommandString.SubstringCount(comment.StartString, comment.EndString)
        Next

        Call ShowStatus(REMOVE_COMMENTS, progressValue, commentCount, New String() {progressValue.ToString, commentCount.ToString})

        For Each sqlComment As StringPair In SQL_COMMENTS
            If CancelFlg Then Exit Sub

            While Not SQLCommandString.Substring(sqlComment.StartString, sqlComment.EndString) = String.Empty
                If CancelFlg Then Exit Sub

                subString = SQLCommandString.Substring(sqlComment.StartString, sqlComment.EndString)
                progressValue += SQLCommandString.SubstringCount(subString)
                SQLCommandString = SQLCommandString.Replace(subString, String.Empty).Trim()

                Call ShowStatus(REMOVE_COMMENTS, progressValue, New String() {progressValue.ToString, commentCount.ToString})
                Application.DoEvents()
            End While
        Next
    End Sub

    Public Sub ExecuteCommands()
        Dim ddlCommand As DDLCommand
        Dim progressValue As Integer = 0

        Call GetSQLCommands()

        Call ShowStatus(EXECUTE_COMMANDS, progressValue, SQLCommands.Count, New String() {progressValue.ToString, SQLCommands.Count.ToString})

        For Each command As String In SQLCommands
            progressValue += 1

            If CancelFlg Then Exit Sub

            ddlCommand = GetDDLCommandType(command)
            Call CallByName(Me, GetMethodName(ddlCommand), CallType.Method, command.Trim)

            Call ShowStatus(EXECUTE_COMMANDS, progressValue, New String() {progressValue.ToString, SQLCommands.Count.ToString})
            Application.DoEvents()
        Next
    End Sub

    Private Sub GetSQLCommands()
        Dim ddlCommandString As String
        Dim commandRegex As String
        Dim replaceString As String

        For Each command As DDLCommand In [Enum].GetValues(GetType(DDLCommand))
            If SQLCommandString.Contains(command.EnumToString()) Then
                ddlCommandString = command.EnumToString()
                commandRegex = command.GetKeywordRegex(String.Empty, "\s+")
                replaceString = Chr(36) & ddlCommandString & Chr(32)
                SQLCommandString = Regex.Replace(SQLCommandString, commandRegex, replaceString)
            End If
        Next

        SQLCommands = SQLCommandString.Split(Chr(36)).ToList
        SQLCommands.Remove(String.Empty)
    End Sub

    Public Sub CreateTable(ByVal command As String)
        Dim table As Table
        Dim columns As List(Of Column)
        Dim tableName As String
        Dim tableGroups As GroupCollection

        tableGroups = Regex.Match(command, CreateTableRegex.TABLE_NAME).Groups
        tableName = tableGroups.Item(TableGroup.TABLE_NAME).ToString
        columns = GetColumns(Regex.Match(command, CreateTableRegex.COLUMN_LIST).ToString.Trim)

        table = New Table(tableName, columns)
        Tables.Add(table)
    End Sub

    Public Sub CommentOnColumn(ByVal command As String)
        Dim tableName As String
        Dim columnName As String
        Dim comment As String
        Dim commentGroups As GroupCollection

        commentGroups = Regex.Match(command, CreateTableRegex.COLUMN_COMMENT).Groups
        tableName = commentGroups.Item(ColumnCommentGroup.TABLE_NAME).ToString
        columnName = commentGroups.Item(ColumnCommentGroup.COLUMN_NAME).ToString
        comment = commentGroups.Item(ColumnCommentGroup.COMMENT).ToString

        Tables.Table(tableName).Column(columnName).Comment = comment
    End Sub

    Public Sub AlterTable(ByVal command As String)
        For Each constraintType As Constraint._Type In [Enum].GetValues(GetType(Constraint._Type))
            If command.Contains(constraintType.EnumToString) Then
                AddConstraint(constraintType, command)
            End If
        Next
    End Sub

    Private Sub AddConstraint(ByVal constraintType As Constraint._Type, ByVal command As String)
        Dim constraintGroups As GroupCollection

        Select Case constraintType
            Case Constraint._Type.NOT_NULL
                constraintGroups = Regex.Match(command, AlterTableRegex.NOT_NULL, RegexOptions.IgnoreCase).Groups
                Call AddNotNullConstraint(constraintGroups)

            Case Constraint._Type.PRIMARY_KEY
                constraintGroups = Regex.Match(command, AlterTableRegex.PRIMARY_KEY, RegexOptions.IgnoreCase).Groups
                Call AddPrimaryKeyConstraint(constraintGroups)

            Case Constraint._Type.UNIQUE
                constraintGroups = Regex.Match(command, AlterTableRegex.UNIQUE, RegexOptions.IgnoreCase).Groups
                Call AddUniqueConstraint(constraintGroups)

            Case Constraint._Type.FOREIGN_KEY
                constraintGroups = Regex.Match(command, AlterTableRegex.FOREIGN_KEY, RegexOptions.IgnoreCase).Groups
                Call AddForeignKeyConstraint(constraintGroups)

            Case Constraint._Type.CHECK
                constraintGroups = Regex.Match(command, AlterTableRegex.CHECK, RegexOptions.IgnoreCase).Groups
                Call AddCheckConstraint(constraintGroups)
        End Select
    End Sub

    Private Sub AddNotNullConstraint(ByVal constraintGroups As GroupCollection)
        Dim constraintName As String = constraintGroups.Item(NotNullGroup.CONSTRAINT_NAME).ToString
        Dim tableName As String = constraintGroups.Item(NotNullGroup.TABLE_NAME).ToString
        Dim columnName As String = constraintGroups.Item(NotNullGroup.COLUMN_NAME).ToString

        Tables.Table(tableName).Column(columnName).AddConstraint(constraintName, Constraint._Type.NOT_NULL)
    End Sub

    Private Sub AddPrimaryKeyConstraint(ByVal constraintGroups As GroupCollection)
        Dim constraintName As String = constraintGroups.Item(PrimaryKeyGroup.CONSTRAINT_NAME).ToString
        Dim tableName As String = constraintGroups.Item(PrimaryKeyGroup.TABLE_NAME).ToString
        Dim columnList As List(Of String) = GetElements(constraintGroups.Item(PrimaryKeyGroup.COLUMN_LIST).ToString)

        For Each columnName As String In columnList
            columnName = columnName.Replace(Chr(34), String.Empty)
            Tables.Table(tableName).Column(columnName).AddConstraint(constraintName, Constraint._Type.PRIMARY_KEY)
        Next
    End Sub

    Private Sub AddUniqueConstraint(ByVal constraintGroups As GroupCollection)
        Dim constraintName As String = constraintGroups.Item(UniqueGroup.CONSTRAINT_NAME).ToString
        Dim tableName As String = constraintGroups.Item(UniqueGroup.TABLE_NAME).ToString
        Dim columnList As List(Of String) = GetElements(constraintGroups.Item(UniqueGroup.COLUMN_LIST).ToString)

        For Each columnName As String In columnList
            columnName = columnName.Replace(Chr(34), String.Empty)
            Tables.Table(tableName).Column(columnName).AddConstraint(constraintName, Constraint._Type.UNIQUE)
        Next
    End Sub

    Private Sub AddForeignKeyConstraint(ByVal constraintGroups As GroupCollection)
        Dim constraintName As String = constraintGroups.Item(ForeignGroup.CONSTRAINT_NAME).ToString
        Dim tableName As String = constraintGroups.Item(ForeignGroup.TABLE_NAME).ToString
        Dim columnList As List(Of String) = GetElements(constraintGroups.Item(ForeignGroup.COLUMN_LIST).ToString)
        Dim refTableName As String = constraintGroups.Item(ForeignGroup.REF_TABLE_NAME).ToString
        Dim refColumnList As List(Of String) = GetElements(constraintGroups.Item(ForeignGroup.REF_COLUMN_LIST).ToString)
        Dim addlClause As List(Of String)

        For Each columnName As String In columnList
            addlClause = New List(Of String)
            addlClause.Add(tableName)
            addlClause.Add(refColumnList.Item(columnList.IndexOf(columnName)))
            columnName = columnName.Replace(Chr(34), String.Empty)
            Tables.Table(tableName).Column(columnName).AddConstraint(constraintName, Constraint._Type.FOREIGN_KEY, addlClause)
        Next
    End Sub

    Private Sub AddCheckConstraint(ByVal constraintGroups As GroupCollection)
        Dim constraintName As String = constraintGroups.Item(CheckGroup.CONSTRAINT_NAME).ToString
        Dim tableName As String = constraintGroups.Item(CheckGroup.TABLE_NAME).ToString
        Dim columnName As String = constraintGroups.Item(CheckGroup.COLUMN_NAME).ToString
        Dim condition As String = constraintGroups.Item(CheckGroup.CONDITION).ToString
        Dim addlClause As New List(Of String)({condition})

        Tables.Table(tableName).Column(columnName).AddConstraint(constraintName, Constraint._Type.CHECK, addlClause)
    End Sub

    '------------
    ' Functions
    '------------
    Private Function GetDDLCommandType(ByVal commandString As String) As DDLCommand
        GetDDLCommandType = Nothing

        For Each command As DDLCommand In [Enum].GetValues(GetType(DDLCommand))
            If commandString.Contains(command.EnumToString()) Then
                Return command
            End If
        Next
    End Function

    Private Function GetMethodName(ByVal command As DDLCommand) As String
        Dim methodName As String

        methodName = command.EnumToString()
        methodName = StrConv(methodName.ToLower, VbStrConv.ProperCase)
        methodName = methodName.Replace(Chr(32), String.Empty)

        Return methodName
    End Function

    Private Function GetColumns(ByVal command As String) As List(Of Column)
        Dim columns As New List(Of Column)
        Dim columnList As List(Of String)
        Dim columnGroups As GroupCollection
        Dim column As Column
        Dim dataType As DataType
        Dim columnName As String
        Dim dataTypeString As String
        Dim dataTypeArgs As String
        Dim autoDefault As String
        Dim defaultValue As String

        columnList = GetElements(command)
        For Each columnString As String In columnList
            columnGroups = Regex.Match(columnString, CreateTableRegex.COLUMN).Groups

            columnName = columnGroups(ColumnGroup.COLUMN_NAME).ToString
            dataTypeString = columnGroups(ColumnGroup.DATA_TYPE).ToString
            dataTypeArgs = columnGroups(ColumnGroup.ARGUMENTS).ToString
            autoDefault = columnGroups(ColumnGroup.AUTO_DEFAULT).ToString

            If autoDefault.Contains("DEFAULT") Then
                defaultValue = columnGroups(ColumnGroup.DEFAULT_VALUE).ToString
            Else
                defaultValue = columnGroups(ColumnGroup.AUTO_DEFAULT).ToString
            End If

            dataType = GetDataType(dataTypeString, dataTypeArgs)
            column = New Column(columnName, dataType, defaultValue)
            columns.Add(column)
        Next

        Return columns
    End Function

    Private Function GetDataType(ByVal dataTypeString As String, Optional ByVal arguments As String = "") As DataType
        Dim dataType As DataType._Type = [Enum].Parse(GetType(DataType._Type), Chr(95) & dataTypeString)
        Return New DataType(dataType, arguments)
    End Function
End Class
