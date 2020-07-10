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
        DEFAULT_VALUE
    End Enum

    Private Enum ColumnCommentGroup
        NONE
        TABLE_NAME
        COLUMN_NAME
        COMMENT
    End Enum

    '-----------
    ' Structure
    '-----------
    Private Structure CreateTableRegex
        Const TABLE_NAME As String = "(?<=CREATE\sTABLE\s)(?:[\""\']?(\w+)[\""\']?\.)?[\""\']*(\w+)[\""\']?"
        Const COLUMN_LIST As String = "\(\s*(" & COLUMN & "[\s\,]+)+\)"
        Const COLUMN As String = "[\""\']?(\w+)[\""\']?\s+(\w+)\s*(\([\w\s\,]+\))?\s*(?:DEFAULT\s+(\w+))?"
        Const COLUMN_COMMENT As String = "[\""\']?(\w+)[\""\']?\.[\""\']?(\w+)[\""\']?\s+IS\s+[\""\']?([\w一-龠ぁ-ゔァ-ヴーａ-ｚＡ-Ｚ０-９々〆〤]+)[\""\']?"
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
        columns = GetColumns(Regex.Match(command, CreateTableRegex.COLUMN_LIST).Value)

        table = New Table(tableName, columns)
        Tables.Add(table)
    End Sub

    Public Sub CommentOnColumn(ByVal command As String)
        Dim tableName As String
        Dim columnName As String
        Dim comment As String
        Dim table As Table
        Dim commentGroups As GroupCollection

        commentGroups = Regex.Match(command, CreateTableRegex.COLUMN_COMMENT).Groups
        tableName = commentGroups.Item(ColumnCommentGroup.TABLE_NAME).ToString
        columnName = commentGroups.Item(ColumnCommentGroup.COLUMN_NAME).ToString
        comment = commentGroups.Item(ColumnCommentGroup.COMMENT).ToString

        table = Tables.Find(Function(tbl) tbl.Name = tableName)
        Call table.GetColumn(columnName).AddComment(comment)
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

    Private Function GetColumns(ByVal columnList As String) As List(Of Column)
        Dim columns As New List(Of Column)
        Dim columnMatches As MatchCollection = Regex.Matches(columnList, CreateTableRegex.COLUMN)

        For Each match As Match In columnMatches
            Dim column As Column
            Dim dataType As DataType
            Dim columnName As String = match.Groups.Item(ColumnGroup.COLUMN_NAME).ToString
            Dim dataTypeString As String = match.Groups.Item(ColumnGroup.DATA_TYPE).ToString
            Dim dataTypeArgs As String = match.Groups.Item(ColumnGroup.ARGUMENTS).ToString
            Dim defaultValue As String = match.Groups.Item(ColumnGroup.DEFAULT_VALUE).ToString

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
