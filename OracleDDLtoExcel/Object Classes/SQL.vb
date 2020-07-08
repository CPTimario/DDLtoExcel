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
        SCHEMA_NAME = 2
        TABLE_NAME = 3
    End Enum

    Private Enum ColumnGroup
        COLUMN_NAME = 1
        DATA_TYPE = 2
        ARGUMENTS = 3
    End Enum

    '-----------
    ' Structure
    '-----------
    Private Structure CreateTableRegex
        Const TABLE_NAME As String = "(?<=CREATE\sTABLE\s)(\""*(\w+)\""\.)?\""*(\w+)\""*"
        Const COLUMN_LIST As String = "\(\s*(" & COLUMN & "[\s\,]+)+\)"
        Const COLUMN As String = "(\""*\w+\""*)\s+(\w+)\s*((\([\w\s\,]+\)))?"
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

        Call GetSQLCommands()
        For Each command As String In SQLCommands
            ddlCommand = GetDDLCommandType(command)
            Call CallByName(Me, GetMethodName(ddlCommand), CallType.Method, command.Trim)
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
        Dim tableName As String = String.Empty
        Dim tableGroups As GroupCollection = Regex.Match(command, CreateTableRegex.TABLE_NAME).Groups

        tableName = tableGroups.Item(TableGroup.TABLE_NAME).ToString
        columns = GetColumns(Regex.Match(command, CreateTableRegex.COLUMN_LIST).Value)

        table = New Table(tableName, columns)
        Tables.Add(table)
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
        Dim columnStrings As List(Of String) = Regex.Split(columnList, CreateTableRegex.COLUMN).ToList

        For Each columnString As String In columnStrings
            Dim column As Column
            Dim columnGroups As GroupCollection = Regex.Match(columnString, CreateTableRegex.COLUMN).Groups
            Dim columnName As String = columnGroups.Item(ColumnGroup.COLUMN_NAME).ToString
            Dim dataType As DataType = GetDataType(columnGroups.Item(ColumnGroup.DATA_TYPE).ToString)

            column = New Column(columnName, dataType)
            columns.Add(column)
        Next

        Return columns
    End Function

    Private Function GetDataType(ByVal dataTypeString As String, Optional ByVal arguments As String = "") As DataType
        Dim dataType As DataType._Type = DirectCast([Enum].Parse(GetType(DataType), dataTypeString), Integer)
        Return New DataType(dataType, arguments)
    End Function
End Class
