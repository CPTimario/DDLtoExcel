Imports System.Text.RegularExpressions
Imports DDLtoExcel.SQL

Public Class Schema
    Public Name As String
    Public FileName As String
    Public Tables As List(Of Table)
    Public SQLCommands As List(Of String)

    Public Sub New(ByVal fileName As String, ByVal command As String)
        Me.FileName = fileName
        Tables = New List(Of Table)
        command = Regex.Replace(command, SQLRegex.SQL_COMMENT, String.Empty, RegexOptions.IgnoreCase)
        SQLCommands = GetSQLCommands(command)
    End Sub

    Private Function GetSQLCommands(ByVal sqlCommand As String) As List(Of String)
        Dim sqlCommandList As List(Of String)
        Dim ddlCommandString As String
        Dim commandRegex As String
        Dim replaceString As String

        For Each command As DDLCommand In [Enum].GetValues(GetType(DDLCommand))
            If sqlCommand.Contains(command.EnumToString()) Then
                ddlCommandString = command.EnumToString()
                commandRegex = command.ToRegex(String.Empty, "\s+")
                replaceString = Chr(36) & ddlCommandString & Chr(32)
                sqlCommand = Regex.Replace(sqlCommand, commandRegex, replaceString, RegexOptions.IgnoreCase)
            End If
        Next

        sqlCommandList = sqlCommand.Split(Chr(36)).ToList
        sqlCommandList.Remove(String.Empty)

        Return sqlCommandList
    End Function

    Public Function Table(ByVal tableName As String) As Table
        Return Tables.Find(Function(findTable) findTable.Name = tableName)
    End Function
End Class
