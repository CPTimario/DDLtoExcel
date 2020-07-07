Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Module SQLCommand
    '---------
    ' Methods
    '---------
    <Extension>
    Public Sub ExecuteCommands(ByVal commands As List(Of String))
        Dim ddlCommand As DDLCommand

        For Each command As String In commands
            ddlCommand = command.GetDDLCommandType()

            Select Case ddlCommand
                Case DDLCommand.ddlCREATE
                    If command.Contains(DDLCommands(ddlCommand) & "TABLE ") Then
                        Call CreateTable(command)
                    End If
                Case DDLCommand.ddlALTER
                    'ALTER
                Case DDLCommand.ddlDROP
                    'DROP
                Case DDLCommand.ddlCOMMENT_ON_COLUMN
                    'COMMENT ON
                Case Else
                    Continue For
            End Select
        Next
    End Sub

    Public Sub CreateTable(ByVal command As String)
        Dim tableName As String = command.Substring(Chr(34), Chr(34))
        Dim table As New Table(tableName.Replace(Chr(34), String.Empty))
        Dim columns As List(Of Column) = command.GetColumns()
    End Sub
    '------------
    ' Functions
    '------------
    <Extension>
    Public Function RemoveComments(ByVal value As String) As String
        RemoveComments = value

        Dim commentCount As Integer = 0
        Dim progressValue As Integer = 0
        Dim subString As String

        For Each comment As StringPair In SqlComments
            commentCount += value.SubstringCount(comment.StartString, comment.EndString)
        Next

        Call ShowStatus(REMOVE_COMMENTS, progressValue, commentCount, New String() {progressValue.ToString, commentCount.ToString})

        For Each sqlComment As StringPair In SqlComments
            If CancelFlg Then Exit Function

            While Not value.Substring(sqlComment.StartString, sqlComment.EndString) = String.Empty
                If CancelFlg Then Exit Function

                subString = value.Substring(sqlComment.StartString, sqlComment.EndString)
                progressValue += value.SubstringCount(subString)
                value = value.Replace(subString, String.Empty).Trim()

                Call ShowStatus(REMOVE_COMMENTS, progressValue, New String() {progressValue.ToString, commentCount.ToString})
                Application.DoEvents()
            End While
        Next

        Return value
    End Function

    <Extension>
    Public Function GetSQLCommands(ByVal commandString As String) As List(Of String)
        Dim sqlCommands As List(Of String)

        For Each command As DDLCommand In DDLCommands.Keys
            Dim ddlCommandString As String

            If command = DDLCommand.ddlCREATE OrElse command = DDLCommand.ddlDROP Then
                For Each createDropObject As DDLCreateDropObject In DDLCreateDropObjects.Keys
                    ddlCommandString = DDLCommands(command) & DDLCreateDropObjects(createDropObject)
                    commandString = commandString.Replace(ddlCommandString, Chr(36) & ddlCommandString)
                Next
            ElseIf command = DDLCommand.ddlALTER Then
                For Each createDropObject As DDLCreateDropObject In DDLCreateDropObjects.Keys
                    ddlCommandString = DDLCommands(command) & DDLCreateDropObjects(createDropObject)
                    commandString = commandString.Replace(ddlCommandString, Chr(36) & ddlCommandString)
                Next
            Else
                ddlCommandString = DDLCommands(command)
                commandString = commandString.Replace(ddlCommandString, Chr(36) & ddlCommandString)
            End If
        Next

        sqlCommands = commandString.Split(Chr(36)).ToList
        sqlCommands.Remove(String.Empty)

        Return sqlCommands
    End Function

    <Extension>
    Private Function GetDDLCommandType(ByVal commandString As String) As DDLCommand
        GetDDLCommandType = Nothing

        For Each command As DDLCommand In DDLCommands.Keys
            If commandString.Contains(DDLCommands(command)) Then
                Return command
            End If
        Next
    End Function

    <Extension>
    Private Function GetColumns(ByVal commandString As String) As List(Of Column)
        GetColumns = New List(Of Column)

        Dim columnClause As String = commandString.Substring(Parenthesis)
        Dim columnStrings As List(Of String) = Regex.Split(columnClause, COLUMN_SYNTAX).ToList
    End Function
End Module
