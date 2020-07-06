Module CommandExtraction
    '---------
    ' Methods
    '---------
    Public Sub RemoveComments(ByRef value As String)
        Dim commentCount As Integer = 0
        Dim progressValue As Integer = 0
        Dim subString As String

        For Each comment As SQlComment In SqlComments
            commentCount += value.SubstringCount(comment.StartString, comment.EndString)
        Next

        Call ShowStatus(REMOVE_COMMENTS, progressValue, commentCount, New String() {progressValue.ToString, commentCount.ToString})

        For Each sqlComment As SQlComment In SqlComments
            While Not value.Substring(sqlComment.StartString, sqlComment.EndString) = String.Empty
                subString = value.Substring(sqlComment.StartString, sqlComment.EndString)
                progressValue += value.SubstringCount(subString)
                value = value.Trim().Replace(subString, String.Empty)

                Call ShowStatus(REMOVE_COMMENTS, progressValue, New String() {progressValue.ToString, commentCount.ToString})
                Application.DoEvents()
            End While
        Next
    End Sub

    '------------
    ' Functions
    '------------
    Public Function GetSQLCommands(ByVal commandString As String) As List(Of String)
        Dim sqlCommands As List(Of String)
        Dim command As String

        For Each ddlCommand As DDLCommand In System.Enum.GetValues(GetType(DDLCommand))
            command = DDLCommands.Item(ddlCommand)
            commandString = commandString.Replace(command, Chr(36) & command)
        Next

        sqlCommands = commandString.Split(Chr(36)).ToList

        Return sqlCommands
    End Function
End Module
