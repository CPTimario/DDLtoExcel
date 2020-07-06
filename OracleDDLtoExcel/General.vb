Imports System.Reflection

Public Module General
    '------------
    ' Structures
    '------------
    Public Structure SQLComment
        Public Const SINGLE_START As String = "--"
        Public Const SINGLE_END As String = vbLf
        Public Const MULTI_START As String = "/*"
        Public Const MULTI_END As String = "*/"
    End Structure

    Public Structure EnumPrefix
        Public Const CONSTRAINT_TYPE As String = "ct"
        Public Const DATA_TYPE As String = "dt"
        Public Const DDL_COMMAND As String = "ddl"
    End Structure

    Public Structure Constraint
        Dim Type As ConstraintType
        Dim Expression As String
        Dim Reference As KeyValuePair(Of String, String)

        Public Sub New(ByVal pType As ConstraintType)
            Type = pType
        End Sub

        Public Sub New(ByVal pType As ConstraintType, ByVal pRefTable As String, ByVal pRefColumn As String)
            Type = pType
            Reference = New KeyValuePair(Of String, String)(pRefTable, pRefColumn)
        End Sub

        Public Sub New(ByVal pType As ConstraintType, ByVal pExpression As String)
            Type = pType
            Expression = pExpression
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

    '---------
    ' Methods
    '---------
    Public Sub RemoveComments(ByRef pString As String)
        Dim intCommentCount As Integer = 0
        Dim intProgress As Integer = 0
        Dim subString As String = String.Empty

        intCommentCount = SubstringCount(pString, SQLComment.SINGLE_START, SQLComment.SINGLE_END)
        intCommentCount += SubstringCount(pString, SQLComment.MULTI_START, SQLComment.MULTI_END)

        Call ShowStatus(REMOVE_COMMENTS, intProgress, intCommentCount, New String() {intProgress.ToString, intCommentCount.ToString})

        While Not pString.Substring(SQLComment.SINGLE_START, SQLComment.SINGLE_END) = String.Empty
            subString = pString.Substring(SQLComment.SINGLE_START, SQLComment.SINGLE_END)
            intProgress += pString.SubstringCount(subString)
            pString = pString.Replace(subString, String.Empty).Trim()
            Call ShowStatus(REMOVE_COMMENTS, intProgress, New String() {intProgress.ToString, intCommentCount.ToString})
        End While

        While Not pString.Substring(SQLComment.MULTI_START, SQLComment.MULTI_END) = String.Empty
            subString = pString.Substring(SQLComment.MULTI_START, SQLComment.MULTI_END)
            intProgress += pString.SubstringCount(subString)
            pString = pString.Replace(subString, String.Empty).Trim()
            Call ShowStatus(REMOVE_COMMENTS, intProgress, New String() {intProgress.ToString, intCommentCount.ToString})
        End While
    End Sub

    '------------
    ' Functions
    '------------
    Public Function GetSQLCommands(ByVal pCommandString As String) As List(Of String)
        Dim sqlCommands As New List(Of String)
        Dim intProgress As Integer = 0
        Dim intCommandCount As Integer = 0
        Dim intIdx As Integer = 0
        Dim strCommand As String = String.Empty
        Dim curCommand As KeyValuePair(Of String, Integer)
        Dim nxtCommand As KeyValuePair(Of String, Integer)

        For Each command As DDLCommand In System.Enum.GetValues(GetType(DDLCommand))
            strCommand = command.ToString.EnumToString(EnumPrefix.DDL_COMMAND) & Chr(32)
            intCommandCount += pCommandString.SubstringCount(strCommand)
        Next

        curCommand = New KeyValuePair(Of String, Integer)(String.Empty, pCommandString.Length)
        nxtCommand = New KeyValuePair(Of String, Integer)(String.Empty, pCommandString.Length)

        Call ShowStatus(GET_COMMANDS, intProgress, intCommandCount, New String() {intProgress.ToString, intCommandCount.ToString})

        For Each command As DDLCommand In System.Enum.GetValues(GetType(DDLCommand))
            strCommand = command.ToString.EnumToString(EnumPrefix.DDL_COMMAND) & Chr(32)
            intIdx = pCommandString.IndexOf(strCommand)

            If intIdx > -1 AndAlso intIdx < curCommand.Value Then
                curCommand = New KeyValuePair(Of String, Integer)(strCommand, intIdx)
            End If
        Next

        Do
            For Each command As DDLCommand In System.Enum.GetValues(GetType(DDLCommand))
                strCommand = command.ToString.EnumToString(EnumPrefix.DDL_COMMAND) & Chr(32)
                intIdx = pCommandString.IndexOf(strCommand, curCommand.Value + curCommand.Key.Length)

                If intIdx > -1 AndAlso intIdx < nxtCommand.Value Then
                    nxtCommand = New KeyValuePair(Of String, Integer)(strCommand, intIdx)
                End If
            Next

            intProgress += 1
            sqlCommands.Add(pCommandString.Substring(curCommand.Value, nxtCommand.Value - curCommand.Value - 1).Trim())
            curCommand = nxtCommand
            nxtCommand = New KeyValuePair(Of String, Integer)(String.Empty, pCommandString.Length)

            Call ShowStatus(GET_COMMANDS, intProgress, New String() {intProgress.ToString, intCommandCount.ToString})
        Loop Until intProgress = intCommandCount

        Return sqlCommands
    End Function
End Module
