Imports System.Runtime.CompilerServices

Module StringUtil
    '------------
    ' Functions
    '------------
    <Extension()>
    Public Function Contains(ByVal pString As String, ByVal pStartString As String, ByVal pEndString As String) As Boolean
        Dim intStartIdx As Integer = -1
        Dim intEndIdx As Integer = -1

        intStartIdx = pString.IndexOf(pStartString)
        If intStartIdx > -1 Then
            intEndIdx = pString.IndexOf(pEndString, intStartIdx + pEndString.Length)
        End If

        Return intStartIdx > -1 And intEndIdx > -1
    End Function

    <Extension()>
    Public Function Substring(ByVal pString As String, ByVal pStartString As String, ByVal pEndString As String) As String
        Dim resultString As String = String.Empty
        Dim intStartIdx As Integer = -1
        Dim intEndIdx As Integer = -1

        If pString.Contains(pStartString, pEndString) Then
            intStartIdx = pString.IndexOf(pStartString)
            intEndIdx = pString.IndexOf(pEndString, intStartIdx + pStartString.Length)
            resultString = pString.Substring(intStartIdx, intEndIdx - intStartIdx + pEndString.Length)
        End If

        Return resultString
    End Function

    <Extension()>
    Public Function SubstringCount(ByVal pString As String, ByVal pStartString As String, ByVal pEndString As String) As Integer
        Dim intCount As Integer = 0
        Dim intOffsetIdx As Integer = 0
        Dim intStartIdx As Integer = -1
        Dim intEndIdx As Integer = -1

        Do
            intStartIdx = pString.IndexOf(pStartString, intOffsetIdx)
            If intStartIdx > -1 Then
                intEndIdx = pString.IndexOf(pEndString, intStartIdx + pStartString.Length)
            End If

            If intStartIdx > -1 And intEndIdx > -1 Then
                intCount += 1
                intOffsetIdx = intEndIdx + pEndString.Length
            End If
        Loop While intStartIdx > -1

        Return intCount
    End Function

    <Extension()>
    Public Function SubstringCount(ByVal pString As String, ByVal pSearchString As String) As Integer
        Dim intCount As Integer = 0
        Dim intOffsetIdx As Integer = 0
        Dim intStartIdx As Integer = -1

        Do
            intStartIdx = pString.IndexOf(pSearchString, intOffsetIdx)

            If intStartIdx > -1 Then
                intCount += 1
                intOffsetIdx = intStartIdx + pSearchString.Length
            End If
        Loop While intStartIdx > -1

        Return intCount
    End Function

    <Extension()>
    Public Function EnumToString(ByVal pEnumName As String, pEnumPrefix As String)
        Return pEnumName.Replace(pEnumPrefix, String.Empty).Replace("_", String.Empty)
    End Function
End Module
