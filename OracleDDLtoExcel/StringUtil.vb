Module StringUtil
    Public Function Substring(ByVal pString As String, ByVal pStart As String, ByVal pEnd As String) As String
        Dim resultString As String = String.Empty
        Dim intStartIdx As Integer = -1
        Dim intEndIdx As Integer = -1

        intStartIdx = pString.IndexOf(pStart)
        If intStartIdx > -1 Then
            intEndIdx = pString.IndexOf(pEnd, intStartIdx + pStart.Length)
        End If

        If intStartIdx > -1 And intEndIdx > -1 Then
            resultString = pString.Substring(intStartIdx, intEndIdx - intStartIdx + pEnd.Length)
        End If

        Return resultString
    End Function

    Public Function RemoveSubstring(ByVal pString As String, ByVal pStart As String, ByVal pEnd As String) As String
        Dim resultString As String = pString

        While Not Substring(resultString, pStart, pEnd) = String.Empty
            resultString = resultString.Replace(Substring(resultString, pStart, pEnd), String.Empty)
        End While

        Return resultString
    End Function
End Module
