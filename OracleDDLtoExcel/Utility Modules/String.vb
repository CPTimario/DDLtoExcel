Imports System.Runtime.CompilerServices

Module StringUtils
    '------------
    ' Functions
    '------------
    <Extension>
    Public Function Contains(ByVal value As String, ByVal searchStringStart As String, Optional ByVal searchStringEnd As String = "", Optional ByVal offsetIndex As Integer = 0) As Boolean
        Dim startIndex As Integer
        Dim endIndex As Integer

        startIndex = value.IndexOf(searchStringStart, offsetIndex)
        If searchStringEnd = String.Empty Then
            endIndex = 0
        ElseIf startIndex > -1 Then
            endIndex = value.IndexOf(searchStringEnd, startIndex + searchStringEnd.Length)
        End If

        Return startIndex > -1 AndAlso endIndex > -1
    End Function

    <Extension>
    Public Function Substring(ByVal value As String, ByVal searchStringStart As String, ByVal searchStringEnd As String) As String
        If value.Contains(searchStringStart, searchStringEnd) Then
            Dim startIndex As Integer = value.IndexOf(searchStringStart)
            Dim endIndex As Integer = value.IndexOf(searchStringEnd, startIndex + searchStringStart.Length)

            Return value.Substring(startIndex, endIndex - startIndex + searchStringEnd.Length)
        End If

        Return String.Empty
    End Function

    '<Extension>
    'Public Function Substring(ByVal value As String, ByVal stringPair As StringPair)
    '    Dim indexStack As New Stack(Of Integer)
    '    Dim startIndex As Integer = value.IndexOf(stringPair.StartString)
    '    Dim length As Integer = 0

    '    For index As Integer = startIndex To value.Length - 1
    '        length += 1
    '        If value(index).ToString.Equals(stringPair.StartString) Then
    '            indexStack.Push(index)
    '        ElseIf value(index).ToString.Equals(stringPair.EndString) Then
    '            indexStack.Pop()

    '            If indexStack.Count = 0 Then
    '                Exit For
    '            End If
    '        End If
    '    Next

    '    If startIndex > -1 And length > 0 Then
    '        Return value.Substring(startIndex, length)
    '    End If

    '    Return String.Empty
    'End Function

    <Extension>
    Public Function SubstringCount(ByVal value As String, ByVal searchStringStart As String, Optional ByVal searchStringEnd As String = "") As Integer
        Dim count As Integer = 0
        Dim offsetIndex As Integer = 0
        Dim startIndex As Integer
        Dim endIndex As Integer

        While value.Contains(searchStringStart, searchStringEnd, offsetIndex)
            count += 1
            startIndex = value.IndexOf(searchStringStart, offsetIndex)

            If searchStringEnd = String.Empty Then
                offsetIndex = startIndex + searchStringStart.Length
            Else
                endIndex = value.IndexOf(searchStringEnd, startIndex + searchStringStart.Length)
                offsetIndex = endIndex + searchStringEnd.Length
            End If
        End While

        Return count
    End Function
End Module
