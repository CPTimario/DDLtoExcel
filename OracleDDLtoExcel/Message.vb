Module Message
    '-----------
    ' Constants
    '-----------
    Public Const INPUT = "Please input {0}."
    Public Const CHOOSE = "Please choose {0}."

    '---------
    ' Methods
    '---------
    Public Sub Show(ByVal pMessage As String, ByVal pIcon As MessageBoxIcon, ByVal pButtons As MessageBoxButtons, ByVal pControl As Control, Optional ByVal pMessageArgs As String() = Nothing)
        Dim message As String = pMessage

        For intIdx As Integer = 0 To pMessageArgs.Length - 1
            message = message.Replace("{" & intIdx & "}", pMessageArgs(intIdx))
        Next

        MessageBox.Show(message, pIcon.ToString("F"), pButtons, pIcon)
        pControl.Focus()
    End Sub
End Module
