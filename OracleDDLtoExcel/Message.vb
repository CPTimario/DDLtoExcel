Module Message
    '-----------
    ' Constants
    '-----------
    Public Const INPUT As String = "Please input {0}."
    Public Const CHOOSE As String = "Please choose {0}."
    Public Const IDLE As String = "Idle"
    Public Const SUCCESS As String = "Success"
    Public Const REMOVE_COMMENTS As String = "Removing comments ({0} of {1} {2}) . . ."

    '---------
    ' Methods
    '---------
    Public Sub Show(ByVal pMessage As String, ByVal pIcon As MessageBoxIcon, ByVal pButtons As MessageBoxButtons, ByVal pControl As Control, Optional ByVal pMessageArgs As String() = Nothing)
        Dim message As String = CreateMessage(pMessage, pMessageArgs)

        MessageBox.Show(message, pIcon.ToString("F"), pButtons, pIcon)
        pControl.Focus()
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByRef pStatusLabel As Label, ByVal pStatus As String, Optional ByVal pMessageArgs As String() = Nothing)
        pStatusLabel.Text = CreateMessage(pStatus, pMessageArgs)
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByRef pStatusLabel As Label, ByVal pStatus As String, ByRef pProgressBar As ProgressBar, ByVal pMinimum As Integer, ByVal pMaximum As Integer)
        pStatusLabel.Text = pStatus
        pProgressBar.Minimum = pMinimum
        pProgressBar.Maximum = pMaximum
        pProgressBar.Value = 0
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByRef pStatusLabel As Label, ByVal pStatus As String, ByRef pProgressBar As ProgressBar, ByVal pValue As Integer, Optional ByVal pMessageArgs As String() = Nothing)
        pStatusLabel.Text = CreateMessage(pStatus, pMessageArgs)
        pProgressBar.Value = pValue
        Application.DoEvents()
    End Sub

    '-----------
    ' Functions
    '-----------
    Private Function CreateMessage(ByVal pMessage As String, ByVal pMessageArgs As String())
        Dim message As String = pMessage

        If pMessageArgs IsNot Nothing Then
            For intIdx As Integer = 0 To pMessageArgs.Length - 1
                message = message.Replace("{" & intIdx & "}", pMessageArgs(intIdx))
            Next
        End If

        Return message
    End Function
End Module
