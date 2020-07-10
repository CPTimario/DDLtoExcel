Module Message
    '-----------
    ' Constants
    '-----------
    Public Const INPUT As String = "Please input {0}."
    Public Const CHOOSE As String = "Please choose {0}."
    Public Const IDLE As String = "Idle"
    Public Const SUCCESS As String = "Success"
    Public Const CANCEL As String = "Cancelled"
    Public Const REMOVE_COMMENTS As String = "Removing comments ({0} of {1} comments) . . ."
    Public Const EXECUTE_COMMANDS As String = "Executing commands ({0} of {1} commands) . . ."

    '---------
    ' Methods
    '---------
    Public Sub Show(ByVal messageSource As String, ByVal icon As MessageBoxIcon, ByVal buttons As MessageBoxButtons, ByVal focusControl As Control, Optional ByVal messageArgs As String() = Nothing)
        Dim message As String = FormatMessage(messageSource, messageArgs)

        MessageBox.Show(message, icon.ToString("F"), buttons, icon)
        focusControl.Focus()
    End Sub

    Public Sub ShowStatus(ByVal messageSource As String, Optional ByVal messageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(messageSource, messageArgs)
        Main.pbProgress.Value = 0
    End Sub

    Public Sub ShowStatus(ByVal messageSource As String, ByVal progressMinimum As Integer, ByVal progressMaximum As Integer, Optional ByVal messageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(messageSource, messageArgs)
        Main.pbProgress.Minimum = progressMinimum
        Main.pbProgress.Maximum = progressMaximum
        Main.pbProgress.Value = 0
    End Sub

    Public Sub ShowStatus(ByVal messageSource As String, ByVal progressValue As Integer, Optional ByVal messageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(messageSource, messageArgs)
        Main.pbProgress.Value = progressValue
    End Sub

    '-----------
    ' Functions
    '-----------
    Private Function FormatMessage(ByVal messageSource As String, ByVal messageArgs As String())
        Dim message As String = messageSource

        If messageArgs IsNot Nothing Then
            message = String.Format(messageSource, messageArgs)
        End If

        Return message
    End Function
End Module
