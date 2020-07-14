Module Message
    '-----------
    ' Constants
    '-----------
    Public Const INPUT_ERROR As String = "Please input {0}."
    Public Const CHOOSE_ERROR As String = "Please choose {0}."
    Public Const IDLE As String = "Idle"
    Public Const SUCCESS As String = "Success"
    Public Const CANCEL As String = "Cancelled"
    Public Const EXECUTE_COMMANDS As String = "Executing commands ({0} of {1} commands) . . ."
    Public Const CREATING_SUMMARY_SHEET As String = "Creating summary sheet ({0} of {1} tables) . . ."
    Public Const CREATING_TABLE_SHEET As String = "Creating sheet for table {0} ({1} of {2} tables) . . ."

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
        If messageArgs IsNot Nothing Then
            Return String.Format(messageSource, messageArgs)
        End If

        Return messageSource
    End Function
End Module
