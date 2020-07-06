Module Message
    '-----------
    ' Constants
    '-----------
    Public Const INPUT As String = "Please input {0}."
    Public Const CHOOSE As String = "Please choose {0}."
    Public Const IDLE As String = "Idle"
    Public Const SUCCESS As String = "Success"
    Public Const REMOVE_COMMENTS As String = "Removing comments ({0} of {1} comments) . . ."
    Public Const GET_COMMANDS As String = "Getting DDL commands ({0} of {1} commands) . . ."

    '---------
    ' Methods
    '---------
    Public Sub Show(ByVal pMessage As String, ByVal pIcon As MessageBoxIcon, ByVal pButtons As MessageBoxButtons, ByVal pControl As Control, Optional ByVal pMessageArgs As String() = Nothing)
        Dim message As String = FormatMessage(pMessage, pMessageArgs)

        MessageBox.Show(message, pIcon.ToString("F"), pButtons, pIcon)
        pControl.Focus()
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByVal pStatus As String, Optional ByVal pMessageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(pStatus, pMessageArgs)
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByVal pStatus As String, ByVal pMinimum As Integer, ByVal pMaximum As Integer, Optional ByVal pMessageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(pStatus, pMessageArgs)
        Main.pbProgress.Minimum = pMinimum
        Main.pbProgress.Maximum = pMaximum
        Main.pbProgress.Value = 0
        Application.DoEvents()
    End Sub

    Public Sub ShowStatus(ByVal pStatus As String, ByVal pValue As Integer, Optional ByVal pMessageArgs As String() = Nothing)
        Main.lblStatus.Text = FormatMessage(pStatus, pMessageArgs)
        Main.pbProgress.Value = pValue
        Application.DoEvents()
    End Sub

    '-----------
    ' Functions
    '-----------
    Private Function FormatMessage(ByVal pMessage As String, ByVal pMessageArgs As String())
        Dim message As String = pMessage

        If pMessageArgs IsNot Nothing Then
            message = String.Format(pMessage, pMessageArgs)
        End If

        Return message
    End Function
End Module
