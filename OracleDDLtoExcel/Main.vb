Imports System.IO

Public Class Main
    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        If ofdOpenFile.ShowDialog = DialogResult.OK Then
            txtFile.Text = ofdOpenFile.FileName
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim textContent As String = String.Empty

        If txtTitle.Text = String.Empty Then
            Message.Show(Message.INPUT, MessageBoxIcon.Error, MessageBoxButtons.OK, txtTitle, New String() {"title"})
        ElseIf txtFile.Text = String.Empty Then
            Message.Show(Message.CHOOSE, MessageBoxIcon.Error, MessageBoxButtons.OK, txtFile, New String() {"file"})
        Else
            textContent = File.ReadAllText(txtFile.Text)
            Call RemoveComments(textContent)
        End If
    End Sub

    Public Sub RemoveComments(ByRef pString As String)
        Dim intCommentCount As Integer = 0
        Dim intProgress As Integer = 0

        intCommentCount = SubstringCount(pString, SINGLE_COMMENT_START, SINGLE_COMMENT_END)
        intCommentCount += SubstringCount(pString, MULTI_COMMENT_START, MULTI_COMMENT_END)

        Call ShowStatus(lblStatus, REMOVE_COMMENTS, pbProgress, intProgress, intCommentCount)

        While Not pString.Substring(SINGLE_COMMENT_START, SINGLE_COMMENT_END) = String.Empty
            intProgress += 1
            Call pString.RemoveSubstring(SINGLE_COMMENT_START, SINGLE_COMMENT_END)
            Call ShowStatus(lblStatus, REMOVE_COMMENTS, pbProgress, intProgress, New String() {intProgress.ToString, intCommentCount.ToString, "comments"})
        End While

        While Not pString.Substring(MULTI_COMMENT_START, MULTI_COMMENT_END) = String.Empty
            intProgress += 1
            Call pString.RemoveSubstring(MULTI_COMMENT_START, MULTI_COMMENT_END)
            Call ShowStatus(lblStatus, REMOVE_COMMENTS, pbProgress, intProgress, New String() {intProgress.ToString, intCommentCount.ToString, "comments"})
        End While

        Call ShowStatus(lblStatus, SUCCESS)
        timDelayIdleMessage.Start()
    End Sub

    Private Sub timDelayIdleMessage_Tick(sender As Object, e As EventArgs) Handles timDelayIdleMessage.Tick
        timDelayIdleMessage.Stop()
        Call ShowStatus(lblStatus, IDLE, pbProgress, 0)
    End Sub
End Class
