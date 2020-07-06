Imports System.IO

Public Class Main
    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        If ofdOpenFile.ShowDialog = DialogResult.OK Then
            txtFile.Text = ofdOpenFile.FileName

            If txtTitle.Text = String.Empty Then
                txtTitle.Text = Path.GetFileNameWithoutExtension(ofdOpenFile.FileName)
            End If
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim content As String
        Dim sqlCommands As New List(Of String)

        If txtTitle.Text = String.Empty Then
            Message.Show(Message.INPUT, MessageBoxIcon.Error, MessageBoxButtons.OK, txtTitle, New String() {"title"})
        ElseIf txtFile.Text = String.Empty Then
            Message.Show(Message.CHOOSE, MessageBoxIcon.Error, MessageBoxButtons.OK, txtFile, New String() {"file"})
        Else
            CancelFlg = False
            btnExport.Enabled = False

            content = File.ReadAllText(txtFile.Text)
            Call RemoveComments(content.Trim)
            sqlCommands = GetSQLCommands(content)

            Call ShowStatus(SUCCESS)
            timDelayIdleMessage.Start()
        End If
    End Sub

    Private Sub timDelayIdleMessage_Tick(sender As Object, e As EventArgs) Handles timDelayIdleMessage.Tick
        timDelayIdleMessage.Stop()
        Call ShowStatus(IDLE, 0)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        'CancelFlg = True
        'Call ShowStatus(CANCEL)
        'timDelayIdleMessage.Start()
    End Sub
End Class
