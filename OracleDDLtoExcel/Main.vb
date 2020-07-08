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
        Dim SQL As SQL

        If txtTitle.Text = String.Empty Then
            Message.Show(Message.INPUT, MessageBoxIcon.Error, MessageBoxButtons.OK, txtTitle, New String() {"title"})
        ElseIf txtFile.Text = String.Empty Then
            Message.Show(Message.CHOOSE, MessageBoxIcon.Error, MessageBoxButtons.OK, txtFile, New String() {"file"})
        Else
            CancelFlg = False
            Call EnableFormComponents(False)

            content = File.ReadAllText(txtFile.Text)

            SQL = New SQL(content)

            Call SQL.RemoveComments()
            If CancelFlg Then Exit Sub

            Call SQL.ExecuteCommands()

            Call ShowStatus(SUCCESS)
            Call EnableFormComponents(True)
            timDelayIdleMessage.Start()
        End If
    End Sub

    Private Sub timDelayIdleMessage_Tick(sender As Object, e As EventArgs) Handles timDelayIdleMessage.Tick
        timDelayIdleMessage.Stop()
        Call ShowStatus(IDLE, 0)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        CancelFlg = True
        Call EnableFormComponents(True)

        Call ShowStatus(CANCEL)
        timDelayIdleMessage.Start()
    End Sub

    Private Sub EnableFormComponents(ByVal value As Boolean)
        btnExport.Enabled = value
        btnOpen.Enabled = value
        txtTitle.Enabled = value
    End Sub
End Class
