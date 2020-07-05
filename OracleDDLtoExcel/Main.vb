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
            textContent = RemoveComments(textContent)
        End If
    End Sub
End Class
