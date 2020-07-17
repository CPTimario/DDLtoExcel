Imports System.IO

Public Class Main
    Private Const ADJUST_HEIGHT As Integer = 32
    Private controlsToAdjust As List(Of Control)
    Private schemas As List(Of SchemaFormControls)

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        controlsToAdjust = New List(Of Control)({btnAddSchema, btnExport, btnCancel, pbProgress, lblStatus})
        schemas = New List(Of SchemaFormControls)({New SchemaFormControls(lblSchema, txtSchema, btnOpenSchema, btnRemoveSchema)})
    End Sub

    Private Sub btnOpenSchema_Click(sender As Button, e As EventArgs) Handles btnOpenSchema.Click
        Dim textbox As TextBox = Controls.Item(sender.Name.Replace("btnOpen", "txt"))
        If ofdOpenFile.ShowDialog = DialogResult.OK Then
            textbox.Text = ofdOpenFile.FileName
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim SQL As SQL
        Dim Excel As Excel
        Dim schemaName As String
        Dim command As String

        If Not InputsEmpty() Then
            CancelFlg = False
            Call EnableFormComponents(False)

            SQL = New SQL()

            For Each schema As SchemaFormControls In schemas
                schemaName = Path.GetFileName(schema.Textbox.Text)
                command = File.ReadAllText(schema.Textbox.Text)
                SQL.Schemas.Add(New Schema(schemaName, command))
            Next

            Call SQL.ExecuteCommands()
            If CancelFlg Then Exit Sub

            Excel = New Excel(SQL.Schemas, txtTitle.Text)
            Call Excel.Export()
            If CancelFlg Then Exit Sub

            Call ShowStatus(SUCCESS)
            Call EnableFormComponents(True)
            timDelayIdleMessage.Start()
        End If
    End Sub

    Private Sub timDelayIdleMessage_Tick(sender As Object, e As EventArgs) Handles timDelayIdleMessage.Tick
        timDelayIdleMessage.Stop()
        Call ShowStatus(IDLE, 0)
        Call EnableFormComponents(True)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If Not CancelFlg Then
            CancelFlg = True
            Call ShowStatus(CANCEL)
            Call EnableFormComponents(True)
            timDelayIdleMessage.Start()
        End If
    End Sub

    Private Sub btnAddSchema_Click(sender As Object, e As EventArgs) Handles btnAddSchema.Click
        Dim label As New Label
        Dim textbox As New TextBox
        Dim open As New Button
        Dim remove As New Button

        label.Location = NewLocation(schemas.Last.Label.Location, ADJUST_HEIGHT)
        label.AutoSize = schemas.Last.Label.AutoSize
        label.Text = schemas.Last.Label.Text
        label.Font = schemas.Last.Label.Font

        textbox.Location = NewLocation(schemas.Last.Textbox.Location, ADJUST_HEIGHT)
        textbox.Size = schemas.Last.Textbox.Size
        textbox.ReadOnly = schemas.Last.Textbox.ReadOnly

        open.Location = NewLocation(schemas.Last.Open.Location, ADJUST_HEIGHT)
        open.Text = schemas.Last.Open.Text
        open.AutoSize = schemas.Last.Open.AutoSize
        open.AutoSizeMode = schemas.Last.Open.AutoSizeMode
        AddHandler open.Click, AddressOf btnOpenSchema_Click

        remove.Location = NewLocation(schemas.Last.Remove.Location, ADJUST_HEIGHT)
        remove.BackgroundImage = schemas.Last.Remove.BackgroundImage
        remove.BackgroundImageLayout = schemas.Last.Remove.BackgroundImageLayout
        remove.Size = schemas.Last.Remove.Size
        AddHandler remove.Click, AddressOf btnRemoveSchema_Click

        schemas.Add(New SchemaFormControls(label, textbox, open, remove))
        Call AdjustSchemaControls()

        Call AdjustForm(ADJUST_HEIGHT)
        Controls.Add(label)
        Controls.Add(textbox)
        Controls.Add(open)
        Controls.Add(remove)
    End Sub

    Private Sub btnRemoveSchema_Click(sender As Button, e As EventArgs) Handles btnRemoveSchema.Click
        Dim index As Integer = sender.Name.Replace("btnRemoveSchema", String.Empty)

        Controls.RemoveByKey("lblSchema" & index)
        Controls.RemoveByKey("txtSchema" & index)
        Controls.RemoveByKey("btnOpenSchema" & index)
        Controls.RemoveByKey("btnRemoveSchema" & index)

        schemas.RemoveAt(index)
        Call AdjustSchemaControls()

        Call AdjustForm(-ADJUST_HEIGHT)
    End Sub

    Private Sub EnableFormComponents(ByVal value As Boolean)
        CancelFlg = value
        btnExport.Enabled = value
        btnAddSchema.Enabled = value
        txtTitle.Enabled = value

        For Each schema As SchemaFormControls In schemas
            schema.Open.Enabled = value
            schema.Remove.Enabled = value
        Next
    End Sub

    Private Sub AdjustForm(ByVal height As Integer)
        Size = New Size(Size.Width, Size.Height + height)
        For Each control As Control In controlsToAdjust
            control.Location = NewLocation(control.Location, height)
        Next
    End Sub

    Private Function NewLocation(ByVal location As Point, ByVal height As Integer) As Point
        Return New Point(location.X, location.Y + height)
    End Function

    Private Sub AdjustSchemaControls()
        Dim height As Integer
        For Each schema As SchemaFormControls In schemas
            height = (ADJUST_HEIGHT * schemas.IndexOf(schema))

            schema.Label.Name = "lblSchema" & schemas.IndexOf(schema)
            schema.Label.Location = NewLocation(New Point(14, 105), height)

            schema.Textbox.Name = "txtSchema" & schemas.IndexOf(schema)
            schema.Textbox.Location = NewLocation(New Point(99, 102), height)

            schema.Open.Name = "btnOpenSchema" & schemas.IndexOf(schema)
            schema.Open.Location = NewLocation(New Point(542, 100), height)

            schema.Remove.Name = "btnRemoveSchema" & schemas.IndexOf(schema)
            schema.Remove.Enabled = schemas.Count > 1
            schema.Remove.Location = NewLocation(New Point(606, 100), height)
        Next
    End Sub

    Private Function InputsEmpty() As Boolean
        If txtTitle.Text.Equals(String.Empty) Then
            Call ShowMessage(INPUT_ERROR, MessageBoxIcon.Error, MessageBoxButtons.OK, New String() {"title"}, txtTitle)
            Return True
        End If

        For Each schema As SchemaFormControls In schemas
            If schema.Textbox.Text.Equals(String.Empty) Then
                Call ShowMessage(INPUT_ERROR, MessageBoxIcon.Error, MessageBoxButtons.OK, New String() {"schema"}, schema.Textbox)
                Return True
            End If
        Next

        Return False
    End Function
End Class
