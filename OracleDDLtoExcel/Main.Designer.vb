<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lblAppTitle = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.lblSchema = New System.Windows.Forms.Label()
        Me.txtSchema = New System.Windows.Forms.TextBox()
        Me.btnOpenSchema = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ofdOpenFile = New System.Windows.Forms.OpenFileDialog()
        Me.pbProgress = New System.Windows.Forms.ProgressBar()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.timDelayIdleMessage = New System.Windows.Forms.Timer(Me.components)
        Me.btnAddSchema = New System.Windows.Forms.Button()
        Me.btnRemoveSchema = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblAppTitle
        '
        Me.lblAppTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAppTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblAppTitle.Name = "lblAppTitle"
        Me.lblAppTitle.Size = New System.Drawing.Size(624, 46)
        Me.lblAppTitle.TabIndex = 0
        Me.lblAppTitle.Text = "Oracle DDL to Excel"
        Me.lblAppTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(12, 73)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(53, 20)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title: "
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(99, 70)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(537, 26)
        Me.txtTitle.TabIndex = 0
        '
        'lblSchema
        '
        Me.lblSchema.AutoSize = True
        Me.lblSchema.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSchema.Location = New System.Drawing.Point(14, 105)
        Me.lblSchema.Name = "lblSchema"
        Me.lblSchema.Size = New System.Drawing.Size(79, 20)
        Me.lblSchema.TabIndex = 0
        Me.lblSchema.Text = "Schema:"
        '
        'txtSchema
        '
        Me.txtSchema.Location = New System.Drawing.Point(99, 102)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.ReadOnly = True
        Me.txtSchema.Size = New System.Drawing.Size(437, 26)
        Me.txtSchema.TabIndex = 0
        Me.txtSchema.TabStop = False
        '
        'btnOpenSchema
        '
        Me.btnOpenSchema.AutoSize = True
        Me.btnOpenSchema.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnOpenSchema.Location = New System.Drawing.Point(542, 100)
        Me.btnOpenSchema.Name = "btnOpenSchema"
        Me.btnOpenSchema.Size = New System.Drawing.Size(58, 30)
        Me.btnOpenSchema.TabIndex = 1
        Me.btnOpenSchema.Text = "Open"
        Me.btnOpenSchema.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.AutoSize = True
        Me.btnExport.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnExport.Location = New System.Drawing.Point(249, 182)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(65, 30)
        Me.btnExport.TabIndex = 4
        Me.btnExport.Text = "Export"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.AutoSize = True
        Me.btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnCancel.Location = New System.Drawing.Point(337, 182)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(68, 30)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ofdOpenFile
        '
        Me.ofdOpenFile.Filter = "SQL|*.sql|TEXT|*.txt"
        Me.ofdOpenFile.Title = "Open File"
        '
        'pbProgress
        '
        Me.pbProgress.Location = New System.Drawing.Point(12, 223)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(624, 23)
        Me.pbProgress.TabIndex = 0
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(12, 249)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(35, 20)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Idle"
        '
        'timDelayIdleMessage
        '
        Me.timDelayIdleMessage.Interval = 3000
        '
        'btnAddSchema
        '
        Me.btnAddSchema.AutoSize = True
        Me.btnAddSchema.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnAddSchema.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnAddSchema.Location = New System.Drawing.Point(274, 134)
        Me.btnAddSchema.Name = "btnAddSchema"
        Me.btnAddSchema.Size = New System.Drawing.Size(111, 30)
        Me.btnAddSchema.TabIndex = 3
        Me.btnAddSchema.Text = "Add Schema"
        Me.btnAddSchema.UseVisualStyleBackColor = True
        '
        'btnRemoveSchema
        '
        Me.btnRemoveSchema.BackgroundImage = Global.OracleDDLtoExcel.My.Resources.Resources.trash
        Me.btnRemoveSchema.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnRemoveSchema.Enabled = False
        Me.btnRemoveSchema.Location = New System.Drawing.Point(606, 100)
        Me.btnRemoveSchema.Name = "btnRemoveSchema"
        Me.btnRemoveSchema.Size = New System.Drawing.Size(30, 30)
        Me.btnRemoveSchema.TabIndex = 2
        Me.btnRemoveSchema.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(648, 277)
        Me.Controls.Add(Me.btnRemoveSchema)
        Me.Controls.Add(Me.btnAddSchema)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.pbProgress)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.btnOpenSchema)
        Me.Controls.Add(Me.txtSchema)
        Me.Controls.Add(Me.lblSchema)
        Me.Controls.Add(Me.txtTitle)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.lblAppTitle)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Oracle DDL to Excel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblAppTitle As Label
    Friend WithEvents lblTitle As Label
    Friend WithEvents txtTitle As TextBox
    Friend WithEvents lblSchema As Label
    Friend WithEvents txtSchema As TextBox
    Friend WithEvents btnOpenSchema As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents ofdOpenFile As OpenFileDialog
    Friend WithEvents pbProgress As ProgressBar
    Friend WithEvents lblStatus As Label
    Friend WithEvents timDelayIdleMessage As Timer
    Friend WithEvents btnAddSchema As Button
    Friend WithEvents btnRemoveSchema As Button
End Class
