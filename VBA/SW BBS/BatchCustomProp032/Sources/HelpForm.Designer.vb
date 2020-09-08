<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelpForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LabelThread = New System.Windows.Forms.Label()
        Me.LabelAuthor = New System.Windows.Forms.Label()
        Me.LabelFifi = New System.Windows.Forms.Label()
        Me.LabelQuestion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LabelThread
        '
        Me.LabelThread.AutoSize = True
        Me.LabelThread.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LabelThread.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelThread.ForeColor = System.Drawing.SystemColors.Highlight
        Me.LabelThread.Location = New System.Drawing.Point(18, 77)
        Me.LabelThread.Name = "LabelThread"
        Me.LabelThread.Size = New System.Drawing.Size(220, 13)
        Me.LabelThread.TabIndex = 0
        Me.LabelThread.Text = "https://forum.solidworks.com/thread/229516"
        '
        'LabelAuthor
        '
        Me.LabelAuthor.AutoSize = True
        Me.LabelAuthor.Location = New System.Drawing.Point(18, 22)
        Me.LabelAuthor.Name = "LabelAuthor"
        Me.LabelAuthor.Size = New System.Drawing.Size(44, 13)
        Me.LabelAuthor.TabIndex = 1
        Me.LabelAuthor.Text = "Author: "
        '
        'LabelFifi
        '
        Me.LabelFifi.AutoSize = True
        Me.LabelFifi.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LabelFifi.Location = New System.Drawing.Point(61, 22)
        Me.LabelFifi.Name = "LabelFifi"
        Me.LabelFifi.Size = New System.Drawing.Size(20, 13)
        Me.LabelFifi.TabIndex = 2
        Me.LabelFifi.Text = "Fifi"
        '
        'LabelQuestion
        '
        Me.LabelQuestion.AutoSize = True
        Me.LabelQuestion.Location = New System.Drawing.Point(18, 58)
        Me.LabelQuestion.Name = "LabelQuestion"
        Me.LabelQuestion.Size = New System.Drawing.Size(150, 13)
        Me.LabelQuestion.TabIndex = 3
        Me.LabelQuestion.Text = "Have a question? Post it here:"
        '
        'HelpForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(263, 116)
        Me.Controls.Add(Me.LabelQuestion)
        Me.Controls.Add(Me.LabelFifi)
        Me.Controls.Add(Me.LabelAuthor)
        Me.Controls.Add(Me.LabelThread)
        Me.Name = "HelpForm"
        Me.Text = "Help"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelThread As Label
    Friend WithEvents LabelAuthor As Label
    Friend WithEvents LabelFifi As Label
    Friend WithEvents LabelQuestion As Label
End Class
