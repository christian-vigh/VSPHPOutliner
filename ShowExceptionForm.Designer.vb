<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ShowExceptionForm
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Message = New System.Windows.Forms.TextBox()
		Me.CloseButton = New System.Windows.Forms.Button()
		Me.SuspendLayout
		'
		'Label1
		'
		Me.Label1.AutoSize = true
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.Label1.Location = New System.Drawing.Point(9, 9)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(124, 13)
		Me.Label1.TabIndex = 0
		Me.Label1.Text = "Exception contents :"
		'
		'Message
		'
		Me.Message.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.Message.Location = New System.Drawing.Point(13, 31)
		Me.Message.Multiline = true
		Me.Message.Name = "Message"
		Me.Message.ReadOnly = true
		Me.Message.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.Message.Size = New System.Drawing.Size(970, 355)
		Me.Message.TabIndex = 1
		Me.Message.WordWrap = false
		'
		'CloseButton
		'
		Me.CloseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.CloseButton.Location = New System.Drawing.Point(908, 402)
		Me.CloseButton.Name = "CloseButton"
		Me.CloseButton.Size = New System.Drawing.Size(75, 23)
		Me.CloseButton.TabIndex = 0
		Me.CloseButton.Text = "&Close"
		Me.CloseButton.UseVisualStyleBackColor = true
		'
		'ShowExceptionForm
		'
		Me.AcceptButton = Me.CloseButton
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.CancelButton = Me.CloseButton
		Me.ClientSize = New System.Drawing.Size(995, 437)
		Me.ControlBox = false
		Me.Controls.Add(Me.CloseButton)
		Me.Controls.Add(Me.Message)
		Me.Controls.Add(Me.Label1)
		Me.MaximizeBox = false
		Me.MinimizeBox = false
		Me.Name = "ShowExceptionForm"
		Me.ShowIcon = false
		Me.ShowInTaskbar = false
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "ShowExceptionForm"
		Me.ResumeLayout(false)
		Me.PerformLayout

End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Message As System.Windows.Forms.TextBox
    Friend WithEvents CloseButton As System.Windows.Forms.Button
End Class
