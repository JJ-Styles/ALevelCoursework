<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PopUP
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
        Me.tbxMessage = New System.Windows.Forms.TextBox()
        Me.btnNo = New System.Windows.Forms.Button()
        Me.btnYes = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbxMessage
        '
        Me.tbxMessage.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.tbxMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbxMessage.Enabled = False
        Me.tbxMessage.Font = New System.Drawing.Font("Comic Sans MS", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxMessage.Location = New System.Drawing.Point(1, 9)
        Me.tbxMessage.Multiline = True
        Me.tbxMessage.Name = "tbxMessage"
        Me.tbxMessage.Size = New System.Drawing.Size(532, 99)
        Me.tbxMessage.TabIndex = 0
        Me.tbxMessage.Text = "This Will restore and archive all database data except from Admin Data." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Do you w" &
    "ish to continue?"
        Me.tbxMessage.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnNo
        '
        Me.btnNo.Font = New System.Drawing.Font("Comic Sans MS", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNo.Location = New System.Drawing.Point(31, 117)
        Me.btnNo.Name = "btnNo"
        Me.btnNo.Size = New System.Drawing.Size(185, 36)
        Me.btnNo.TabIndex = 1
        Me.btnNo.Text = "No"
        Me.btnNo.UseVisualStyleBackColor = True
        '
        'btnYes
        '
        Me.btnYes.Font = New System.Drawing.Font("Comic Sans MS", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnYes.Location = New System.Drawing.Point(320, 117)
        Me.btnYes.Name = "btnYes"
        Me.btnYes.Size = New System.Drawing.Size(185, 36)
        Me.btnYes.TabIndex = 2
        Me.btnYes.Text = "Yes"
        Me.btnYes.UseVisualStyleBackColor = True
        '
        'PopUP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ClientSize = New System.Drawing.Size(540, 175)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnYes)
        Me.Controls.Add(Me.btnNo)
        Me.Controls.Add(Me.tbxMessage)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PopUP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PopUP"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbxMessage As TextBox
    Friend WithEvents btnNo As Button
    Friend WithEvents btnYes As Button
End Class
