<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportClass
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LbxName = New System.Windows.Forms.ListBox()
        Me.LbxQualification = New System.Windows.Forms.ListBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.tbxPassFailRatio = New System.Windows.Forms.TextBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(20, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(218, 45)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Class Report:"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(196, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(251, 45)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Pass/Fail Ratio:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(462, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 45)
        Me.Label4.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(244, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(0, 45)
        Me.Label5.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(57, 202)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(169, 45)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Students:"
        '
        'LbxName
        '
        Me.LbxName.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LbxName.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.LbxName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.LbxName.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbxName.FormattingEnabled = True
        Me.LbxName.ItemHeight = 45
        Me.LbxName.Location = New System.Drawing.Point(177, 278)
        Me.LbxName.Name = "LbxName"
        Me.LbxName.Size = New System.Drawing.Size(270, 360)
        Me.LbxName.TabIndex = 6
        '
        'LbxQualification
        '
        Me.LbxQualification.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LbxQualification.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.LbxQualification.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.LbxQualification.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbxQualification.FormattingEnabled = True
        Me.LbxQualification.ItemHeight = 45
        Me.LbxQualification.Location = New System.Drawing.Point(613, 278)
        Me.LbxQualification.Name = "LbxQualification"
        Me.LbxQualification.Size = New System.Drawing.Size(270, 360)
        Me.LbxQualification.TabIndex = 7
        '
        'PictureBox3
        '
        Me.PictureBox3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.PictureBox3.Image = Global.NEA_Prototype.My.Resources.Resources.HomeButton
        Me.PictureBox3.Location = New System.Drawing.Point(1183, 611)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(69, 59)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 15
        Me.PictureBox3.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TextBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(252, 55)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(261, 45)
        Me.TextBox1.TabIndex = 16
        '
        'tbxPassFailRatio
        '
        Me.tbxPassFailRatio.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.tbxPassFailRatio.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.tbxPassFailRatio.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbxPassFailRatio.Enabled = False
        Me.tbxPassFailRatio.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxPassFailRatio.Location = New System.Drawing.Point(453, 124)
        Me.tbxPassFailRatio.Name = "tbxPassFailRatio"
        Me.tbxPassFailRatio.Size = New System.Drawing.Size(261, 45)
        Me.tbxPassFailRatio.TabIndex = 17
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Font = New System.Drawing.Font("Comic Sans MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel1.Location = New System.Drawing.Point(1177, 9)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(75, 23)
        Me.LinkLabel1.TabIndex = 18
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Sign Out"
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Comic Sans MS", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(605, 202)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(224, 45)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Qualification:"
        '
        'ReportClass
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ClientSize = New System.Drawing.Size(1264, 682)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.tbxPassFailRatio)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.LbxQualification)
        Me.Controls.Add(Me.LbxName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ReportClass"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ReportClass"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LbxName As ListBox
    Friend WithEvents LbxQualification As ListBox
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents tbxPassFailRatio As TextBox
    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents Label7 As Label
End Class
