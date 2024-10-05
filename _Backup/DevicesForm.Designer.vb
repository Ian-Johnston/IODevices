Namespace IODeviceForms  'internal namespace

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class DevicesForm
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
            Me.components = New System.ComponentModel.Container
            Me.Label1 = New System.Windows.Forms.Label
            Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
            Me.chk_gpibcmd = New System.Windows.Forms.CheckBox
            Me.lblstatus = New System.Windows.Forms.Label
            Me.txt_list = New System.Windows.Forms.TextBox
            Me.SuspendLayout()
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(21, 9)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(81, 13)
            Me.Label1.TabIndex = 0
            Me.Label1.Text = " Open Devices:"
            '
            'Timer1
            '
            Me.Timer1.Interval = 200
            '
            'chk_gpibcmd
            '
            Me.chk_gpibcmd.AutoSize = True
            Me.chk_gpibcmd.Checked = True
            Me.chk_gpibcmd.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chk_gpibcmd.Location = New System.Drawing.Point(227, 8)
            Me.chk_gpibcmd.Name = "chk_gpibcmd"
            Me.chk_gpibcmd.Size = New System.Drawing.Size(144, 17)
            Me.chk_gpibcmd.TabIndex = 4
            Me.chk_gpibcmd.Text = "show queued commands"
            Me.chk_gpibcmd.UseVisualStyleBackColor = True
            '
            'lblstatus
            '
            Me.lblstatus.Location = New System.Drawing.Point(12, 246)
            Me.lblstatus.Name = "lblstatus"
            Me.lblstatus.Size = New System.Drawing.Size(445, 29)
            Me.lblstatus.TabIndex = 5
            Me.lblstatus.Text = "          "
            '
            'txt_list
            '
            Me.txt_list.Location = New System.Drawing.Point(12, 31)
            Me.txt_list.Multiline = True
            Me.txt_list.Name = "txt_list"
            Me.txt_list.Size = New System.Drawing.Size(445, 212)
            Me.txt_list.TabIndex = 6
            '
            'DevicesForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(471, 284)
            Me.Controls.Add(Me.lblstatus)
            Me.Controls.Add(Me.txt_list)
            Me.Controls.Add(Me.chk_gpibcmd)
            Me.Controls.Add(Me.Label1)
            Me.Name = "DevicesForm"
            Me.Text = "IO Devices"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Timer1 As System.Windows.Forms.Timer
        Friend WithEvents chk_gpibcmd As System.Windows.Forms.CheckBox
        Friend WithEvents lblstatus As System.Windows.Forms.Label
        Friend WithEvents txt_list As System.Windows.Forms.TextBox
    End Class
End Namespace