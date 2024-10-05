Namespace IODeviceForms

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class IOmsgForm
#Region "Code généré par le Concepteur Windows Form "
        <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
            MyBase.New()
            'Cet appel est requis par le Concepteur Windows Form.
            InitializeComponent()
        End Sub
        'Form remplace la méthode Dispose pour nettoyer la liste des composants.
        <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
            If Disposing Then
                If Not components Is Nothing Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(Disposing)
        End Sub
        'Requise par le Concepteur Windows Form
        Private components As System.ComponentModel.IContainer
        Public ToolTip1 As System.Windows.Forms.ToolTip
        Public WithEvents Timer1 As System.Windows.Forms.Timer
        Public WithEvents cmd_abort As System.Windows.Forms.Button
        Public WithEvents cmd_ok As System.Windows.Forms.Button
        Public WithEvents lbl_retry As System.Windows.Forms.Label
        Public WithEvents msg1 As System.Windows.Forms.Label
        'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
        'Elle peut être modifiée à l'aide du Concepteur Windows Form.
        'Ne la modifiez pas à l'aide de l'éditeur de code.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container
            Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
            Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
            Me.cmd_abort = New System.Windows.Forms.Button
            Me.cmd_ok = New System.Windows.Forms.Button
            Me.lbl_retry = New System.Windows.Forms.Label
            Me.msg1 = New System.Windows.Forms.Label
            Me.txt = New System.Windows.Forms.TextBox
            Me.cmd_abortall = New System.Windows.Forms.Button
            Me.SuspendLayout()
            '
            'Timer1
            '
            Me.Timer1.Interval = 500
            '
            'cmd_abort
            '
            Me.cmd_abort.BackColor = System.Drawing.SystemColors.Control
            Me.cmd_abort.Cursor = System.Windows.Forms.Cursors.Default
            Me.cmd_abort.ForeColor = System.Drawing.SystemColors.ControlText
            Me.cmd_abort.Location = New System.Drawing.Point(6, 150)
            Me.cmd_abort.Name = "cmd_abort"
            Me.cmd_abort.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.cmd_abort.Size = New System.Drawing.Size(73, 36)
            Me.cmd_abort.TabIndex = 3
            Me.cmd_abort.Text = "abort retry"
            Me.cmd_abort.UseVisualStyleBackColor = False
            '
            'cmd_ok
            '
            Me.cmd_ok.BackColor = System.Drawing.SystemColors.Control
            Me.cmd_ok.Cursor = System.Windows.Forms.Cursors.Default
            Me.cmd_ok.ForeColor = System.Drawing.SystemColors.ControlText
            Me.cmd_ok.Location = New System.Drawing.Point(270, 153)
            Me.cmd_ok.Name = "cmd_ok"
            Me.cmd_ok.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.cmd_ok.Size = New System.Drawing.Size(73, 33)
            Me.cmd_ok.TabIndex = 2
            Me.cmd_ok.Text = "Close"
            Me.cmd_ok.UseVisualStyleBackColor = False
            '
            'lbl_retry
            '
            Me.lbl_retry.BackColor = System.Drawing.SystemColors.Window
            Me.lbl_retry.Cursor = System.Windows.Forms.Cursors.Default
            Me.lbl_retry.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lbl_retry.ForeColor = System.Drawing.SystemColors.WindowText
            Me.lbl_retry.Location = New System.Drawing.Point(2, 118)
            Me.lbl_retry.Name = "lbl_retry"
            Me.lbl_retry.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.lbl_retry.Size = New System.Drawing.Size(311, 29)
            Me.lbl_retry.TabIndex = 7
            Me.lbl_retry.Text = "retrying ..."
            Me.lbl_retry.Visible = False
            '
            'msg1
            '
            Me.msg1.BackColor = System.Drawing.SystemColors.Window
            Me.msg1.Cursor = System.Windows.Forms.Cursors.Default
            Me.msg1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.msg1.ForeColor = System.Drawing.SystemColors.WindowText
            Me.msg1.Location = New System.Drawing.Point(3, 9)
            Me.msg1.Name = "msg1"
            Me.msg1.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.msg1.Size = New System.Drawing.Size(331, 18)
            Me.msg1.TabIndex = 0
            Me.msg1.Text = " error while sending data to "
            '
            'txt
            '
            Me.txt.Location = New System.Drawing.Point(6, 30)
            Me.txt.Multiline = True
            Me.txt.Name = "txt"
            Me.txt.Size = New System.Drawing.Size(328, 85)
            Me.txt.TabIndex = 8
            '
            'cmd_abortall
            '
            Me.cmd_abortall.BackColor = System.Drawing.SystemColors.Control
            Me.cmd_abortall.Location = New System.Drawing.Point(100, 147)
            Me.cmd_abortall.Name = "cmd_abortall"
            Me.cmd_abortall.Size = New System.Drawing.Size(139, 39)
            Me.cmd_abortall.TabIndex = 9
            Me.cmd_abortall.Text = "abort all pending commands on this device"
            Me.cmd_abortall.UseVisualStyleBackColor = False
            '
            'IOmsgForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Window
            Me.ClientSize = New System.Drawing.Size(346, 194)
            Me.Controls.Add(Me.cmd_abortall)
            Me.Controls.Add(Me.txt)
            Me.Controls.Add(Me.cmd_abort)
            Me.Controls.Add(Me.cmd_ok)
            Me.Controls.Add(Me.lbl_retry)
            Me.Controls.Add(Me.msg1)
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.ForeColor = System.Drawing.SystemColors.WindowText
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.Location = New System.Drawing.Point(121, 116)
            Me.MaximizeBox = False
            Me.Name = "IOmsgForm"
            Me.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.Text = "IO Device error message"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents txt As System.Windows.Forms.TextBox
        Friend WithEvents cmd_abortall As System.Windows.Forms.Button
#End Region
    End Class

End Namespace