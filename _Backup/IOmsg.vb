Option Strict Off
Option Explicit On
Imports System.Drawing
Imports System.Windows.Forms
Imports IODevices
Namespace IODeviceForms  'internal namespace

    Friend Class IOmsgForm  'one form per device
        Inherits System.Windows.Forms.Form

        Private query As IOQuery  'set by calling thread
        Public Sub showit(ByVal q As IOQuery) 'called by gpib thread

            Dim cberr As Boolean = (q.status And 256) <> 0
            If cberr Then q.status -= 256

            query = q
            Text = q.device.devname + " error"
            msg1.ForeColor = Color.Red
            lbl_retry.ForeColor = Color.Red

            lbl_retry.Visible = True
            'cmd_abort.Visible = False

            If query.status > 0 Then
                If (query.status And 2) = 0 Then             'bit 2 on send(0)/recv(1)
                    msg1.Text = "error while sending data to " + query.device.devname
                Else
                    msg1.Text = "error while receiving data from " + query.device.devname
                End If
                'compose info:

                txt.Text = ""


                txt.Text &= "address: " & query.device.devaddr
                txt.Text &= vbCrLf & "command: " & query.cmd

                If (query.status And 4) > 0 Then
                    txt.Text &= vbCrLf & "interface returned error n°" & query.errcode.ToString
                End If

                txt.Text &= vbCrLf & query.errmsg
                'if retry:
                If query.task.retry Then
                    lbl_retry.Text = " retrying ..."
                    cmd_abort.Visible = True
                    Timer1.Enabled = True
                Else
                    lbl_retry.Text = "IO operation abandoned"
                    cmd_abort.Visible = False

                    Timer1.Enabled = False
                End If

            ElseIf cberr Then
                msg1.Text = query.device.devname + " : error in callback "
                txt.Text = "unhandled exception in user callback function"

                txt.Text &= vbCrLf & "command: " & query.cmd

                txt.Text &= vbCrLf & query.errmsg

                lbl_retry.Text = "IO operation completed"
                cmd_abort.Visible = False

                Timer1.Enabled = False

            End If





            Show()

        End Sub

        Private Sub cmd_abort_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_abort.Click

            query.AbortRetry()

            Close()

        End Sub



        Private Sub cmd_ok_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_ok.Click

            Close()

        End Sub


        Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
            'make it blinking

            If lbl_retry.Visible Then
                lbl_retry.Visible = False
            Else
                lbl_retry.Visible = True
            End If

            Me.Show()

        End Sub

        Private Sub cmd_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        End Sub

        Private Sub cmd_abortall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_abortall.Click

            query.AbortRetry()  'abort this (blocking or async)

            query.AbortAll() 'abort all async commands in queue
            Close()
            'Hide()
        End Sub

        Private Sub IOmsgForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            txt.Multiline = True
        End Sub


        Public Sub shutdownmsg(ByVal devname As String) 'called by iodevice when shutting down

            Try
                Text = devname
                msg1.ForeColor = Color.Red
                lbl_retry.Visible = False

                cmd_abort.Visible = False
                cmd_abortall.Visible = False
                cmd_ok.Visible = False

                msg1.Text = "shutting down " + devname + " ..."


                Show()
                WindowState = FormWindowState.Normal
                BringToFront()

            Catch ex As Exception

            End Try



        End Sub

        Private Sub txt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt.TextChanged

        End Sub
    End Class
End Namespace