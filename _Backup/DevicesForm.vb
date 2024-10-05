Imports IODevices
Namespace IODeviceForms  'internal namespace (used within assembly to access devices and message forms)

    Friend Class DevicesForm


        Private enableupdate As Boolean = True



        Private Sub updatelist()

            'update  list
            Dim sl() As String


            Dim details As Boolean = chk_gpibcmd.Checked


            sl = IODevice.GetDeviceList(details)

            lblstatus.Text = IODevice.statusmsg

            If Not sl Is Nothing Then

                txt_list.Lines = sl

            End If



        End Sub

        Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

            If enableupdate And Not IODevice.updated Then
                updatelist()

                IODevice.updated = True
            End If

        End Sub

        Private Sub lb_gpib_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        End Sub

        Private Sub DevicesForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Timer1.Interval = 50
            Timer1.Enabled = True
        End Sub

        Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblstatus.Click

        End Sub

        Private Sub DevicesForm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
            txt_list.Height = Height - 100
            txt_list.Width = Width - 40
            lblstatus.Top = Height - 60
        End Sub

        Private Sub txt_list_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_list.MouseDown
            enableupdate = False
        End Sub

        Private Sub txt_list_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_list.MouseUp
            enableupdate = True
        End Sub
    End Class

End Namespace