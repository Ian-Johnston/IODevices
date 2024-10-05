Imports System.Net.Sockets
Imports System.Text


Namespace IODevices
    Public Class PrologixDeviceEthernet
        Inherits IODevice

        Private tcpClient As TcpClient
        Private networkStream As NetworkStream
        Private ipAddress As String
        Private port As Integer
        Private gpibaddr As String
        Private Shared Lf As String = vbLf
        Private Shared termstr As String = Lf

        ' Address format: n:IPaddress:port
        Public Sub New(name As String, addr As String)
            MyBase.New(name, addr)

            ' Parse the address to extract GPIB address, IP address, and port
            Dim parr As String() = devaddr.ToUpper().Split(":".ToCharArray())
            gpibaddr = parr(0)
            ipAddress = parr(1)
            port = Integer.Parse(parr(2))

            OpenConnection()

            ' Configure EOI
            SendRawCommand("++eoi 1")
            SendRawCommand("++eot_enable 1")
            SendRawCommand("++eot_char 10")

            interfacename = "Prologix Ethernet"
            statusmsg = ""

            AddToList()
        End Sub

        Private Sub OpenConnection()
            Try
                tcpClient = New TcpClient(ipAddress, port)
                networkStream = tcpClient.GetStream()
                networkStream.ReadTimeout = 1000
                networkStream.WriteTimeout = 1000
            Catch ex As Exception
                Throw New Exception("Unable to connect to the device at " & ipAddress & ":" & port)
            End Try
        End Sub

        Protected Overrides Sub DisposeDevice()
            Try
                SendRawCommand("++addr " & gpibaddr & termstr & "++loc")
                Threading.Thread.Sleep(500)
                networkStream.Close()
                tcpClient.Close()
            Catch
                ' Handle any cleanup exceptions
            End Try
        End Sub

        Protected Overrides Function Send(cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                SendRawCommand("++addr " & gpibaddr)
                SendRawCommand(cmd)
                Return 0
            Catch ex As Exception
                errcode = -1
                errmsg = ex.Message
                Return 2
            End Try
        End Function

        Private Sub SendRawCommand(cmd As String)
            Dim data As Byte() = Encoding.ASCII.GetBytes(cmd & Lf)
            networkStream.Write(data, 0, data.Length)
        End Sub

        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                SendRawCommand("++spoll " & gpibaddr)
                Dim resp As String = ReceiveRawResponse()
                statusbyte = Byte.Parse(resp)
                mav = (statusbyte And MAVmask) <> 0
                Return 0
            Catch ex As Exception
                errcode = -1
                errmsg = ex.Message
                Return 2
            End Try
        End Function

        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                SendRawCommand("++addr " & gpibaddr & termstr & "++read")
                Dim resp As String = ReceiveRawResponse()
                arr = Encoding.ASCII.GetBytes(resp)
                EOI = True
                Return 0
            Catch ex As Exception
                errcode = -1
                errmsg = ex.Message
                Return 2
            End Try
        End Function

        Private Function ReceiveRawResponse() As String
            Dim responseBuffer As Byte() = New Byte(4095) {}
            Dim bytesRead As Integer = networkStream.Read(responseBuffer, 0, responseBuffer.Length)
            Return Encoding.ASCII.GetString(responseBuffer, 0, bytesRead).TrimEnd()
        End Function

        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                SendRawCommand("++addr " & gpibaddr & termstr & "++clr ")
                Return 0
            Catch ex As Exception
                errcode = -1
                errmsg = "cannot flush buffer"
                Return -1
            End Try
        End Function

    End Class
End Namespace
