Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Text
Imports System.IO.Ports


Namespace IODevices

    Public Class SerialDevice
        Inherits IODevice

        Protected defaultreadtimeout As Integer = 5          'port read timeout should be set to short value 

        Protected defaultwritetimeout As Integer = 1000


        Protected commport As SerialPort = Nothing
        Protected commerr As String

        Protected portname As String

        Protected fullportspec As String

        Private Shared CrLf As String = vbCr & vbLf 'System.Text.Encoding.UTF8.GetString(new byte[2] { 13, 10 }); //CrLf;
        Private Shared Cr As String = vbCr ' System.Text.Encoding.UTF8.GetString(new byte[1] { 13}); 
        Private Shared Lf As String = vbLf ' System.Text.Encoding.UTF8.GetString(new byte[1] { 10 }); 

        Private _enableDataReceivedEvent As Boolean = False
        Public Property EnableDataReceivedEvent() As Boolean


            Get
                Return _enableDataReceivedEvent
            End Get
            Set(ByVal value As Boolean)
                If Not _enableDataReceivedEvent AndAlso value Then
                    delayread = 100    'use events to interrupt waiting delays, then delays can be set long
                    delayrereadontimeout = 100
                    AddHandler commport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
                    _enableDataReceivedEvent = True
                End If
                If _enableDataReceivedEvent AndAlso Not value Then
                    delayread = 10
                    delayrereadontimeout = 20
                    RemoveHandler commport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
                    _enableDataReceivedEvent = False
                End If
            End Set
        End Property




        Public Sub New(ByVal name As String, ByVal addr As String)
            MyBase.New(name, addr)


            init(name, addr, "", 4096)
        End Sub

        'more complete constructor

        Public Sub New(ByVal name As String, ByVal addr As String, ByVal termstr As String, ByVal buffersize As Integer)
            MyBase.New(name, addr)
            'init base class storing name and addr


            init(name, addr, termstr, buffersize)
        End Sub

        Private Sub init(ByVal name As String, ByVal addr As String, ByVal termstr As String, ByVal buffersize As Integer)
            'termstr defines the "end of line" character or sequence of caracters (eg. Cr, Lf, CRLF etc),  will override the setting in "addr"

            fullportspec = addr


            interfacelockid = -1
            'no bus lock (serial ports perfectly thread-safe)

            statusmsg = "opening comm port for device " & name
            OpenComm(fullportspec, buffersize)

            If Not String.IsNullOrEmpty(termstr) Then
                commport.NewLine = termstr
            End If
            'override settings in addr if any

            devaddr = portname
            'store short addr for display  (baud etc removed)
            interfacename = "serial"

            EnableDataReceivedEvent = True

            statusmsg = ""
            AddToList()



        End Sub


        Private Sub DataReceivedHandler(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
            WakeUp()
        End Sub

        Protected Overrides Sub DisposeDevice()
            Try
                commport.Close()
            Catch
            End Try

        End Sub

        Protected Overrides Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'send cmd, return 0 if ok, 1 if timeout,  other if other error



            commport.DiscardInBuffer()


            Dim retval = 0


            Try

                commport.WriteLine(cmd)
            Catch generatedExceptionName As TimeoutException
                errcode = 1
                errmsg = "write timeout"

                retval = 1
            Catch ex As Exception
                retval = 2
                errmsg = ex.Message

                errcode = -1
            End Try

            Return retval


        End Function


        '--------------------------
        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'poll for status, return MAV bit 
            '  return 0 if ok, 1 if timeout,  other if other error


            mav = commport.BytesToRead > 0
            Return 0

        End Function

        ''--------------------
        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer

            '  Public Shared Sub Receive(ByVal board As Integer, ByVal address As Short, ByVal buffer As Byte(), 
            'ByVal count As Integer, ByVal termination As Integer)

            ' commport.DiscardInBuffer();   'bad if reread on timeout or not eoi! will destroy data



            Dim retval = 0

            Dim respstr As String = Nothing

            Try
                respstr = commport.ReadLine()
                arr = System.Text.Encoding.UTF8.GetBytes(respstr)


                'because of readline


                EOI = True
            Catch generatedExceptionName As TimeoutException
                errcode = 1
                errmsg = "read timeout"

                retval = 1
            Catch ex As Exception
                retval = 2
                errmsg = ex.Message

                errcode = -1
            End Try

            Return retval




        End Function




        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer

            '        Public Shared Sub DevClear(ByVal board As Integer, ByVal address As Short)


            Try
                commport.DiscardOutBuffer()
                commport.DiscardInBuffer()
                Return 0
            Catch
                errcode = -1
                errmsg = "cannot flush buffer"
                Return -1
            End Try


        End Function




        '*********** other private functions



        Private Function OpenComm(ByVal portspec As String, ByVal buffersize As Integer) As Object

            ' return 0 if ok, -1 if error


            commerr = ""

            If commport Is Nothing Then
                commport = New SerialPort()
                commport.ReadBufferSize = buffersize
            Else
                If commport.IsOpen Then
                    Try
                        commport.Close()
                    Catch ex As Exception
                        commerr = ex.Message
                        Return -1
                    End Try
                End If
            End If


            'extract name:


            Dim parr As String() = portspec.ToUpper().Split(":".ToCharArray())
            '"com1:9600,N,8,2"
            portname = parr(0)

            'settings
            commport.PortName = portname


            Dim sarr As String() = parr(1).Split(",".ToCharArray())

            commport.BaudRate = Integer.Parse(sarr(0))

            Select Case sarr(1)
                Case "E"
                    commport.Parity = Parity.Even
                    Exit Select
                Case "M"
                    commport.Parity = Parity.Mark
                    Exit Select
                Case "N"
                    commport.Parity = Parity.None
                    Exit Select
                Case "O"
                    commport.Parity = Parity.Odd
                    Exit Select
                Case "S"
                    commport.Parity = Parity.Space
                    Exit Select
                Case Else

                    Throw New Exception(" invalid comm parity specification")
            End Select

            commport.DataBits = Integer.Parse(sarr(2))

            Select Case Integer.Parse(sarr(3))
                Case 0
                    commport.StopBits = StopBits.None
                    Exit Select
                Case 1
                    commport.StopBits = StopBits.One
                    Exit Select
                Case 2
                    commport.StopBits = StopBits.Two
                    Exit Select
                Case Else
                    Throw New Exception(" invalid comm stop bits specification")
            End Select



            If sarr.Length > 4 AndAlso Not String.IsNullOrEmpty(sarr(4).Trim()) Then
                Select Case sarr(4).Trim()
                    Case "CR"
                        commport.NewLine = Cr
                        Exit Select
                    Case "LF"
                        commport.NewLine = Lf
                        Exit Select
                    Case "CRLF"
                        commport.NewLine = CrLf
                        Exit Select
                    Case Else
                        commport.NewLine = sarr(4).Trim()
                        Exit Select

                End Select
            End If

            'exemple "COM1:9600,N,8,2"  -> "com1",9600,Parity.None,8, stopbits.Two

            commport.WriteTimeout = defaultreadtimeout
            commport.ReadTimeout = defaultwritetimeout


			commport.Handshake = Handshake.None
            commport.Open()
            ' dont catch errors in constructor


			'added in 2020
			
			commport.DtrEnable = True
			commport.RtsEnable = True

            commport.DiscardOutBuffer()
            commport.DiscardInBuffer()


            Return 0
        End Function


    End Class

End Namespace
