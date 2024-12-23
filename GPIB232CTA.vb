
Imports System
Imports System.Threading
Imports System.Text
Imports System.IO.Ports



Namespace IODevices
	'NI based on SerialDevice code
	'address format:  n:COMc:baudrate
	'where n is the gpib address and c the COM port number
	' e.g.  8:COM7:115000  for device at gpib addr.8, NI board on COM7, configured for 115000 bauds

	Public Class GPIB_NI_232CT_A
		Inherits IODevice
		Protected Structure ComInfo
			Public COMportNum As Integer
			Public COMport As SerialPort
			Public NumDevices As Integer
		End Structure
		Protected MyPort As ComInfo
		Protected Shared openPorts As New List(Of ComInfo)
		Protected defaultreadtimeout As Integer = 1000 'port read timeout, can be set short
		Protected defaultwritetimeout As Integer = 1000
		'		Protected ComInfo As ComInfo
		'		Protected commport As SerialPort = Nothing
		Protected commerr As String
		'		Protected commportnum As Integer 'e.g. will be 3 for "COM3" etc.
		Protected baudrate As String 'will be like "9600"
		Protected gpibaddr As String
		Private Shared Lf As String = vbLf ' System.Text.Encoding.ASCII.GetString(new byte[1] { 10 }); 
		Private Shared termstr As String = Lf 'comm port will be configured with (++eot_char also)
		Private readinitiated As Boolean = False
		Private _enableDataReceivedEvent As Boolean = False

		Public Property EnableDataReceivedEvent() As Boolean 'use events to quit waiting delays, then delays can be set long
			Get
				Return _enableDataReceivedEvent
			End Get
			Set
				If Not _enableDataReceivedEvent AndAlso Value Then
					delayread = 100
					delayrereadontimeout = 100
					AddHandler MyPort.COMport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
					_enableDataReceivedEvent = True
				End If
				If _enableDataReceivedEvent AndAlso Not Value Then
					delayread = 10
					delayrereadontimeout = 20
					RemoveHandler MyPort.COMport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
					_enableDataReceivedEvent = False
				End If
			End Set
		End Property

		'address format:  n:COMc:baudrate
		'where n is the gpib address and c the COM port number
		' e.g.  8:COM7:115200  for device at gpib addr.8, prologix board on COM7, configured for 115200 bauds

		Public Sub New(name As String, addr As String)
			MyBase.New(name, addr)

			statusmsg = "opening comm port for device " & name

			OpenComm()

			'configure EOI
			'MyPort.COMport.WriteLine("++eoi 1") 'append EOI
			'MyPort.COMport.WriteLine("++eot_enable 1")
			'MyPort.COMport.WriteLine("++eot_char 10") 'translate EOI to LF
			interfacename = "GPIB-232CT-A Serial"

			EnableDataReceivedEvent = True

			statusmsg = ""

			'interface locking needed here (different threads talk to the same com port) 
			interfacelockid = 25 + MyPort.COMportNum

			AddToList()
		End Sub


		Private Sub DataReceivedHandler(sender As Object, e As SerialDataReceivedEventArgs)
			WakeUp()
		End Sub

		Protected Overrides Sub DisposeDevice()
			Try
				Dim removed As Boolean
				'MyPort.COMport.WriteLine("wrt " & gpibaddr & termstr & "LOCAL 7" & gpibaddr)
				MyPort.COMport.WriteLine("clr")
				Thread.Sleep(500)
				MyPort.NumDevices = MyPort.NumDevices - 1
				If (MyPort.NumDevices = 0) Then
					MyPort.COMport.Close()
					For Each rec In openPorts
						If rec.COMportNum = MyPort.COMportNum Then
							removed = openPorts.Remove(rec)
							MyPort = Nothing
						End If
					Next
				End If
			Catch ex As Exception
			End Try

		End Sub

		Protected Overrides Function Send(cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
			'send cmd, return 0 if ok, 1 if timeout,  other if other error

			MyPort.COMport.DiscardInBuffer()

			Dim retval = 0

			Try
				MyPort.COMport.WriteLine("wrt " & gpibaddr & termstr) 'set address
				MyPort.COMport.WriteLine(cmd)
			Catch generatedExceptionName As TimeoutException
				errcode = 1
				errmsg = "write timeout"

				retval = 1
			Catch ex As Exception
				retval = 2
				errmsg = ex.Message

				errcode = -1
			End Try

			readinitiated = False  'rearm for new reading
			Return retval
		End Function

		'--------------------------
		Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
			'poll for status, return MAV bit 
			'  return 0 if ok, 1 if timeout,  other if other error
			MyPort.COMport.DiscardInBuffer()  'only input after ++read command is considered

			Dim retval = 0

			Dim cmd As String = "rsp " & gpibaddr
			Dim resp As String
			mav = False

			Try
				MyPort.COMport.WriteLine(cmd)
				'MyPort.COMport.WriteLine("rd #100 " & gpibaddr)
				resp = MyPort.COMport.ReadLine()
				statusbyte = Byte.Parse(resp)

				mav = (statusbyte And MAVmaskGBIP232CTA) <> 0
			Catch generatedExceptionName As TimeoutException
				errcode = 1
				errmsg = "spoll: timeout"              'will retry on timeout
				retval = 1
			Catch ex As Exception
				retval = 2
				errmsg = ex.Message

				errcode = -1
			End Try

			Return retval

		End Function

		''--------------------
		Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer
			'this function will be repeatedly called if there is timeout on "Readline"
			'however for Prologix the "++read" command should not be repeated therefore 
			'we store the status in "readinitiated" flag. 
			' However we have problem if timeout is caused by gpib and not ReadLine,
			'then ++read would have to be repeated but with Prologix we cannot make the difference
			'so that polling should always be used with this interface to avoid such situation
			Dim retval = 0
			Dim respstr As String = Nothing

			If Not readinitiated Then
				Try
					Dim cmd As String
					'					cmd = "++addr " & gpibaddr & termstr & "++read eoi" 'read until eoi or timeout
					cmd = "rd #100 " & gpibaddr 'read until eoi or timeout
					MyPort.COMport.WriteLine(cmd)
					readinitiated = True
				Catch generatedExceptionName As TimeoutException
					errcode = 1
					errmsg = "write timeout"   'will retry
					Return 1
				Catch ex As Exception
					errmsg = ex.Message
					errcode = -1
					Return -1
				End Try
			End If

			Try

				respstr = MyPort.COMport.ReadLine() 'can timeout but should not cause error:
				'once read is initiated only reading will be retried on timeout 
				' (then we can set short timeout on the COM port)

				arr = System.Text.Encoding.ASCII.GetBytes(respstr)

				EOI = True   'EOI translated to end of line

				readinitiated = False  'reset for next reading

			Catch generatedExceptionName As TimeoutException
				errcode = 1
				errmsg = "read timeout" 'will not show up until cumulative timeout delay "readtimeout"

				retval = 1  'will retry
			Catch ex As Exception
				retval = 2
				errmsg = ex.Message

				errcode = -1
			End Try

			Return retval

		End Function


		Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer

			'        Public Shared Sub DevClear(ByVal board As Integer, ByVal address As Short)

			Dim cmd As String = "clr "

			Try
				MyPort.COMport.WriteLine(cmd)
				MyPort.COMport.DiscardOutBuffer()
				MyPort.COMport.DiscardInBuffer()
				Return 0
			Catch
				errcode = -1
				errmsg = "cannot flush buffer"
				Return -1
			End Try

		End Function

		'*********** other private functions

		'try to open port then parse devaddr to get commport, baudrate and gpibaddr,
		'and configure the port     



		Private Function OpenComm() As Integer
			' return 0 if ok, -1 if error
			Dim COMportNum As Integer
			commerr = ""
			'----- extract fields from devaddr: -----
			Dim parr As String() = devaddr.ToUpper().Split(":".ToCharArray())
			' e.g. "8:com1:9600"
			gpibaddr = parr(0)
			Dim COMportName As String = parr(1)
			baudrate = parr(2)
			Try
				COMportNum = Integer.Parse(COMportName.Remove(0, 3))
			Catch
				Throw New Exception("invalid COM port specification: " & COMportName)
			End Try
			'----- search openPorts list for MyPort.COMportNum -----
			MyPort = ComInfoByPort(COMportNum)
			If MyPort.COMport Is Nothing Then
				'----- MyPort not found in openPorts -----
				MyPort.COMport = New SerialPort()
				MyPort.COMportNum = COMportNum
				MyPort.COMport.ReadBufferSize = 4096
				'----- Apply settings -----
				MyPort.COMport.BaudRate = Integer.Parse(baudrate)
				MyPort.COMport.PortName = COMportName
				MyPort.COMport.Parity = Parity.None
				MyPort.COMport.StopBits = StopBits.One
				MyPort.COMport.DataBits = 8
				MyPort.COMport.NewLine = termstr 'LF, if modified then ++eot  should be updated too
				MyPort.COMport.WriteTimeout = defaultwritetimeout
				MyPort.COMport.ReadTimeout = defaultreadtimeout
				' RTS/CTS handshaking as suggested for prologix, to check
				MyPort.COMport.Handshake = Handshake.RequestToSend
				MyPort.COMport.DtrEnable = True

				MyPort.COMport.Open()
				' dont catch errors in constructor
				MyPort.COMport.DiscardOutBuffer()
				MyPort.COMport.DiscardInBuffer()
				MyPort.NumDevices = 1
				openPorts.Add(MyPort)
				Return 0
			Else
				If (MyPort.NumDevices > 0) Then
					'----- Add 1 device to MyPort -----
					MyPort.NumDevices = MyPort.NumDevices + 1
					For Each rec In openPorts
						If rec.COMportNum = MyPort.COMportNum Then
							rec.NumDevices = MyPort.NumDevices
						End If
					Next
					Return 0
				End If
			End If
		End Function

		'----- find ComInfo in openPorts list using COMportNum -----
		Private Shared Function ComInfoByPort(ByVal Num As Integer) As ComInfo
			If openPorts Is Nothing Then
				Return Nothing
			End If
			For Each Port As ComInfo In openPorts
				If Port.COMportNum = Num Then
					Return Port
				End If
			Next
			Return Nothing
		End Function
		'----------------------------------------------------------------------
	End Class

End Namespace
