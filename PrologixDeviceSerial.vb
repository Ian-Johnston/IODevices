
Imports System
Imports System.Threading
Imports System.Text
Imports System.IO.Ports


Namespace IODevices
	'PrologixDevice based on SerialDevice code
	'address format:  n:COMc:baudrate
	'where n is the gpib address and c the COM port number
	' e.g.  8:COM7:115000  for device at gpib addr.8, prologix board on COM7, configured for 115000 bauds
	Public Structure ComInfo
		public commportnum as integer
		public shared  commport as SerialPort
		public shared  numdevices As integer
	End Structure

	Public Class PrologixDeviceSerial
		Inherits IODevice

		Protected defaultreadtimeout As Integer = 100 'port read timeout, can be set short
		Protected defaultwritetimeout As Integer = 1000
		Protected ComInfo As ComInfo
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
					AddHandler ComInfo.commport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
					_enableDataReceivedEvent = True
				End If
				If _enableDataReceivedEvent AndAlso Not Value Then
					delayread = 10
					delayrereadontimeout = 20
					RemoveHandler ComInfo.commport.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceivedHandler)
					_enableDataReceivedEvent = False
				End If
			End Set
		End Property

		'address format:  n:COMc:baudrate
		'where n is the gpib address and c the COM port number
		' e.g.  8:COM7:115000  for device at gpib addr.8, prologix board on COM7, configured for 115000 bauds

		Public Sub New(name As String, addr As String)
			MyBase.New(name, addr)

			statusmsg = "opening comm port for device " & name

			OpenComm()

			'configure EOI
			ComInfo.commport.WriteLine("++eoi 1") 'append EOI
			ComInfo.commport.WriteLine("++eot_enable 1")
			ComInfo.commport.WriteLine("++eot_char 10") 'translate EOI to LF
			interfacename = "Prologix Serial"

			EnableDataReceivedEvent = True

			statusmsg = ""

			'interface locking needed here (different threads talk to the same com port) 
			interfacelockid = 25 + ComInfo.commportnum

			AddToList()
		End Sub


		Private Sub DataReceivedHandler(sender As Object, e As SerialDataReceivedEventArgs)
			WakeUp()
		End Sub

		Protected Overrides Sub DisposeDevice()
			Try
				ComInfo.commport.WriteLine("++addr " & gpibaddr & termstr & "++loc")
				Thread.Sleep(500)
				ComInfo.numdevices = ComInfo.numdevices - 1
				If (ComInfo.numdevices = 0) Then
					ComInfo.commport.Close()
				End If

			Catch
			End Try

		End Sub

		Protected Overrides Function Send(cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
			'send cmd, return 0 if ok, 1 if timeout,  other if other error

			ComInfo.commport.DiscardInBuffer()

			Dim retval = 0

			Try
				ComInfo.commport.WriteLine("++addr " & gpibaddr) 'set address
				ComInfo.commport.WriteLine(cmd)
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
			ComInfo.commport.DiscardInBuffer()  'only input after ++read command is considered

			Dim retval = 0

			Dim cmd As String = "++spoll " & gpibaddr   'prologix command
			Dim resp As String
			mav = False

			Try
				ComInfo.commport.WriteLine(cmd)
				resp = ComInfo.commport.ReadLine()
				statusbyte = Byte.Parse(resp)

				mav = (statusbyte And MAVmask) <> 0
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
					cmd = "++addr " & gpibaddr & termstr & "++read" 'read until eoi or timeout
					ComInfo.commport.WriteLine(cmd)
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

				respstr = ComInfo.commport.ReadLine() 'can timeout but should not cause error:
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

			Dim cmd As String = "++addr " & gpibaddr & termstr & "++clr "

			Try
				ComInfo.commport.WriteLine(cmd)
				ComInfo.commport.DiscardOutBuffer()
				ComInfo.commport.DiscardInBuffer()
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

			commerr = ""
			'extract fields from devaddr:
			Dim parr As String() = devaddr.ToUpper().Split(":".ToCharArray())
			' e.g. "8:com1:9600"
			gpibaddr = parr(0)
			Dim commportname As String = parr(1)
			baudrate = parr(2)

			Try
				ComInfo.commportnum = Integer.Parse(commportname.Remove(0, 3))
			Catch
				Throw New Exception("invalid COM port specification: " & commportname)
			End Try
			If (ComInfo.numdevices > 0) Then
				ComInfo.numdevices = ComInfo.numdevices + 1
				Return 0
			End If

			If ComInfo.commport Is Nothing Then
				ComInfo.commport = New SerialPort()
				ComInfo.commport.ReadBufferSize = 4096
			Else
				If ComInfo.commport.IsOpen Then
					Try
						ComInfo.commport.Close()
					Catch ex As Exception
						commerr = ex.Message
						Return -1
					End Try
				End If
			End If


			'settings
			ComInfo.commport.BaudRate = Integer.Parse(baudrate)
			ComInfo.commport.PortName = commportname

			ComInfo.commport.Parity = Parity.None
			ComInfo.commport.StopBits = StopBits.One
			ComInfo.commport.DataBits = 8
			ComInfo.commport.NewLine = termstr 'LF, if modified then ++eot  should be updated too
			ComInfo.commport.WriteTimeout = defaultreadtimeout
			ComInfo.commport.ReadTimeout = defaultwritetimeout

			' RTS/CTS handshaking as suggested for prologix, to check
			ComInfo.commport.Handshake = Handshake.RequestToSend
			ComInfo.commport.DtrEnable = True

			ComInfo.commport.Open()
			' dont catch errors in constructor

			ComInfo.commport.DiscardOutBuffer()
			ComInfo.commport.DiscardInBuffer()

			Return 0
		End Function

	End Class

End Namespace
