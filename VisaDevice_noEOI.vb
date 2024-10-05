
Imports System


Namespace IODevices
	
Public Class VisaDevice_noEOI
	 
        Inherits VisaDevice
	
	   'constructor: added termchar
	   Public Sub New(ByVal name As String, ByVal addr As String, ByVal interlocked As Boolean, ByVal termchar As Byte)
	   	MyBase.New(name, addr, interlocked)
	   	
	   	Me.termchar = termchar
		termstr = System.Text.Encoding.UTF8.GetString(New Byte(0) {termchar})
	   	'set visa attributes
	   	SetAttribute(VI_ATTR_TERMCHAR_EN, 1)  'stop read on termchar
	   	SetAttribute(VI_ATTR_TERMCHAR, termchar)

        End sub
	   	
	   	
		Protected Overrides Function Send(cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
			Dim l As Integer = cmd.Length
			Dim cmdc As Byte() = System.Text.Encoding.UTF8.GetBytes(cmd)
			If l > 0 Then 
				If cmdc(l - 1) <> termchar Then  'append termchar if not already there
                    cmd = cmd & termstr
                End If
			End If

			Return MyBase.Send(cmd, errcode, errmsg)
		End Function
		
		Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer

			Dim retval As Integer = 0

			Dim err As Boolean = False
			Dim result As UInteger = 0

			Dim cnt As Integer = 0
			Dim tmo As Boolean = False

			retval = 0
			result = VisaDll.viRead(devid, buffer, buffer.Length, cnt)

			err = Not (result = 0 Or result = VI_SUCCESS_MAX_CNT Or result = VI_SUCCESS_TERM_CHAR)

			EOI = (result = 0 Or result = VI_SUCCESS_TERM_CHAR)

			tmo = (result = VI_ERROR_TMO)

			If err Then
				errcode = Convert.ToInt32(result And VI_errmask)

				If tmo Then
					retval = 1
					errmsg = " read timeout"
				Else
					retval = 2
					errmsg = " error in 'viRead', code: " & result.ToString("X")
				End If
			Else
				arr = New Byte(cnt - 1) {}
				Array.Copy(buffer, arr, cnt)
			End If
			Return retval
		End Function
	   	
	   	'added field
		Protected termchar As Byte
		Protected termstr As String
	   	'constants
	   
	   	Protected Const VI_SUCCESS_TERM_CHAR As UInteger = &H3FFF0005UI 'if eoi disabled 
	   	
	   	Protected const VI_ATTR_TERMCHAR_EN As UInteger = &H3FFF0038
        Protected const VI_ATTR_TERMCHAR As UInteger = &H3FFF0018
	   	
	   	
	   	
End Class

End Namespace 
