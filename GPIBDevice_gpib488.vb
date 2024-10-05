
Imports System.Text
Imports System.Runtime.InteropServices

Namespace IODevices



    'uses gpib488.dll  used by  MCC and Keithley boards

    ' to change the name of the dll change the line:    Private Const _GPIBDll As String = "gpib488.dll"

    Public Class GPIBDevice_gpib488


        Inherits IODevice


        Protected buffer() As Byte
        Public Property BufferSize() As Integer
            Get
                Return buffer.Length
            End Get
            Set(ByVal value As Integer)
                buffer = New Byte(value) {}
            End Set
        End Property
        Public Property IOTimeoutCode() As Integer 'see GpibConst
            Get
                Return _timeoutcode
            End Get
            Set(ByVal value As Integer)
                _timeoutcode = value
                GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcTMO, value)
            End Set
        End Property

        Protected gpibaddress As Integer
        Protected gpibboard As Integer

        Protected devid As Integer  'device id as returned by ibdev

        Private Shared boardinitialized() As Boolean
        Private _timeoutcode As Integer = GpibDll.GpibConst.T3s


        Shared Sub New()

            boardinitialized = New Boolean(10) {}   'max 10 boards?

            Dim i As Integer
            For i = 0 To boardinitialized.Length - 1 : boardinitialized(i) = False : Next


        End Sub

        Sub New(ByVal name As String, ByVal addr As String)

            MyBase.New(name, addr) 'init base class storing name and addr
            create(name, addr, 32 * 1024)
        End Sub
        Sub New(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer)

            MyBase.New(name, addr) 'init base class storing name and addr
            create(name, addr, defaultbuffersize)
        End Sub

        Sub create(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer)

            statusmsg = "trying to create device '" + name + "' at address " + addr
            IODevice.ParseGpibAddr(addr, gpibboard, gpibaddress)


            If Not boardinitialized(gpibboard) Then
                GpibDll.SendIFC(gpibboard)
                boardinitialized(gpibboard) = True
            End If


            'catchinterfaceexceptions = False  'set when debugging read/write routines

            BufferSize = defaultbuffersize

            interfacelockid = 21
            interfacename = "gpib488"
            statusmsg = ""

            'try to create device

            Try
                devid = GpibDll.ibdev(gpibboard, gpibaddress, 0, _timeoutcode, 1, 0)

                Dim devsta As Integer = GpibDll.ThreadIbsta()
                If (devsta And GpibDll.GpibConst.EERR) <> 0 Then
                    Throw New Exception("cannot get device descriptor on board " & gpibboard)
                End If


                statusmsg = "sending clear to device " & name
                GpibDll.ibclr(devid)


            Catch ex As System.AccessViolationException
                'in this dll happens when USB-GPIB board not connected!!! (however sendIFC has no problem!

                Throw New Exception("exception thrown when trying to create device: GPIB board not connected?")
            End Try
            

            'EOI configuration

            GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcEOSwrt, 0)
            GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcEOT, 1)

            AddToList()
            statusmsg = ""
        End Sub
        Protected Overrides Sub DisposeDevice()

            If devid <> 0 Then
                GpibDll.ibonl(devid, 0)
            End If

        End Sub


        Protected Overrides Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'send cmd, return 0 if ok, 1 if timeout,  other if other error

            Dim retval As Integer
            Dim sta As Integer
            Dim err As Boolean
            Dim tmo As Boolean = False
            Try

                retval = 0

                sta = GpibDll.ibwrt(devid, cmd, Len(cmd))

                err = sta And GpibDll.GpibConst.EERR

                If err Then
                    errcode = GpibDll.ThreadIberr()
                    tmo = (errcode = GpibDll.GpibConst.EABO)
                    If tmo Then
                        retval = 1
                        errmsg = " write timeout"
                    Else
                        retval = 2
                        Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If s <> "" Then
                            errmsg = "error in 'send':" + s
                        Else
                            errmsg = "error in 'send' "
                        End If

                    End If
                End If


            Catch ex As Exception
                err = True
                retval = 2
                errmsg = "exception in ibwrt:\n" + ex.Message
            End Try


            Return retval
        End Function


        '--------------------------
        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer 'poll for status, return MAV bit 
            'spoll,  return 0 if ok, 1 if timeout,  other if other error


            Dim retval As Integer
            Dim sta As Integer
            Dim err As Boolean
            Dim tmo As Boolean = False


            Try  'reading

                retval = 0

                sta = GpibDll.ibrsp(devid, statusbyte)


                err = sta And GpibDll.GpibConst.EERR
                mav = statusbyte And MAVmask  'SerialPollFlags.MessageAvailable


                If err Then
                    errcode = GpibDll.ThreadIberr()
                    tmo = (errcode = GpibDll.GpibConst.EABO)
                    If tmo Then
                        retval = 1
                        errmsg = "serial poll timeout"
                    Else
                    	retval = 2
                    	 Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If s <> "" Then
                            errmsg = "serial poll error:" + s
                        Else
                            errmsg = "serial poll error"
                        End If
                        
                    End If
                End If


            Catch ex As Exception
                retval = 2
                errmsg = "exception in ibrsp:\n" + ex.Message

            End Try

            Return retval

        End Function




        ''--------------------
        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer

            Dim retval As Integer = 0

            Dim err As Boolean
            Dim sta As Integer
            Dim cnt As Integer





            Try  'reading


                cnt = buffer.Length
                sta = GpibDll.ibrd(devid, buffer, cnt)


                err = sta And GpibDll.GpibConst.EERR
                If err Then
                    errcode = GpibDll.ThreadIberr()
                    If errcode = GpibDll.GpibConst.EABO Then
                        retval = 1
                        errmsg = "receive timeout"
                    Else
                        retval = 2
                        Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If s <> "" Then
                            errmsg = s
                        Else
                            errmsg = "error in 'receive' "
                        End If



                    End If
                Else

                    cnt = GpibDll.ThreadIbcnt()
                    arr = New Byte(cnt - 1) {}
                    Array.Copy(buffer, arr, cnt)

                    EOI = sta And GpibDll.GpibConst.EEND
                    retval = 0
                End If

            Catch ex As Exception
                retval = 2
                errmsg = "exception in ibrd:\n" + ex.Message


            End Try

            Return retval


        End Function


        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer


            Dim retval As Integer = 0



            Dim err As Boolean
            Dim sta As Integer


            Try

                sta = GpibDll.ibclr(devid)
                err = sta And GpibDll.GpibConst.EERR

                If err Then
                    errcode = GpibDll.ThreadIberr()
                    retval = 1
                    Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                    If s <> "" Then
                        errmsg = "error in 'cleardevice': " + s
                    Else
                        errmsg = "error in 'cleardevice' "
                    End If

                Else
                    retval = 0
                End If
            Catch ex As Exception
                retval = 1
                errmsg = "exception in ibclr:\n" + ex.Message + "\n cannot clear device "
            End Try

            Return retval
        End Function




        '********************************************************************

        ' dll import functions
        Friend Class GpibDll

            Private Const _GPIBDll As String = "gpib488.dll"



            Protected Friend Class GpibConst
                'status constants
                Public Const EERR = &H8000  ' Error detected
                Public Const TIMO = &H4000  ' 
                Public Const EEND = &H2000  ' EOI or EOS detected
                'some errors

                Public Const EABO As Integer = 6 'Timeout
                Public Const ECIC As Integer = 1     ' Board must be CIC for this function
                Public Const ENOL As Integer = 2 ' no listeners 
                Public Const ENEB As Integer = 7     ' Invalid board specified

                'timeout option
                Public Const T10ms As Integer = 7
                Public Const T30ms As Integer = 8
                Public Const T100ms As Integer = 9
                Public Const T300ms As Integer = 10
                Public Const T1s As Integer = 11
                Public Const T3s As Integer = 12
                Public Const T10s As Integer = 13
                'eot options
                Public Const NULLend As Integer = &H0
                Public Const NLend As Integer = &H1
                Public Const DABend As Integer = &H2

                ' some ibconfig() options
                Public Const IbcPAD As Integer = &H1
                Public Const IbcSAD As Integer = &H2
                Public Const IbcTMO As Integer = &H3
                Public Const IbcEOT As Integer = &H4
               
                Public Const IbcEOSrd As Integer = &HC
                Public Const IbcEOSwrt As Integer = &HD
                Public Const IbcEOScmp As Integer = &HE
                Public Const IbcEOSchar As Integer = &HF

                Public Shared Function errmsg(ByVal errno As Integer) As String
                    Dim s As String = ""

                    Select Case errno 'most common errors
                        Case ECIC : s = "Board is not CIC"
                        Case ENOL : s = "no listeners"
                        Case ENEB : s = "Invalid board specified"

                    End Select
                    Return s
                End Function


            End Class



            <DllImport(_GPIBDll, EntryPoint:="SendIFC")> _
            Private Shared Sub _SendIFC(ByVal board As Integer)
            End Sub
            Protected Friend Shared Sub SendIFC(ByVal board As Integer)
                _SendIFC(board)
            End Sub

            <DllImport(_GPIBDll, EntryPoint:="ibdev")> _
            Private Shared Function _ibdev(ByVal ubrd As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
            End Function
            Protected Friend Shared Function ibdev(ByVal board As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
                Return _ibdev(board, pad, sad, tmo, eot, eos)
            End Function
            <DllImport(_GPIBDll, EntryPoint:="ibonl")> _
             Private Shared Function _ibonl(ByVal ud As Integer, ByVal v As Integer) As UInteger
            End Function
            Protected Friend Shared Function ibonl(ByVal ud As Integer, ByVal v As Integer) As UInteger
                Return _ibonl(ud, v)
            End Function

            <DllImport(_GPIBDll, EntryPoint:="ibconfig")> _
            Private Shared Function _ibconfig(ByVal ud As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
            End Function
            Protected Friend Shared Function ibconfig(ByVal ud As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
                Return _ibconfig(ud, opt, v)
            End Function

            <DllImport(_GPIBDll, EntryPoint:="ibwrt")> _
            Private Shared Function _ibwrt(ByVal ud As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal buf As String, ByVal count As IntPtr) As Integer
            End Function

            'here the IntPtr function is used to detect and adapt to the 32 vs 64 bit environment

            Protected Friend Shared Function ibwrt(ByVal ud As Integer, ByVal buf As String, ByVal count As Integer) As Integer
                Return _ibwrt(ud, buf, New IntPtr(count))
            End Function



            <DllImport(_GPIBDll, EntryPoint:="ibrd")> _
            Private Shared Function _ibrd(ByVal ud As Integer, <MarshalAs(UnmanagedType.LPArray), Out()> ByVal buffer As Byte(), ByVal count As IntPtr) As Integer
            End Function


            Protected Friend Shared Function ibrd(ByVal ud As Integer, ByVal buffer As Byte(), ByVal count As Integer) As Integer

                Return _ibrd(ud, buffer, New IntPtr(count))
            End Function


            <DllImport(_GPIBDll, EntryPoint:="ibclr")> _
            Private Shared Function _ibclr(ByVal ud As Integer) As Integer
            End Function
            Protected Friend Shared Function ibclr(ByVal ud As Integer) As Integer
                Return _ibclr(ud)
            End Function


            <DllImport(_GPIBDll, EntryPoint:="ibrsp")> _
            Private Shared Function _ibrsp(ByVal ud As Integer, ByRef spr As Byte) As Integer
            End Function
            Protected Friend Shared Function ibrsp(ByVal ud As Integer, ByRef spr As Byte) As Integer
                Return _ibrsp(ud, spr)
            End Function

            <DllImport(_GPIBDll, EntryPoint:="ThreadIbsta")> _
            Private Shared Function _ThreadIbsta() As Integer
            End Function
            Protected Friend Shared Function ThreadIbsta() As Integer
                Return _ThreadIbsta()
            End Function

            <DllImport(_GPIBDll, EntryPoint:="ThreadIberr")> _
            Private Shared Function _ThreadIberr() As Integer
            End Function
            Protected Friend Shared Function ThreadIberr() As Integer
                Return _ThreadIberr()
            End Function

            <DllImport(_GPIBDll, EntryPoint:="ThreadIbcnt")> _
            Private Shared Function _ThreadIbcnt() As Integer
            End Function
            Protected Friend Shared Function ThreadIbcnt() As Integer
                Return _ThreadIbcnt()
            End Function




        End Class


    End Class
End Namespace