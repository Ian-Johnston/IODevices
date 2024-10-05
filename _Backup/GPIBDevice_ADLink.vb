
Imports System

Imports System.Data
Imports System.Diagnostics
Imports System.Collections
Imports System.Collections.Generic

Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Threading

Namespace IODevices


    'uses gpib-32.dll  from ADLINK (but call syntax compatible with the NI dll of the same name: can use this library for NI boards too)
    'this name conflicts with old versions of dlls by NI, Keithley or MCC 
    ' to change the name of the dll change the line:    Private Const _GPIBDll As String = "gpib-32.dll"



    Public Class GPIBDevice_ADLink
        Inherits IODevice


        Protected buffer As Byte()
        Public Property BufferSize() As Integer
            Get
                Return buffer.Length
            End Get
            Set(ByVal value As Integer)
                buffer = New Byte(value) {}
            End Set
        End Property

        Public Property IOTimeoutCode() As Integer
            'see GpibConst
            Get
                Return _timeoutcode
            End Get
            Set(ByVal value As Integer)
                _timeoutcode = value
                GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcTMO, value)
            End Set
        End Property

        Protected gpibaddress As Byte

        Protected gpibboard As Integer
        'device id as returned by ibdev
        Protected devid As Integer = 0

        Private _timeoutcode As Integer = GpibDll.GpibConst.T3s     '  default value



        'variables and code used by notify
        Private Shared boardud As Integer()
        'board unit descriptors for all boards
        Public Shared delaynotify As Integer = 5
        'delay before rearming
        Private Shared notifylist As List(Of GPIBDevice_ADLink)
        Private Shared locklist As New Object()
        'list of devices to notify
        Private Shared notifymask As Integer()
        ' masks for each board
        Private Shared userdata As Integer
        'not used 
        Private Shared cbdelegate As GpibDll.NotifyCallback

        Private _enableNotify As Boolean = False


        Public Overrides Property EnableNotify() As Boolean

            'for pending notify events


            Get
                Return _enableNotify
            End Get
            Set(ByVal value As Boolean)
                If Not _enableNotify AndAlso value Then
                    SyncLock locklist
                        If notifycount(gpibboard) = 0 Then
                            notifymask(gpibboard) = GpibDll.GpibConst.SRQI
                            GpibDll.ibnotify(boardud(gpibboard), GpibDll.GpibConst.SRQI, cbdelegate, userdata)
                        End If
                        notifylist.Add(Me)
                    End SyncLock
                    _enableNotify = True
                End If
                If _enableNotify AndAlso Not value Then
                    SyncLock locklist
                        notifylist.Remove(Me)
                        If notifycount(gpibboard) = 0 Then
                            notifymask(gpibboard) = 0 'let pending notify events finish
                            Thread.Sleep(delaynotify)
                            GpibDll.ibnotify(boardud(gpibboard), 0, cbdelegate, userdata)
                        End If
                    End SyncLock
                    _enableNotify = False
                End If
            End Set
        End Property


        Private Shared Function notifycount(ByVal boardnum As Integer) As Integer
            Dim count As Integer = 0
            For Each device As GPIBDevice_ADLink In notifylist
                If boardnum = device.gpibboard Then
                    count += 1
                End If
            Next

            Return count
        End Function

        'notify callback:  public delegate uint NotifyCallback(int ud, int ibsta, int iberr, int ibcnt, [MarshalAs(UnmanagedType.AsAny)] object RefData);//refdata not used here

        Public Shared Function cbnotify(ByVal ud As Integer, ByVal ibsta As Integer, ByVal iberr As Integer, ByVal ibcnt As Integer, ByRef RefData As Integer) As Integer
            'refdata not used here
            Dim board As Integer = 0
            Dim retval As Integer

            SyncLock locklist
                For Each device As GPIBDevice_ADLink In notifylist
                    If ud = boardud(device.gpibboard) Then
                        device.WakeUp() 'interrupt waiting for next read/poll trial 

                        board = device.gpibboard  'identify board n°
                    End If
                Next
                retval = notifymask(board)
            End SyncLock

            If retval <> 0 Then Thread.Sleep(delaynotify) 'delay before rearming
            Return retval 'return mask to rearm for next notify (or 0 if disabling : safer if notify disabled while there are pending events)
        End Function



        'static constructor
        Shared Sub New()
            Const maxboards As Integer = 10

            boardud = New Integer(maxboards - 1) {}
            notifymask = New Integer(maxboards - 1) {}

            Dim i As Integer = 0
            For i = 0 To boardud.Length - 1
                boardud(i) = 0
                notifymask(i) = 0
            Next

            notifylist = New List(Of GPIBDevice_ADLink)()

            cbdelegate = DirectCast(AddressOf cbnotify, GpibDll.NotifyCallback)
        End Sub



        Public Sub New(ByVal name As String, ByVal addr As String)
            MyBase.New(name, addr)
            'init base class storing name and addr
            create(name, addr, 32 * 1024)
        End Sub

        Public Sub New(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer)
            MyBase.New(name, addr)
            'init base class storing name and addr
            create(name, addr, defaultbuffersize)
        End Sub

        'common part of constructor

        Private Sub create(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer)



            IODevice.ParseGpibAddr(addr, gpibboard, gpibaddress)

            Try

                If boardud(gpibboard) <= 0 Then
                    Dim boardname As String = "GPIB" & gpibboard.ToString().Trim()

                    statusmsg = "trying to initialize " & boardname
                    Dim bud As Integer = GpibDll.ibfind(boardname)

                    If bud <> -1 Then
                        boardud(gpibboard) = bud
                    Else
                        Throw New Exception("cannot find GPIB board n°" & gpibboard)
                    End If

                    'GpibDll.SendIFC(gpibboard);   //some devices don't like it...


                    GpibDll.ibconfig(boardud(gpibboard), GpibDll.GpibConst.IbcAUTOPOLL, 0) ' disable autopolling
                End If


                'catchinterfaceexceptions = False  'set when debugging read/write routines

                BufferSize = defaultbuffersize

                interfacelockid = 20

                interfacename = "ADLink"


                'try to create device


                statusmsg = "trying to create device '" & name & "' at address " & addr
                devid = GpibDll.ibdev(gpibboard, gpibaddress, 0, _timeoutcode, 1, 0)

                Dim devsta As Integer = GpibDll.ThreadIbsta()
                If (devsta And GpibDll.GpibConst.EERR) <> 0 Then
                    Throw New Exception("cannot get device descriptor on board " & gpibboard)
                End If


                statusmsg = "sending clear to device " & name

                GpibDll.ibclr(devid)

            Catch ex As System.AccessViolationException 'in this dll may happen when USB-GPIB board not connected!!! (however sendIFC has no problem!)

                Throw New Exception("exception thrown when trying to create device: GPIB board not connected?")
            End Try


            'EOI configuration

            Dim sta As Integer = 0
            sta = GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcEOSwrt, 0)
            sta = GpibDll.ibconfig(devid, GpibDll.GpibConst.IbcEOT, 1)

            AddToList()
            statusmsg = ""
        End Sub

        Protected Overrides Sub DisposeDevice()
            If devid <> 0 Then
                EnableNotify = False
                GpibDll.ibonl(devid, 0)
            End If

        End Sub


        Protected Overrides Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'send cmd, return 0 if ok, 1 if timeout,  other if other error

            Dim retval As Integer = 0
            Dim sta As Integer = 0
            Dim err As Boolean = False
            Dim tmo As Boolean = False

            Try
                retval = 0

                sta = GpibDll.ibwrt(devid, cmd, cmd.Length)


                err = (sta And GpibDll.GpibConst.EERR) <> 0

                If err Then
                    errcode = GpibDll.ThreadIberr()
                    tmo = (errcode = GpibDll.GpibConst.EABO)
                    If tmo Then
                        retval = 1
                        errmsg = " write timeout"
                    Else
                        retval = 2

                        Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If Not String.IsNullOrEmpty(s) Then
                            errmsg = " error in 'send':" & s
                        Else
                            errmsg = " error in 'send'"
                        End If
                    End If



                End If
            Catch ex As Exception
                err = True
                retval = 2
                errmsg = "exception in ibwrt:\n" & ex.Message
            End Try


            Return retval
        End Function


        '--------------------------
        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'poll for status, return MAV bit 
            'spoll,  return 0 if ok, 1 if timeout,  other if other error


            Dim retval As Integer = 0
            Dim sta As Integer = 0
            Dim err As Boolean = False
            Dim tmo As Boolean = False


            'reading
            Try

                retval = 0

                sta = GpibDll.ibrsp(devid, statusbyte)


                err = (sta And GpibDll.GpibConst.EERR) <> 0
                mav = (statusbyte And MAVmask) <> 0

                'status=1 tmo on send,  =3 tmo on rcv, =4 other err on send, =6 other err on rcv


                If err Then
                    errcode = GpibDll.ThreadIberr()
                    tmo = (errcode = GpibDll.GpibConst.EABO)
                    If tmo Then
                        retval = 1
                        errmsg = "serial poll timeout"
                    Else
                        retval = 2
                        Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If Not String.IsNullOrEmpty(s) Then
                            errmsg = "serial poll error:" & s
                        Else
                            errmsg = "serial poll error"

                        End If
                    End If


                End If
            Catch ex As Exception
                retval = 2

                errmsg = ex.Message
            End Try

            Return retval

        End Function




        ''--------------------
        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer


            Dim retval As Integer = 0

            Dim err As Boolean = False
            Dim sta As Integer = 0
            Dim cnt As Integer = 0



            'reading
            Try


                cnt = buffer.Length
                sta = GpibDll.ibrd(devid, buffer, cnt)


                err = (sta And GpibDll.GpibConst.EERR) <> 0
                If err Then
                    errcode = GpibDll.ThreadIberr()
                    If errcode = GpibDll.GpibConst.EABO Then
                        retval = 1
                        errmsg = "receive timeout"
                    Else
                        retval = 2
                        Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                        If Not String.IsNullOrEmpty(s) Then
                            errmsg = s
                        Else
                            errmsg = "error in 'receive' "
                        End If

                    End If
                Else

                    cnt = GpibDll.ThreadIbcnt()
                    arr = New Byte(cnt - 1) {}
                    Array.Copy(buffer, arr, cnt)

                    EOI = (sta And GpibDll.GpibConst.EEND) <> 0
                    retval = 0

                End If
            Catch ex As Exception
                retval = 2


                errmsg = ex.Message
            End Try

            Return retval


        End Function


        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer

            Dim retval As Integer = 0


            Dim err As Boolean = False
            Dim sta As Integer = 0


            Try
                sta = GpibDll.ibclr(devid)


                err = (sta And GpibDll.GpibConst.EERR) <> 0

                If err Then
                    errcode = GpibDll.ThreadIberr()
                    retval = 1
                    Dim s As String = GpibDll.GpibConst.errmsg(errcode)
                    If Not String.IsNullOrEmpty(s) Then
                        errmsg = "error in 'cleardevice': " & s
                    Else
                        errmsg = "error in 'cleardevice' "
                    End If
                Else
                    retval = 0
                End If
            Catch ex As Exception
                retval = 1
                errmsg = ex.Message & vbLf & " cannot clear device "
            End Try

            Return retval
        End Function


        '********************************************************************

        ' dll import functions
        Friend Class GpibDll

            Private Const _GPIBDll As String = "gpib-32.dll"


            Protected Friend Class GpibConst

                'status constants
                Public Const EERR = &H8000  ' Error detected
                Public Const TIMO = &H4000  ' 
                Public Const EEND = &H2000  ' EOI or EOS detected

                'some errors

                Public Const EABO As Integer = 6 'Timeout
                Public Const ECIC As Integer = 1     ' Board must be CIC for this function
                Public Const ENOL As Integer = 2 ' no listeners 
                Public Const EADR As Integer = 3     ' Board not addressed correctly
                Public Const ENEB As Integer = 7     ' Invalid board specified
                Public Const EBUS As Integer = 14    ' Command error on bus

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

                'some  ibconfig() options
                Public Const IbcPAD As Integer = &H1
                Public Const IbcSAD As Integer = &H2
                Public Const IbcTMO As Integer = &H3
                Public Const IbcEOT As Integer = &H4
                Public Const IbcPPC As Integer = &H5

                Public Const IbcEOSrd As Integer = &HC
                Public Const IbcEOSwrt As Integer = &HD
                Public Const IbcEOScmp As Integer = &HE
                Public Const IbcEOSchar As Integer = &HF
                Public Const IbcAUTOPOLL As Integer = &H7

                Public Const SRQI As Integer = &H1000'mask for SRQ board level notify

                Public Shared Function errmsg(ByVal errno As Integer) As String
                    Dim s As String = ""

                    Select Case errno
                        'most common errors
                        Case ECIC
                            s = "Board is not CIC"
                            Exit Select
                        Case ENOL
                            s = "no listeners"
                            Exit Select
                        Case ENEB
                            s = "Invalid board specified"
                            Exit Select
                        Case EADR
                            s = "Board not addressed correctly"
                            Exit Select
                        Case EBUS
                            s = "Command error on bus"

                            Exit Select
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


            <DllImport(_GPIBDll, EntryPoint:="ibfind")> _
            Private Shared Function _ibfind(<MarshalAs(UnmanagedType.LPStr)> ByVal name As String) As Integer
            End Function
            Protected Friend Shared Function ibfind(ByVal name As String) As Integer
                Return _ibfind(name)
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
            Private Shared Function _ibwrt(ByVal ud As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal buf As String, ByVal count As Integer) As Integer
            End Function
            Protected Friend Shared Function ibwrt(ByVal ud As Integer, ByVal buf As String, ByVal count As Integer) As Integer
                Return _ibwrt(ud, buf, count)
            End Function
            <DllImport(_GPIBDll, EntryPoint:="ibrd")> _
            Private Shared Function _ibrd(ByVal ud As Integer, <MarshalAs(UnmanagedType.LPArray), Out()> ByVal buffer As Byte(), ByVal count As Integer) As Integer
            End Function

            Protected Friend Shared Function ibrd(ByVal ud As Integer, ByVal buffer As Byte(), ByVal count As Integer) As Integer

                Return _ibrd(ud, buffer, count)
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


            ' notify  event handler functions

            ' C prototype for callback handler:
            'int __stdcall Callback (int ud,int ibsta,int iberr,long ibcntl,void * RefData)
            <UnmanagedFunctionPointer(CallingConvention.StdCall)> _
            Public Delegate Function NotifyCallback(ByVal ud As Integer, ByVal ibsta As Integer, ByVal iberr As Integer, ByVal ibcntl As Integer, ByRef RefData As Integer) As Integer
            'refdata not used here


            'uint ibnotify (int ud,int mask,GpibNotifyCallback_t Callback,void * RefData)
            <DllImport(_GPIBDll, EntryPoint:="ibnotify")> _
            Public Shared Function ibnotify(ByVal ud As Integer, ByVal mask As Integer, <MarshalAs(UnmanagedType.FunctionPtr)> ByVal callback As NotifyCallback, ByRef RefData As Integer) As UInteger
            End Function




        End Class


    End Class

End Namespace
