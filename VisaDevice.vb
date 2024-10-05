
Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Text


Namespace IODevices

     'PW 2020  (last mod: buffer and some constants changed from private to protected)
    'uses "visa32.dll"  which exists in 32bit (windows/syswow64)  and 64 bit((windows/system32) versions
    'dll name defined in the line
    '         private const string _VisaDll = "visa32.dll";   

    Public Class VisaDevice
        Inherits IODevice


        Protected buffer As Byte()
        Public Property BufferSize() As Integer
            Get
                Return buffer.Length
            End Get
            'should not be set during operation

            Set(ByVal value As Integer)
                buffer = New Byte(value - 1) {}
            End Set
        End Property

        Public Property IOTimeout() As UInteger  'get or set the device timeout attribute (in ms)

            Get
                Dim timeout As UInteger
                GetAttribute(VI_ATTR_TMO_VALUE, timeout)
                Return timeout
            End Get
            Set(ByVal value As UInteger)
                Dim timeout As UInteger = value
                SetAttribute(VI_ATTR_TMO_VALUE, timeout)
            End Set
        End Property

        Protected Shared session As Integer = 0
        'resource manager session, common for all devices
        Protected devid As Integer = 0

        Private Const clearonstartup As Boolean = True   'if send clear when creating device


        'Visa constants
        Protected Const VI_errmask As UInteger = &H1111UI

        Protected Const VI_ERROR As UInteger = &HBFFF0000UI
        Protected Const VI_ERROR_TMO As UInteger = &HBFFF0015UI
        Protected Const VI_ERROR_NLISTENERS As UInteger = &HBFFF005FUI
        Protected Const VI_ERROR_RSRC_NFOUND As UInteger = &HBFFF0011UI
        Protected Const VI_GPIB_REN_ADDRESS_GTL As UInt16 = 6

        Protected Const VI_SUCCESS_MAX_CNT As UInteger = &H3FFF0006UI' when EOI not set
        ' variables and constants related to notify
        Protected Const VI_EVENT_SERVICE_REQ As UInteger = &H3FFF200BUI
        Protected Const VI_HNDLR As Integer = 2
        Private userhandle As Integer = 0
        Private cbdelegate As VisaDll.NotifyCallback 'unmanaged callback delegate
        Protected Const VI_ATTR_TMO_VALUE As UInteger = &H3FFF001AUI
        Private CrLf As String = vbCrLf   'C#: "System.Text.Encoding.UTF8.GetString(new byte[2] { 13, 10 }); //CrLf;

        Private _enableNotify As Boolean = False
        Public Overrides Property EnableNotify() As Boolean

            'in Visa each device has to install its own handler

            Get
                Return _enableNotify
            End Get
            Set(ByVal value As Boolean)
                If Not _enableNotify AndAlso value Then
                    cbdelegate = DirectCast(AddressOf cbnotify, VisaDll.NotifyCallback)
                    VisaDll.viInstallHandler(devid, VI_EVENT_SERVICE_REQ, cbdelegate, userhandle)
                    VisaDll.viEnableEvent(devid, VI_EVENT_SERVICE_REQ, VI_HNDLR, 0)
                    _enableNotify = True
                End If
                If _enableNotify AndAlso Not value Then
                    VisaDll.viDisableEvent(devid, VI_EVENT_SERVICE_REQ, VI_HNDLR)
                    VisaDll.viUninstallHandler(devid, VI_EVENT_SERVICE_REQ, cbdelegate, userhandle)
                    _enableNotify = False
                End If
            End Set
        End Property

        'handler for Notify

        ' C prototype:
        'ViStatus _VI_FUNCH viEventHandler(ViSession vi, ViEventType eventType, ViEvent context, ViAddr userHandle)
        ' _VI_FUNCH :  __stdcall in 32 , __fastcall in 64

        Private Function cbnotify(ByVal vid As Integer, ByVal eventType As UInteger, ByVal context As UInteger, ByRef userHandle As Integer) As UInteger
            If vid = devid And eventType = VI_EVENT_SERVICE_REQ Then
                WakeUp()
            End If
            Thread.Sleep(1) 'yield before rearming          
            Return 0

        End Function

        'constructors
        Public Sub New(ByVal name As String, ByVal addr As String)
            MyBase.New(name, addr)
            create(name, addr, 32 * 1024, false)
        End Sub



        Public Sub New(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer)
            MyBase.New(name, addr)
            create(name, addr, defaultbuffersize, false)
        End Sub
        
        
        'constructor overloads added in Dec 2018:
        'constructor
        Public Sub New(ByVal name As String, ByVal addr As String, ByVal interlocked As Boolean)
            MyBase.New(name, addr)
            create(name, addr, 32 * 1024, interlocked )
        End Sub



        Public Sub New(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer, ByVal interlocked As Boolean)
            MyBase.New(name, addr)
            create(name, addr, defaultbuffersize,  interlocked)
        End Sub
        
        

'common part of constructor
' modified Dec 2018:"interlocked" parameter added 

        Private Sub create(ByVal name As String, ByVal addr As String, ByVal defaultbuffersize As Integer, ByVal interlocked As Boolean)

            Dim accessmode As Integer = 0
            Dim opentimeout As Integer = 300

            Dim result As UInteger = 0


            'open a session

            If session = 0 Then
                statusmsg = "opening Visa session..."

                result = VisaDll.viopenDefaultRM(session)
                If result >= VI_ERROR Then
                    Throw New Exception("cannot open Visa session, code " & result.ToString("X"))
                End If
            End If


            'try to create device

            statusmsg = "trying to create device '" & name & "' at address " & addr

            result = VisaDll.viOpen(session, addr, accessmode, opentimeout, devid)


            If result >= VI_ERROR Then
                Dim msg As String = "could not open Visa device at address " & addr

                If (result = VI_ERROR_NLISTENERS) OrElse (result = VI_ERROR_RSRC_NFOUND) Then
                    msg += CrLf & "no devices detected at this adress"
                End If

                msg += CrLf & "error code: " & result.ToString("X")

                Throw New System.Exception(msg)
            End If

            If clearonstartup Then
                statusmsg = "sending clear to device " & name
                result = VisaDll.viClear(devid)


                If result >= VI_ERROR Then
                    Dim msg As String = "could not clear Visa device at address " & addr
                    If result = VI_ERROR_NLISTENERS Then
                        msg += CrLf & "no listeners detected at this adress"
                    Else
                        msg += CrLf & "error code " & result.ToString("X")
                    End If
                    'Throw New Exception(msg)           ' Removed 11/03/22 an attempt to stop viClear errors
                End If
            End If

            'catchinterfaceexceptions = False  'set when debugging read/write routines

            BufferSize = defaultbuffersize

'interfacelockid modified in 2018:
         
          If Not interlocked then 'standard configuration
            If addr.ToUpper().Contains("GPIB") Then
                Dim gpibboard As Integer
                Dim gpibaddr As Byte
                IODevice.ParseGpibAddr(addr, gpibboard, gpibaddr)
                interfacelockid = gpibboard + 10
            Else
                interfacelockid = -1 'no interface lock for non-gpib interfaces
            End If
          Else 'interlocked configuration
                interfacelockid = 25
            End If	
          	
            interfacename = "Visa"
            statusmsg = ""
            AddToList()


        End Sub






        Protected Overrides Sub DisposeDevice()

            'release unmanaged resources
            EnableNotify = False
            If devid <> 0 Then
            	VisaDll.viGpibControlREN(devid, VI_GPIB_REN_ADDRESS_GTL)
                VisaDll.viClose(devid)
                devid = 0
            End If



        End Sub


        'finalizer: now implemented in the base class
        ' Protected Overrides Sub Finalize()
           

        Protected Overrides Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'send cmd, return 0 if ok, 1 if timeout,  other if other error

            Dim retval As Integer = 0

            Dim err As Boolean = False
            Dim tmo As Boolean = False
            Dim resultwrite As UInteger = 0
            Dim retcount As Integer = 0

            retval = 0

            resultwrite = VisaDll.viWrite(devid, cmd, retcount)

            err = (resultwrite >= VI_ERROR) Or (cmd.Length <> retcount)
            tmo = (resultwrite = VI_ERROR_TMO)

            If err Then
                errcode = Convert.ToInt32(resultwrite And VI_errmask)
                If tmo Then
                    retval = 1
                    errmsg = " write timeout"
                Else
                    retval = 2
                    errmsg = " error in 'viWrite', "
                    If resultwrite = VI_ERROR_NLISTENERS Then
                        errmsg += "no listeners detected at this adress"
                    Else
                        errmsg += "error code " & resultwrite.ToString("X")
                    End If

                End If
            End If

            Return retval
        End Function


        '--------------------------
        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            'poll for status, return MAV bit 
            'spoll,  return 0 if ok, 1 if timeout,  other if other error

            Dim retval As Integer = 0
            Dim result As UInteger = 0
            Dim err As Boolean = False
            Dim tmo As Boolean = False
            Dim stb As Short = 0


            retval = 0
            result = VisaDll.viReadSTB(devid, stb)
            statusbyte = Convert.ToByte(stb And 255)
            mav = (statusbyte And MAVmask) <> 0
            'SerialPollFlags.MessageAvailable=16
            err = (result > VI_ERROR)
            tmo = (result = VI_ERROR_TMO)

            'status=1 tmo on send,  =3 tmo on rcv, =4 other err on send, =6 other err on rcv


            If err Then
                errcode = Convert.ToInt32(result And VI_errmask)

                If tmo Then
                    retval = 1
                    errmsg = "serial poll timeout"
                Else
                    retval = 2
                    errmsg = "serial poll error, code: " & result.ToString("X")
                End If
            End If

            Return retval

        End Function

        ''--------------------
        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer

            Dim retval As Integer = 0

            Dim err As Boolean = False
            Dim result As UInteger = 0

            Dim cnt As Integer = 0
            Dim tmo As Boolean = False

            retval = 0
            result = VisaDll.viRead(devid, buffer, buffer.Length, cnt)

            err = Not (result = 0 Or result = VI_SUCCESS_MAX_CNT)

            EOI = result = 0

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


        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Dim retval As Integer = 0
            Dim result As UInteger = 0


            retval = 0


            result = VisaDll.viClear(devid)


            If result <> 0 Then
                retval = 1
                errcode = Convert.ToInt32(result And VI_errmask)
                errmsg = "error in viClear, code: " & result.ToString("X")
            End If

            Return retval
        End Function

        'other functions specific to visa : attributes

        Public Function SetAttribute(ByVal attribute As UInteger, ByVal attrState As Integer) As UInteger
            Return VisaDll.viSetAttribute_Int32(devid, attribute, attrState)
        End Function
        Public Function SetAttribute(ByVal attribute As UInteger, ByVal attrState As UInteger) As UInteger
            Return VisaDll.viSetAttribute_UInt32(devid, attribute, attrState)
        End Function



        Public Function GetAttribute(ByVal attribute As UInteger, ByRef attrState As Integer) As UInteger
            Return VisaDll.viGetAttribute_Int32(devid, attribute, attrState)
        End Function
        Public Function GetAttribute(ByVal attribute As UInteger, ByRef attrState As UInteger) As UInteger
            Return VisaDll.viGetAttribute_UInt32(devid, attribute, attrState)
        End Function



    End Class


    '********************************************************************

    ' dll import functions
    Class VisaDll


        Private Const _VisaDll As String = "visa32.dll"

        <DllImport(_VisaDll, EntryPoint:="viOpenDefaultRM")> _
        Private Shared Function _viopenDefaultRM(ByRef sesn As Integer) As UInteger
        End Function
        Public Shared Function viopenDefaultRM(ByRef sesn As Integer) As UInteger
            Return _viopenDefaultRM(sesn)
        End Function

        <DllImport(_VisaDll, EntryPoint:="viOpen")> _
        Private Shared Function _viopen(ByVal sesn As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal rsrcName As String, ByVal accessMode As Integer, ByVal openTimeout As Integer, ByRef v As Integer) As UInteger
        End Function
        Public Shared Function viOpen(ByVal sesn As Integer, ByVal rsrcName As String, ByVal accessMode As Integer, ByVal openTimeout As Integer, ByRef v As Integer) As UInteger

            Return _viopen(sesn, rsrcName, accessMode, openTimeout, v)

        End Function
        <DllImport(_VisaDll, EntryPoint:="viWrite")> _
        Private Shared Function _viWrite(ByVal vi As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal buf As String, ByVal maxcount As Integer, ByRef retcount As Integer) As UInteger
        End Function
        Public Shared Function viWrite(ByVal vi As Integer, ByVal cmd As String, ByRef retcount As Integer) As UInteger

            Return _viWrite(vi, cmd, cmd.Length, retcount)


        End Function


        <DllImport(_VisaDll, EntryPoint:="viRead")> _
        Private Shared Function _viRead(ByVal vi As Integer, <MarshalAs(UnmanagedType.LPArray), Out()> ByVal buf As Byte(), ByVal maxcount As Integer, ByRef retcount As Integer) As UInteger
        End Function
        Public Shared Function viRead(ByVal vi As Integer, ByVal buf As Byte(), ByVal maxcount As Integer, ByRef cnt As Integer) As UInteger


            Return _viRead(vi, buf, maxcount, cnt)

        End Function

        <DllImport(_VisaDll, EntryPoint:="viClear")> _
        Private Shared Function _viClear(ByVal vid As Integer) As UInteger
        End Function
        Public Shared Function viClear(ByVal vid As Integer) As UInteger
            Return _viClear(vid)
        End Function


        <DllImport(_VisaDll, EntryPoint:="viReadSTB")> _
        Private Shared Function _viReadSTB(ByVal vid As Integer, ByRef stb As Short) As UInteger
        End Function
        Public Shared Function viReadSTB(ByVal vid As Integer, ByRef stb As Short) As UInteger
            Return _viReadSTB(vid, stb)
        End Function


        <DllImport(_VisaDll, EntryPoint:="viClose")> _
        Private Shared Function _viClose(ByVal vid As Integer) As UInteger
        End Function
        Public Shared Function viClose(ByVal vid As Integer) As UInteger
            Return _viClose(vid)
        End Function
        'set/get attribute defined for most common attribute types: Int32,UInt32 (may be extended if needed for other types)

        <DllImport(_VisaDll, EntryPoint:="viSetAttribute")> _
        Friend Shared Function viSetAttribute_Int32(ByVal vid As Integer, ByVal attribute As UInteger, ByVal attrState As Integer) As UInteger
        End Function


        <DllImport(_VisaDll, EntryPoint:="viSetAttribute")> _
        Friend Shared Function viSetAttribute_UInt32(ByVal vid As Integer, ByVal attribute As UInteger, ByVal attrState As UInteger) As UInteger
        End Function


        <DllImport(_VisaDll, EntryPoint:="viGetAttribute")> _
        Friend Shared Function viGetAttribute_Int32(ByVal vid As Integer, ByVal attribute As UInteger, ByRef attrState As Integer) As UInteger
        End Function

        <DllImport(_VisaDll, EntryPoint:="viGetAttribute")> _
        Friend Shared Function viGetAttribute_UInt32(ByVal vid As Integer, ByVal attribute As UInteger, ByRef attrState As UInteger) As UInteger
        End Function


        '  event handler functions

        ' C prototype for callback handler:
        'ViStatus _VI_FUNCH viEventHandler(ViSession vi, ViEventType eventType, ViEvent context, ViAddr userHandle)
        ' _VI_FUNCH :  __stdcall in 32 , __fastcall in 64
        <UnmanagedFunctionPointer(CallingConvention.StdCall)> _
        Public Delegate Function NotifyCallback(ByVal vid As Integer, ByVal eventType As UInteger, ByVal context As UInteger, ByRef userHandle As Integer) As UInteger
        'userHandle not used here

        'ViStatus viInstallHandler(ViSession vi, ViEventType eventType, ViHndlr handler, ViAddr userHandle)
        <DllImport(_VisaDll, EntryPoint:="viInstallHandler")> _
        Public Shared Function viInstallHandler(ByVal vid As Integer, ByVal eventType As UInteger, <MarshalAs(UnmanagedType.FunctionPtr)> ByVal callback As NotifyCallback, ByRef userHandle As Integer) As UInteger
        End Function

        '  ViStatus viUninstallHandler(ViSession vi, ViEventType eventType, ViHndlr handler, ViAddr userHandle)

        <DllImport(_VisaDll, EntryPoint:="viUninstallHandler")> _
        Public Shared Function viUninstallHandler(ByVal vid As Integer, ByVal eventType As UInteger, <MarshalAs(UnmanagedType.FunctionPtr)> ByVal callback As NotifyCallback, ByRef userHandle As Integer) As UInteger
        End Function

        'ViStatus viEnableEvent(ViSession vi, ViEventType eventType, ViUInt16 mechanism, ViEventFilter context)


        <DllImport(_VisaDll, EntryPoint:="viEnableEvent")> _
        Public Shared Function viEnableEvent(ByVal vid As Integer, ByVal eventType As UInteger, ByVal mechanism As UInt16, ByVal context As UInteger) As UInteger
        End Function

        'ViStatus viDisableEvent(ViSession vi, ViEventType eventType, ViUInt16 mechanism)
        <DllImport(_VisaDll, EntryPoint:="viDisableEvent")> _
        Public Shared Function viDisableEvent(ByVal vid As Integer, ByVal eventType As UInteger, ByVal mechanism As UInt16) As UInteger
        End Function
        
        
        'ViStatus  viGpibControlREN(ViSession vi, ViUInt16 mode);
   
        
         <DllImport(_VisaDll, EntryPoint:="viGpibControlREN")> _
         Friend Shared Function viGpibControlREN(ByVal vid As Integer, ByVal mode As UInt16) As UInteger
        End Function

        
    End Class

End Namespace
