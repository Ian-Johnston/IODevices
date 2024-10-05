'Class IODevice:  multithreaded GPIB/Visa/serial communication
'(C)P.Wzietek 2016




'exported classes:
'IODevice   'abstract (MustInherit) class from which real devices are derived
'IOQuery   'class passed to callback functions

'all definitions contained in "IODevices" namespace (in VB do not define a "default" namespace for this project)


' exported IODevice class methods:

'shared methods:
'Shared Function GetDeviceList
'Shared Function DeviceByName
'ShowDevices
'  Shared Sub ParseGpibAddr    used by constructor


'instance methods:


'1st version: call callback function when data ready
' 'signature of callback functions:
'  Public Delegate Sub IOCallback(ByVal q As IOQuery)  



'standard:
'Public Function QueryAsync(ByVal cmd As String, _
'                       ByVal callback As IOCallback, ByVal retry As Boolean) As Integer

'complete:

'Public Function QueryAsync(ByVal cmd As String, _
'                      ByVal callback As IOCallback, ByVal retry As Boolean, ByVal cbwait As Boolean, _
'                       ByVal tag As Integer) As Integer

'2nd version : update textbox with data string
'Public Function QueryAsync(ByVal cmd As String, ByVal text As TextBox, _
'                         ByVal retry As Boolean) As Integer

'Public Function QueryAsync(ByVal cmd As String, ByVal text As TextBox, _
'                    ByVal retry As Boolean, ByVal tag As Integer) As Integer


'send command, no response:

'default behavior :
'Public Function SendAsync(ByVal cmd As String, ByVal retry As Boolean)

'complete (with callback)
'Public Function SendAsync(ByVal cmd As String, _
'                    ByVal callback As IOCallback, ByVal retry As Boolean, _
'                    ByVal cbwait As Boolean, ByVal tag As Integer) As Integer






'blocking versions  ********************

'send command and wait for response

'Public Function QueryBlocking(ByVal cmd As String, ByRef q As IOQuery, ByVal retry As Boolean) As Integer 
'Public Function QueryBlocking(ByVal cmd As String, ByRef resp As String, ByVal retry As Boolean) As Integer 
'Public Function QueryBlocking(ByVal cmd As String, ByRef resparr As Byte(), ByVal retry As Boolean) As Integer 


'send command, no response

'Public Function SendBlocking(ByVal cmd As String, ByVal retry As Boolean) As Integer

'other functions

'    Public Function IsBlocking() As Boolean
'        true when blocking call in progress

'    Public Function PendingTasks() As Integer
'      return number of queries in the queue
'    Public Function PendingTasks(ByVal cmd As String) As Integer
'     'same for a specific command: number of copies of specific command in the queue

'    Public Sub WaitAsync()
'       waits until queries queued before the call are done (not until the queue is empty - this may never happen if queries called by timers)

'    Public Sub AbortAllTasks()

'    Public Sub AddToList()  'register device in "devicelist",  should be called by child class constructors (is not called in the base class constructor to avoid registering ill-defined objects when error)

'version 2: adds possibility to use asynchronous notifying:

' Protected Sub WakeUp()   interrupts waiting for next reading or poll trial



Imports System.Runtime.InteropServices
Imports System.Drawing

Imports IODeviceForms
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Threading
Imports System.Reflection
Imports System.Windows.Forms

'Imports System.Runtime.ExceptionServices  'tentative patch for accessviolation exception (blackdog tests)

Namespace IODevices



    Public Class IOQuery
        Public cmd As String
        'query identifier 
        Public tag As Integer
        Public ReadOnly Property ResponseAsString() As String
            'response as string or byte arr depending on query type
            Get
                Return task.parent.ByteArrayToString(resparr)
            End Get
        End Property

        Public ReadOnly Property ResponseAsByteArray() As Byte()
            Get
                Return resparr
            End Get
        End Property

        Public status As Integer
        '0:ok, otherwise combination:
        'bit 1:timeout, bit 2 on send(0)/recv(1), bit 3 : other error (see errcode), , bit 4: aborted, bit 5: poll error, bit 8 callback error
        'so if not aborted: status=1 tmo on send; =3 tmo on rcv, =4 other err on send, =6 other err on rcv, 
        ' if aborted add 8 , if  poll timeout add 16, 

        'interface error code (if status>0)
        Public errcode As Integer
        'error message
        Public errmsg As String
        'when function called
        Public timecall As DateTime
        'when device unlocked and operation started
        Public timestart As DateTime
        Public timeend As DateTime
        'query type :  1: without response (cmd only) '2: with response
        Public type As Integer
        Public ReadOnly Property device() As IODevice
            Get
                If task IsNot Nothing Then
                    Return task.parent
                Else
                    Return Nothing
                End If
            End Get
        End Property

        'abort this task(async or blocking)
        Public Sub AbortRetry()
            task.abort = True
        End Sub
        'abort all queued async commands and active blocking command
        Public Sub AbortAll()
            task.parent.AbortAllTasks()

        End Sub

        'private fields
        Protected Friend resparr As Byte()
        'used to access  task fields (abort etc)
        Friend task As IODevice.IOTask

        'constructor
        Friend Sub New(ByVal qtype As Integer, ByVal command As String, ByVal qtag As Integer)

            cmd = command
            resparr = Nothing
            type = qtype
            errmsg = ""
            timestart = DateTime.MinValue
            timeend = DateTime.MinValue
            timecall = DateTime.MinValue

            task = Nothing
        End Sub
    End Class
    'end of IOQuery


    '************             class IODevice                  *******************************
    '*****************************************************************************

    Public MustInherit Class IODevice
        Implements IDisposable

        Public Delegate Sub IOCallback(ByVal q As IOQuery)
        'signature of callback functions



        '************class IODevice public  variables



        'optional message (status etc.) to display in devices form (eg used by constructors during init)
        Public Shared statusmsg As String


        '-------------------device instance variables

        'used to limit the number of threads 

        Public maxtasks As Integer
        'device name
        Public devname As String

        Public devaddr As String

        'delay to wait between operations (ms)
        Public delayop As Integer
        'default delay between cmd and read :
        Public delayread As Integer
        ' delay before poll/read, especially useful to avoid blocking gpib bus by slow devices when polling is not available

        'default delay before retrying read after timeout
        Public delayrereadontimeout As Integer
        '                                        or delay between polls if polling used 

        'delayread may be overwritten on per task basis
        'cumulative timeout for read (to replace effect of delayread)
        Public readtimeout As Integer
        'delay before retry on timeout
        Public delayretry As Integer

        'if true repeat read if EOI not detected (eg. buffer too small)
        Public checkEOI As Boolean
        'if true serial poll before reading to not to block bus
        Public enablepoll As Boolean
        Public MAVmask As Byte = 16  'for GPIB, USBTMC-USB488, VXI-11: standard (488.2) mask for MAV status (bit 5 of the status byte), change it for devices not compliant with 488.2 

        'remove crlf in ByteArrayToString function 
        Public stripcrlf As Boolean
        'for blocking commands during delays and retry loop
        Public eventsallowed As Boolean


        Public showmessages As Boolean ' error window enabled
        Public Const showdevicesonstartup As Boolean = True


        Public lastasyncquery As IOQuery         '  volatile
        'set to false when debugging a new interface
        Public catchinterfaceexceptions As Boolean
        Public catchcallbackexceptions As Boolean
        'if callback on each retry
        Public callbackonretry As Boolean

        'to override when available
        Public Overridable Property EnableNotify() As Boolean
            Get
                Return False
            End Get

            Set(ByVal value As Boolean)
                If value Then
                    Throw New NotImplementedException("Notify not implemented for interface '" + interfacename + "'")
                End If
            End Set
        End Property

        '*****************    private IODevice variables        *********************************************************
        '   

        'shared variables and functions
        'store refs of all created devices 
        Protected Shared devlist As List(Of IODevice)
        'shared objects to lock a common bus during read/write, index=interfacelockid 
        Private Shared lockbus As Object()
        'will be set to main form
        Private Shared frm As Form

        Friend Shared devform As New DevicesForm()
        'to signal when DevicesForm has to update
        Friend Shared updated As Boolean = False



        'private instance variables:

        'task queue private variables:
        'to trigger wakeup of async thread
        Protected queue_event As EventWaitHandle = New AutoResetEvent(False)
        Private asyncthread As Thread
        'lock queue during operations on tasks
        Private ReadOnly lockqueue As New Object()

        Protected tasks As Queue(Of IOTask)
        'other private
        'set to to true when disposing to terminate properly (so that timer calls cannot fill queue)
        Private disposing As Boolean
        'set to true when DisposeDevice() called, so it can be called in main class finalizer
        Private devicedisposed As Boolean = False

        'index in lockbus: should be set by derived classes to distinguish between interfaces that can be used concurently
        Protected interfacelockid As Integer

        Protected interfacename As String = ""
        'last error message for errors not related to query (queue full, blocking call in progress) to display in devices form
        Protected Friend liberrmsg As String
        'to calculate delay
        Private lastoptime As DateTime

        'to lock device during cmd-resp sequence
        Private ReadOnly lockdev As New Object()

        'per device error message form
        Private msgfrm As IODeviceForms.IOmsgForm


        'currently existing async task or nothing if finished (volatile)
        Private currentasynctask As IOTask
        'currently existing blocking task or nothing if finished 
        Private currentblockingtask As IOTask
        ' task that currently has lock on device (may switch between async/blocking during retry loops)
        Protected currentactivetask As IOTask

        'currentactivequery: to use in implementation eg if need to know the command to format the response accordingly
        Protected ReadOnly Property currentactivequery() As IOQuery
            Get
                Dim ca As IOTask = currentactivetask 'make a copy of volatile objects
                If ca Is Nothing Then
                    Return Nothing
                Else
                    Return ca.query
                End If
            End Get
        End Property

        'event used to wait (sleep) in async routines so that notify can break waiting
        Private notify_event As EventWaitHandle = New AutoResetEvent(False)
        'equivalent boolean flag used in waiting loops in blocking calls
        Private notify_flag As Boolean

        '   **********  public methods

        'static constructor
        Shared Sub New()
            devlist = New List(Of IODevice)()

            Dim i As Integer = 0
            lockbus = New Object(100) {}
            For i = 0 To lockbus.Length - 1
                lockbus(i) = New Object()
            Next
            frm = Application.OpenForms(0)
            'set to Main Form 
            ' (null ref ok if callbacks can be executed on a not-GUI thread, without invoke, and if messages are disabled)

            If frm Is Nothing OrElse (frm.IsDisposed Or frm.Disposing) Then
                Throw New Exception("error creating IODevice class: main form null or disposed")
            End If

            'other init: send ifc etc:  in derived classes

            '
            updated = False


            If showdevicesonstartup Then ShowDevices()
        End Sub


        '------------------- IODevice public shared methods


        'Public Shared Sub ShowDevices()
        'If devform Is Nothing OrElse devform.IsDisposed OrElse devform.Disposing Then
        '       devform = New DevicesForm()
        'End If
        ''if disposing
        'Try
        '       devform.Show()
        '      devform.WindowState = FormWindowState.Normal
        '     devform.BringToFront()
        'Catch
        'End Try

        'End Sub


        'Public Shared Sub DisposeAll()

        '   devform.Close()
        ''faster if signal abort for all before shutting down one by one
        'For Each device As IODevice In devlist
        '       device.AbortAllTasks()
        'Next


        'While devlist.Count > 0

        '       devlist(0).Dispose()
        '      Application.DoEvents()
        'End While

        'End Sub








        ' Position the IODevices pop-up immediately to the right of the main app form.
        ' There is a big hack in here because on first run the pop-up was positioned fine, but on subsequent pop-ups it set itself x2 in both x and y direction.
        ' It seems to pop-up appeared twice, the main app under btncreate sub effectively calls the pop-up twice
        ' Declare the Windows API function to get window rectangle
        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function

        ' RECT structure to store window's position and size
        Private Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        ' Store the instance of the form
        'Friend Shared devform As DevicesForm

        ' Static counter to track the number of times ShowDevices has been called
        Private Shared showCount As Integer = 0

        ' Shared method to show the devices form
        Public Shared Sub ShowDevices()

            ' Increment the call counter
            showCount += 1

            ' Instantiate the DevicesForm if it doesn't exist or has been disposed
            If devform Is Nothing OrElse devform.IsDisposed OrElse devform.Disposing Then
                updated = False
                devform = New DevicesForm()
            End If

            ' Get the current process (main application)
            Dim mainApp As Process = Process.GetCurrentProcess()

            ' Get the main window handle
            Dim mainWindowHandle As IntPtr = mainApp.MainWindowHandle

            ' Get the position of the main window
            Dim mainWindowRect As New RECT()

            If GetWindowRect(mainWindowHandle, mainWindowRect) Then

                ' Base offsets
                Dim baseXOffset As Integer = 530
                Dim baseYOffset As Integer = 0

                ' Adjust offsets based on how many times the form has been shown.....this is a bit of a hack due to the pop-up effectively being called twice by the main app.
                Dim xOffset As Integer = 0
                Dim yOffset As Integer = 0
                If showCount < 2 Then   ' first two times the pop-up appears the offset is zero so it positions correctly
                    xOffset = 0
                    yOffset = 0
                    Exit Sub         ' Update: this added to abort, it seems that the pop-up appears twice when FIRST run after starting the app and this abort fixes that. Need to get to bottom of why this double pop-up happens on first run of WinGPIB
                End If
                If showCount >= 2 Then
                    xOffset = baseXOffset * 2   ' more than two times the position needs adjusted
                    yOffset = baseYOffset * 2
                End If

                Dim x As Integer = mainWindowRect.Left + xOffset
                Dim y As Integer = mainWindowRect.Top + yOffset

                ' only set position of pop-up if Dock mode enabled
                If My.Settings.dataDock = True Then
                    ' Set the location of the form
                    devform.StartPosition = FormStartPosition.Manual
                    devform.Location = New Point(x, y)
                End If

                ' Ensure form is shown and brought to the front
                devform.Show()
                devform.BringToFront()

            End If


            Try
                devform.WindowState = FormWindowState.Normal
            Catch ex As Exception
                ' Handle any exceptions here if needed
                Debug.WriteLine($"Error: {ex.Message}")
            End Try
        End Sub


        ' Method to dispose all resources
        Public Shared Sub DisposeAll()
            ' Check if the form is not null and is still open before closing
            If devform IsNot Nothing AndAlso Not devform.IsDisposed AndAlso devform.Visible Then
                devform.Close()
            End If

            ' Faster if signal abort for all before shutting down one by one
            For Each device As IODevice In devlist
                device.AbortAllTasks()
            Next

            ' Dispose of all devices
            While devlist.Count > 0
                devlist(0).Dispose()
                'devlist.RemoveAt(0) ' Ensure you remove the disposed item from the list
                Application.DoEvents() ' Allow UI thread to process other events
            End While
        End Sub








        'used by DevicesForm
        Public Shared Function GetDeviceList(ByVal details As Boolean) As String()
            Dim sl As String() = Nothing
            Dim n = devlist.Count
            If n = 0 Then
                Return Nothing
            End If
            Dim i As Short = 0
            Dim s As String = Nothing
            Dim sc As String = ""
            Dim st As String = ""
            Dim si As String = ""
            Dim scmd As String = Nothing


            Dim cbt As IOTask = Nothing
            Dim cat As IOTask = Nothing
            Dim ct As IOTask = Nothing

            For i = 0 To n - 1
                If devlist(i) Is Nothing Then
                    Continue For
                End If
                si = "(" & devlist(i).interfacename & ")"
                s = devlist(i).devname & "@"

                st = ""

                cbt = devlist(i).currentblockingtask

                'make a copy of volatile objects!!!
                Interlocked.Exchange(cat, devlist(i).currentasynctask)
                Interlocked.Exchange(ct, devlist(i).currentactivetask)



                sc = Convert.ToString(devlist(i).devaddr)

                If ct IsNot Nothing Then
                    If ct.blocking Then

                        st += ", blocking: "
                    Else
                        st += ", async: "
                    End If
                    st += "'" & ct.query.cmd & "'  "

                    If devlist(i).PendingTasks() > 0 Then
                        st += ",  pending:" & devlist(i).PendingTasks().ToString()
                    End If
                    If Not String.IsNullOrEmpty(ct.query.errmsg) Then
                        st += ",  error:" & ct.query.errmsg

                        If (cbt IsNot Nothing AndAlso cbt.retry) Or (cat IsNot Nothing AndAlso cat.retry) Then
                            st += " (retrying...)"

                        End If
                    End If
                End If
                If Not String.IsNullOrEmpty(devlist(i).liberrmsg) Then
                    st += ", " & devlist(i).liberrmsg
                End If

                If sl Is Nothing Then
                    sl = New String(0) {}
                Else
                    Array.Resize(sl, sl.Length + 1)
                End If

                sl(sl.Length - 1) = si & "  " & s & sc & st

                'queue
                If details Then


                    'make sure the task list does not change during loop
                    SyncLock devlist(i).lockqueue

                        For Each task As IOTask In devlist(i).tasks
                            If (task IsNot Nothing) Then
                                scmd = "    " & "'" & task.query.cmd & "'"
                                Array.Resize(sl, sl.Length + 1)
                                sl(sl.Length - 1) = scmd
                            End If
                        Next

                    End SyncLock

                End If
            Next
            Return sl
        End Function

        'find device in list using name
        Public Shared Function DeviceByName(ByVal name As String) As IODevice

            For Each d As IODevice In devlist
                If d.devname = name Then
                    Return d
                End If
            Next

            Return Nothing
        End Function
        ' helper function to interpret gpib or visa type address  (simple version, might be rewritten using regex class)
        ' address may be just a number eg "9", then board will be set to 0
        'or "0:9",  "GPIB0::9", "GPIB0:9", "GPIB0::9::INSTR" etc 
        'will 

        Public Shared Sub ParseGpibAddr(ByVal address As String, ByRef board As Integer, ByRef gpibaddr As Byte)

            'interpret address




            Dim sarr As String() = Nothing


            sarr = address.ToUpper().Split(":".ToCharArray())


            Try
                Select Case sarr.Length
                    Case 0

                        Throw New Exception("invalid address format: " & address)
                    Case 1
                        board = 0
                        gpibaddr = Byte.Parse(sarr(0))

                        Exit Select
                    Case Else
                        If sarr(0).Contains("GPIB") Then

                            Dim bs As String = sarr(0).Substring(4)

                            If bs.Length = 0 Then
                                board = 0
                            Else
                                board = Integer.Parse(sarr(0).Substring(4))

                            End If
                        Else

                            board = Integer.Parse(sarr(0))
                        End If

                        Dim idx As Integer = 1

                        While idx < sarr.Length AndAlso String.IsNullOrEmpty(sarr(idx))
                            idx += 1
                        End While
                        If idx = sarr.Length Then
                            Throw New Exception("invalid address format: " & address)
                        End If
                        gpibaddr = Byte.Parse(sarr(idx))

                        Exit Select

                End Select
            Catch
                Throw New Exception("invalid address format: " & address)
            End Try
            If gpibaddr = 0 Then
                Throw New Exception("invalid address format: " & address)
            End If

        End Sub



        '************class IODevice public instance methods *******************************
        '*******************************************


        ' constructor

        Public Sub New(ByVal name As String, ByVal addr As String)
            If frm Is Nothing OrElse (frm.IsDisposed Or frm.Disposing) Then
                Throw New Exception("error creating IODevice " & name & ": main form null or disposed")
            End If

            devname = name
            devaddr = addr
            'name and addr used only to display by devicesform and in error messages

            'set  default options
            delayop = 1
            delayread = 20
            delayrereadontimeout = 80
            delayretry = 1000
            readtimeout = 5000
            'cumulative timeout, independent of interface settings
            showmessages = True

            catchinterfaceexceptions = True
            'may be useful to set to false during debugging interface routines
            catchcallbackexceptions = True
            callbackonretry = True

            maxtasks = 50
            checkEOI = True
            eventsallowed = True
            enablepoll = True
            stripcrlf = True

            'create task queue: 
            tasks = New Queue(Of IOTask)()

            'start async thread:
            currentasynctask = Nothing
            currentblockingtask = Nothing
            currentactivetask = Nothing
            asyncthread = New Thread(AddressOf AsyncThreadProc)
            asyncthread.IsBackground = True
            asyncthread.Start()

            lastoptime = DateTime.Now

            disposing = False
            'add device to global list:
            '  moved to child class constructors (will not be called if exception occurs)



            updated = False
        End Sub

        ' interrupts waiting for next reading or poll trial
        ' may be called e.g. by "Notify" callback (if defined by the implementation) when data ready
        'can be called from any thread
        Protected Sub WakeUp()

            'we don't know from which thread it will be called therefore signal is sent to both blocking and async tasks (will be rearmed ion next command anyway)
            notify_flag = True
            'used in waitdelay 
            Try 'may cause error on disposing
                notify_event.Set()
                'for async calls
            Catch
            End Try
        End Sub


        Public Sub Dispose() Implements IDisposable.Dispose
            If disposing Then
                Return
            End If

            AbortAllTasks()

            disposing = True
            'prevent new tasks to be appended, signal async thread to exit

            If showmessages Then
                Try
                    msg_end_TS()
                Catch
                End Try
            End If



            queue_event.[Set]()
            'set event to wake up async thread 
            EnqueueTask(Nothing)
            'signal async thread to exit

            'wait 3s for the async  thread to finish gracefully
            If Not asyncthread.Join(3000) Then
                asyncthread.Abort()
            End If


            queue_event.Close()
            notify_event.Close()

            devlist.Remove(Me)
            If msgfrm IsNot Nothing AndAlso Not msgfrm.IsDisposed Then
                msgfrm.Close()
            End If



            ' Release external unmanaged resources
            If Not devicedisposed Then
                Try
                    DisposeDevice()
                    devicedisposed = True
                    'will  be tested in finalizer
                    GC.SuppressFinalize(Me)
                Catch
                End Try
            End If


        End Sub
        'finalizer: if not disposed properly make sure the unmanaged 
        'resources are released (eg notify handler uninstalled) before the object is garbage-collected!

        Protected Overrides Sub Finalize()
            Try
                If Not devicedisposed Then
                    DisposeDevice()
                End If
            Finally
                devicedisposed = True
                MyBase.Finalize()
            End Try
        End Sub



        Public Function PendingTasks() As Integer

            Dim retval As Integer = 0
            SyncLock lockqueue
                retval = tasks.Count
            End SyncLock

            Return retval

        End Function

        Public Function PendingTasks(ByVal t As DateTime) As Integer
            'version counting only tasks with queries called before or at t

            Dim p As Short = 0

            SyncLock lockqueue

                For Each task As IOTask In tasks
                    If (task IsNot Nothing) Then
                        If task.query.timecall <= t Then
                            p += 1
                        End If
                    End If
                Next
            End SyncLock

            Return p



        End Function
        Public Function PendingTasks(ByVal cmd As String) As Integer
            'version counting only tasks with certain commands
            'case insensitive

            Dim p As Short = 0

            'make sure async thread does not remove a task during the loop
            SyncLock lockqueue

                For Each task As IOTask In tasks
                    If (task IsNot Nothing) Then
                        If task.query.cmd.ToUpper() = cmd.ToUpper() Then
                            p += 1
                        End If
                    End If
                Next
            End SyncLock

            Return p



        End Function

        Public Function PendingTasks(ByVal tag As Integer) As Integer
            'version counting only tasks with certain tag values


            Dim p As Short = 0

            'make sure async thread does not remove a task during the loop
            SyncLock lockqueue

                For Each task As IOTask In tasks
                    If (task IsNot Nothing) Then
                        If task.query.tag = tag Then
                            p += 1
                        End If
                    End If
                Next
            End SyncLock

            Return p



        End Function



        Public Sub AbortAllTasks()

            SyncLock lockqueue

                For Each task As IOTask In tasks
                    If (task IsNot Nothing) Then
                        task.abort = True

                    End If
                Next
                Try
                    If currentasynctask IsNot Nothing Then
                        currentasynctask.abort = True
                    End If
                    If currentblockingtask IsNot Nothing Then
                        currentblockingtask.abort = True
                    End If
                    If currentactivetask IsNot Nothing Then
                        currentactivetask.abort = True

                    End If
                Catch
                End Try
            End SyncLock

        End Sub

        Public Sub WaitAsync()
            WaitAsync(DateTime.Now.AddTicks(2))
            'make sure for last command 
        End Sub

        'wait until async queries queued before the call are done (usually set ts to "Now")
        Public Sub WaitAsync(ByVal ts As DateTime)


            Dim t As IOTask = Nothing
            Dim bp As Boolean = False
            Dim bt As Boolean = False
            Dim p As Integer = 0
            Dim pt As Integer = 0


            Do
                t = currentasynctask
                p = PendingTasks(ts)
                pt = PendingTasks()
                bp = (p = 0)
                bt = (t Is Nothing OrElse t.query.timecall > ts)


                Application.DoEvents()
                Thread.Sleep(1)
            Loop While Not ((bp And bt) Or disposing)



        End Sub

        Public Function IsBlocking() As Boolean

            Return currentblockingtask IsNot Nothing

        End Function





        'query 1st version: callback

        Public Function QueryAsync(ByVal cmd As String, ByVal callback As IOCallback, ByVal retry As Boolean) As Integer

            Return makeQueryAsync(cmd, Nothing, callback, True, retry, 0)

        End Function

        Public Function QueryAsync(ByVal cmd As String, ByVal callback As IOCallback, ByVal retry As Boolean, ByVal cbwait As Boolean, ByVal tag As Integer) As Integer

            Return makeQueryAsync(cmd, Nothing, callback, cbwait, retry, tag)

        End Function
        '2nd version : textbox
        Public Function QueryAsync(ByVal cmd As String, ByVal text As TextBox, ByVal retry As Boolean) As Integer

            Return makeQueryAsync(cmd, text, Nothing, False, retry, 0)
            ''

        End Function
        Public Function QueryAsync(ByVal cmd As String, ByVal text As TextBox, ByVal retry As Boolean, ByVal tag As Integer) As Integer

            Return makeQueryAsync(cmd, text, Nothing, False, retry, tag)

        End Function

        'send command and wait for response
        Protected Function makeQueryAsync(ByVal cmd As String, ByVal text As TextBox, ByVal callback As IOCallback, ByVal cbwait As Boolean, ByVal retry As Boolean, ByVal tag As Integer) As Integer

            'return 0 if thread run, -1 if too many tasks, -2 if other error

            If disposing Then
                Return -2
            End If



            If PendingTasks() >= maxtasks Then
                liberrmsg = "async queue full"

                Return -1
            End If


            'make query
            Dim query As New IOQuery(2, cmd, tag)
            'type 2wait for  response

            Dim task As New IOTask(Me, query)

            query.timecall = DateTime.Now



            task.txt = text
            task.callback = callback
            task.cbwait = cbwait
            task.retry = retry
            task.events = False



            'append 
            EnqueueTask(task)
            Return 0

        End Function



        'send command, no response
        Protected Function makeSendAsync(ByVal cmd As String, ByVal callback As IOCallback, ByVal cbwait As Boolean, ByVal retry As Boolean, ByVal tag As Integer) As Integer


            'return 0 if ok, -1 if too many tasks, -2 if disabled 

            If disposing Then
                Return -2
            End If



            If PendingTasks() >= maxtasks Then
                liberrmsg = "async queue full"

                Return -1
            End If

            Dim query As New IOQuery(1, cmd, tag)
            'type 1 no resp

            Dim task As New IOTask(Me, query)

            query.timecall = DateTime.Now

            task.callback = callback
            task.cbwait = cbwait
            task.retry = retry
            task.events = False

            'append 
            EnqueueTask(task)

            Return 0

        End Function
        'send command, no response,
        Public Function SendAsync(ByVal cmd As String, ByVal callback As IOCallback, ByVal retry As Boolean, ByVal cbwait As Boolean, ByVal tag As Integer) As Integer

            Return makeSendAsync(cmd, callback, cbwait, retry, tag)

        End Function
        'default behavior : no callback
        Public Function SendAsync(ByVal cmd As String, ByVal retry As Boolean) As Integer

            Return makeSendAsync(cmd, Nothing, False, retry, 0)

        End Function



        'blocking versions  ********************

        'send command and wait for response


        Public Function QueryBlocking(ByVal cmd As String, ByRef q As IOQuery, ByVal retry As Boolean) As Integer


            'return -1 if blocking call in progress (may happen if events allowed), -2 disposing
            'otherwise return status as in IOQuery (0 if ok)


            q = New IOQuery(2, cmd, 0)



            If disposing Then
                q.status = -2
                Return q.status
            End If


            If currentblockingtask IsNot Nothing Then
                q.status = -1
                q.errmsg = "blocking call in progress"
                'necessary because wait loop until bus unlocked (otherwise will wait forever since is on the same thread)
                Return q.status
            End If



            Dim t As New IOTask(Me, q)

            t.query.timecall = DateTime.Now

            t.cbwait = False
            'no meaning here
            t.retry = retry
            t.events = eventsallowed

            t.delayread = delayread


            'run directly from calling thread but lock device so that Async functions can be used concurrently
            t.blocking = True
            Interlocked.Exchange(currentblockingtask, t)
            updated = False

            t.TaskProc()
            t.blocking = False



            Interlocked.Exchange(currentblockingtask, Nothing)
            updated = False

            Return q.status

        End Function




        Public Function QueryBlocking(ByVal cmd As String, ByRef resp As String, ByVal retry As Boolean) As Integer


            'return -1 if blocking call in progress (may happen if events allowed), -2 disabled
            'otherwise return status as in IOQuery (0 if ok)

            resp = Nothing

            If disposing Then
                Return -2
            End If

            If currentblockingtask IsNot Nothing Then
                liberrmsg = "blocking call in progress"
                Return -1
            End If

            Dim q As IOQuery = Nothing
            'or resp = ""  'to avoid null ref exceptions for lazy programmers?
            Dim st As Integer = QueryBlocking(cmd, q, retry)

            If q IsNot Nothing Then
                resp = q.ResponseAsString
            End If

            Return st

        End Function

        '********************************
        Public Function QueryBlocking(ByVal cmd As String, ByRef resparr As Byte(), ByVal retry As Boolean) As Integer

            'return -1 if blocking call in progress (may happen if events allowed), -2 disabled
            'otherwise return status as in IOQuery (0 if ok)

            resparr = Nothing

            If disposing Then
                Return -2
            End If

            If currentblockingtask IsNot Nothing Then
                liberrmsg = "blocking call in progress"
                Return -1
            End If

            Dim q As IOQuery = Nothing

            Dim st As Integer = QueryBlocking(cmd, q, retry)

            If q IsNot Nothing Then
                resparr = q.ResponseAsByteArray
            End If

            Return st


        End Function


        'send command, no response

        Public Function SendBlocking(ByVal cmd As String, ByVal retry As Boolean) As Integer


            'return -1 if blocking call in progress (may happen if events allowed)
            'return -2 if disposing
            'otherwise return status as in IOQuery 

            If disposing Then
                Return -2
            End If


            If currentblockingtask IsNot Nothing Then
                liberrmsg = "blocking call in progress"
                Return -1
            End If

            Dim query As New IOQuery(1, cmd, 0)



            Dim t As New IOTask(Me, query)

            t.query.timecall = DateTime.Now


            t.cbwait = False
            'no meaning here
            t.retry = retry
            t.events = eventsallowed

            t.delayread = delayread
            'in blocking commands may help not to block bus and GUI events when polling is not available

            'run directly from calling thread but lock device so that Async functions can be used concurrently
            t.blocking = True
            Interlocked.Exchange(currentblockingtask, t)
            updated = False

            t.TaskProc()
            t.blocking = False

            query = t.query


            Interlocked.Exchange(currentblockingtask, Nothing)
            updated = False

            Return t.query.status

        End Function


        Public Overridable Function ByteArrayToString(ByVal arr As Byte()) As String

            Dim s As String = Nothing

            Try
                Dim len = arr.Length
                'may cause exception if no array
                'remove terminating LF, CR if any
                If stripcrlf Then
                    If len > 0 AndAlso arr(len - 1) = 10 Then
                        len = len - 1
                    End If
                    'ignore lf at end
                    If len > 0 AndAlso arr(len - 1) = 13 Then
                        len = len - 1
                        'ignore cr
                    End If
                End If

                s = System.Text.Encoding.UTF8.GetString(arr, 0, len)
            Catch

                s = Nothing
            End Try

            Return s
        End Function




        '*****************  interface abstract methods that have to be defined
        '
        ' all functions  should return 0 if ok, 1 if timeout,  other value if other error
        ' the functions should catch all interface exceptions (otherwise the "catchinterfaceexceptions" flag should be set)
        ' if there is an error (returned value different from 0 or 1)  the information should be returned in errcode and errmsg
        ' errcode and errmsg are just for display and are not interpreted by the class

        Protected MustOverride Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer
        Protected MustOverride Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
        Protected MustOverride Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
        'poll for status, return status byte and MAV bit 
        'statusbyte is for user info (displayed in message window) and also to allows to easily override this method reinterpreting the status byte for devices not compliant with 488.2 (eg.Lakeshore),
        'the function should interpret it and set the mav flag which will be used by the class to decide when to start reading data
        '(for interfaces that don't implement this feature (serial) the function will set it to true)
        ' the function will be called repeatedly until it sets mav to true (or until cumulative timeout period "readtimeout" elapses)
        Protected MustOverride Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer
        'if the function returns 1 (timeout), the function will be called repeatedly (until cumulative timeout period "readtimeout" elapses)
        'EOI is the "message complete" status flag. If set to false the function will be called repeatedly and data chunks assembled until it is set to true (or timeout period elapses)
        '(usually EOI will be set to false when buffer gets full)
        Protected MustOverride Sub DisposeDevice() 'release unmanaged resources, if any


        'related overridable methods : 
        '      Protected Overridable Function ByteArrayToString(ByVal arr() As Byte) As String




        'must be called only by child class constructors
        Public Sub AddToList()


            devlist.Add(Me)

            'force updating DevicesForm
            updated = False

        End Sub

        '*******************************  IODevice private methods


        'queue handling with lock:

        Private Sub EnqueueTask(ByVal task As IOTask)
            SyncLock lockqueue
                If tasks.Count <= maxtasks Or task Is Nothing Then
                    tasks.Enqueue(task)
                End If
            End SyncLock
            queue_event.[Set]()
            updated = False
            'set event to wake up async thread

        End Sub
        'return true if there was a task (could be nothing)
        Private Function DequeueTask() As Boolean
            SyncLock lockqueue
                If tasks.Count > 0 Then
                    currentasynctask = tasks.Dequeue()
                    updated = False
                    Return True
                Else

                    Return False

                End If
            End SyncLock

        End Function


        'asyncthread
        Private Sub AsyncThreadProc()

            While Not disposing

                If DequeueTask() Then
                    If currentasynctask Is Nothing Then
                        'end async thread if null task has been queued
                        'do task
                        Return
                    Else
                        If Not disposing Then
                            currentasynctask.TaskProc()
                            'give a chance to blocking calls 
                            Thread.Sleep(1)
                        End If

                        'no more tasks
                    End If
                Else
                    currentasynctask = Nothing
                    'signal when current task finished
                    If Not disposing Then
                        queue_event.WaitOne(200)
                        ' No more tasks - wait for a signal (timeout 200 to ensure exit when disposing) 

                    End If
                End If
            End While
        End Sub

        Private Shared Function Ismainthread() As Boolean

            If frm Is Nothing Then
                'form not defined eg console app ?
                Return True
            Else
                If Not (frm.IsDisposed Or frm.Disposing) Then
                    Return Not frm.InvokeRequired
                Else
                    'Debug.Print("error in IOTask.Ismainthread: form disposed")
                    'safer
                    Return False
                End If
            End If


        End Function




        'must be invoked from gui thread
        'make/show error message form when closing device
        Private Sub msg_end(ByVal msg As String)


            If msgfrm Is Nothing OrElse msgfrm.IsDisposed Then
                msgfrm = New IODeviceForms.IOmsgForm()
            End If
            msgfrm.shutdownmsg(msg)
        End Sub

        'in case dispose called by another thread?
        Private Sub msg_end_TS()


            Try

                If frm.InvokeRequired Then

                    frm.Invoke(New Action(Of String)(AddressOf msg_end), New Object() {devname})
                Else
                    msg_end(devname)
                End If
            Catch
            End Try

        End Sub


        '************************inner class GpibTask: one instance for each query created (tasks will be queued)

        Protected Friend Class IOTask
            'copy of parent to access its members
            Protected Friend parent As IODevice


            Protected Friend query As IOQuery
            'fields used to return result
            Protected Friend txt As TextBox

            Protected Friend callback As IOCallback
            'delay between cmd and read : to avoid blocking gpib bus by slow devices
            Protected Friend delayread As Integer

            'set true when called on main thread from blocking functions
            Protected Friend blocking As Boolean
            'thread wait (not locked) until callback resumes
            Protected Friend cbwait As Boolean
            'retry on timeout
            Protected Friend retry As Boolean
            'abort signal
            Protected Friend abort As Boolean
            'if doevents allowed on retry loop: set false in async and "eventsallowed" in blocking
            Protected Friend events As Boolean
            Private CrLf As String = vbCr & vbLf
            'System.Text.Encoding.UTF8.GetString(new byte[2] { 13, 10 });

            'constructor

            Public Sub New(ByVal parentdev As IODevice, ByVal q As IOQuery)

                parent = parentdev
                query = q
                query.task = Me



                delayread = parent.delayread
                txt = Nothing

                callback = Nothing
                abort = False

                blocking = False
            End Sub
            'private methods

            'wait when in GUI thread
            Private Sub waitevents(ByVal delms As Integer, ByVal allowbreak As Boolean)

                Const sleepms As Short = 3

                Dim t1 As DateTime = Nothing
                Dim msecs As Double = 0

                If delms < 0 Then
                    Return
                End If

                t1 = DateTime.Now


                Do
                    msecs = DateTime.Now.Subtract(t1).TotalMilliseconds
                    'total!!! 

                    If events Then
                        System.Windows.Forms.Application.DoEvents()
                    End If
                    'yield shortly 
                    Thread.Sleep(sleepms)
                Loop While Not (msecs >= delms OrElse (allowbreak And parent.notify_flag))

                parent.notify_flag = False  'auto rearm

            End Sub

            'wait delay (in ms) counted from "fromtime" (usually set from=Now)

            '<HandleProcessCorruptedStateExceptions> _
            Private Sub waitdelay(ByVal fromtime As DateTime, ByVal delay As Integer, ByVal allowbreak As Boolean)


                Dim m As Integer = 0

                m = delay - DateTime.Now.Subtract(fromtime).Milliseconds

                If m > 0 Then
                    If Not Ismainthread() Then
                        If allowbreak Then
                            Try
                                parent.notify_event.WaitOne(m)  ' to enable wake up by Wakeup()
                            Catch
                            End Try

                        Else

                            Thread.Sleep(m)
                        End If
                    Else
                        waitevents(m, allowbreak)
                    End If
                End If





            End Sub
            Private Sub rearmforwakeup()  'used when sending new command

                If Ismainthread() Then
                    parent.notify_flag = False
                Else
                    parent.notify_event.Reset()
                End If

            End Sub







            'poll for status, wait (readtimeout) until MAV bit set 
            Private Sub waitMAV()

                'polling frequency defined by delayrereadontimeout

                Dim rereadflag As Boolean = False

                Dim pollcount As Integer = 0

                Dim startpoll As DateTime = Nothing
                Dim mav As Boolean = False
                Dim statusbyte As Byte = 0

                Dim result As Integer = 0
                Dim errmsg As String = ""

                startpoll = DateTime.Now
                '-----------reread:  loop here to repeat poll 
                Do
                    rereadflag = False

                    Dim buslocked As Boolean = False


                    'set <0 if no bus lock needed (eg. serial)
                    If parent.interfacelockid >= 0 Then
                        buslocked = False
                        'if already locked by another device
                        While Not buslocked And Not abort And Not parent.disposing
                            '******************************** critical section for the bus: locked at global class level
                            ' use Monitor.TryEnter  instead of synclock to avoid deadlocks or freezing GUI (when a blocking query waits for lock, if it was freezing GUI  then "Invoke" would be blocked too) 
                            buslocked = Monitor.TryEnter(IODevice.lockbus(parent.interfacelockid), 5)
                            If Not buslocked Then
                                Thread.Sleep(1)
                                'true only when the call belongs to a task performing a "blocking" routine from gui thread
                                If blocking Then
                                    'must give a chance to process events here, to avoid freezing GUI, when blocking call waits until unlocked by asyncthread
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                            End If
                        End While
                    End If


                    'may abort/dispose on waiting for lock 
                    If buslocked Or parent.interfacelockid < 0 Then

                        Try
                            result = parent.PollMAV(mav, statusbyte, query.errcode, errmsg)
                        Catch ex As Exception
                            query.errmsg = "exception in 'PollMav': " & CrLf & ex.Message
                            If ex.InnerException IsNot Nothing Then
                                query.errmsg += CrLf & ex.InnerException.Message
                            End If
                            result = -1
                            query.errcode = -1
                            query.status = 6 + 16
                            If Not parent.catchinterfaceexceptions Then
                                Throw
                                're-throw the exception with same stack
                            End If
                        Finally
                            If buslocked Then
                                Monitor.[Exit](IODevice.lockbus(parent.interfacelockid))
                            End If
                        End Try


                        pollcount += 1

                        'status=1 tmo on send,  =3 tmo on rcv, =4 other err on send, =6 other err on rcv
                        If mav And (result = 0) Then

                            query.status = 0
                        Else

                            Dim d As Double = DateTime.Now.Subtract(startpoll).TotalMilliseconds

                            If d < parent.readtimeout Then
                                rereadflag = True
                                'retry, (dont set timeout status yet)
                                pollcount += 1
                            Else
                                rereadflag = False
                                'status=1 tmo on send; =3 tmo on rcv, =4 other err on send, =6 other err on rcv
                                query.status = 3 + 16
                                'definitive read timout : poll problem


                                Select Case result
                                    Case 0 : query.errmsg = "poll timeout: MAV not set, status byte=" & statusbyte
                                    Case 1 : query.errmsg = "poll timeout: cannot get status byte"
                                    Case Else : query.errmsg = "poll error : " & errmsg
                                        query.status = 6 + 16
                                End Select

                            End If


                        End If


                        If rereadflag Then
                            waitdelay(DateTime.Now, parent.delayrereadontimeout, True)
                        End If

                    End If
                Loop While Not (Not rereadflag Or abort Or parent.disposing)
                '(though no way to abort from msg window)




            End Sub



            Private Sub sendcmd()

                Dim result As Integer = 0

                Dim buslocked As Boolean = False


                'set <0 if no bus lock needed (eg. serial)
                If parent.interfacelockid >= 0 Then
                    buslocked = False
                    'if already locked by another device
                    While Not buslocked And Not abort And Not parent.disposing
                        '******************************** critical section for the bus: locked at global class level
                        ' use Monitor.TryEnter  instead of synclock to avoid deadlocks or freezing GUI (when a blocking query waits for lock, if it was freezing GUI  then "Invoke" would be blocked too) 
                        buslocked = Monitor.TryEnter(IODevice.lockbus(parent.interfacelockid), 5)
                        If Not buslocked Then
                            Thread.Sleep(1)
                            'true only when the call belongs to a task performing a "blocking" routine from gui thread
                            If blocking Then
                                'must give a chance to process events here, to avoid freezing GUI, when blocking call waits until unlocked by asyncthread
                                System.Windows.Forms.Application.DoEvents()
                            End If
                        End If
                    End While
                End If


                'may abort/dispose on waiting for lock  
                If buslocked Or parent.interfacelockid < 0 Then

                    Try

                        result = parent.Send(query.cmd, query.errcode, query.errmsg)
                    Catch ex As Exception
                        query.errmsg = "exception in 'Send': " & CrLf & ex.Message
                        If ex.InnerException IsNot Nothing Then
                            query.errmsg += CrLf & ex.InnerException.Message
                        End If
                        result = -1
                        query.errcode = -1
                        query.status = 4
                        If Not parent.catchinterfaceexceptions Then
                            Throw
                            're-throw the exception with same stack
                        End If
                    Finally
                        If buslocked Then
                            Monitor.[Exit](IODevice.lockbus(parent.interfacelockid))
                        End If
                    End Try


                    'so status=1 tmo on send;
                    ' =3 tmo on rcv, =4 other err on send, =6 other err on rcv
                    If result = 0 Then

                        query.status = 0
                    Else
                        'timeout 
                        If result = 1 Then
                            query.status = 1
                            'send tmo
                            ' Message "aborted" not clear...

                            query.errmsg = "write timeout"
                        Else
                            query.status = 4

                        End If
                    End If
                End If

            End Sub


            ''--------------------

            Private Sub getresponse()
                Dim rereadflag As Boolean = False
                Dim readcount As Integer = 0
                'for tests
                Dim startread As DateTime = Nothing

                Dim result As Integer = 0


                startread = DateTime.Now
                '-----------reread:  loop here to repeat read on GPIB timeout without setting timeout status (slow devices)
                Do
                    rereadflag = False
                    Dim EOI As Boolean = True
                    Dim arr As Byte() = Nothing

                    Dim buslocked As Boolean = False


                    'set <0 if no bus lock needed (eg. serial)
                    If parent.interfacelockid >= 0 Then
                        buslocked = False
                        'if already locked by another device
                        While Not buslocked And Not abort And Not parent.disposing
                            '******************************** critical section for the bus: locked at global class level
                            ' use Monitor.TryEnter  instead of synclock to avoid deadlocks or freezing GUI (when a blocking query waits for lock, if it was freezing GUI  then "Invoke" would be blocked too) 
                            buslocked = Monitor.TryEnter(IODevice.lockbus(parent.interfacelockid), 5)
                            If Not buslocked Then
                                Thread.Sleep(1)
                                'true only when the call belongs to a task performing a "blocking" routine from gui thread
                                If blocking Then
                                    'must give a chance to process events here, to avoid freezing GUI, when blocking call waits until unlocked by asyncthread
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                            End If
                        End While
                    End If


                    'may abort/dispose on waiting for lock 
                    If buslocked Or parent.interfacelockid < 0 Then


                        Try

                            result = parent.ReceiveByteArray(arr, EOI, query.errcode, query.errmsg)
                        Catch ex As Exception
                            query.errmsg = "exception in 'ReceiveByteArray': " & CrLf & ex.Message
                            If ex.InnerException IsNot Nothing Then
                                query.errmsg += CrLf & ex.InnerException.Message
                            End If
                            result = -1
                            query.errcode = -1
                            query.status = 6
                            If Not parent.catchinterfaceexceptions Then
                                Throw
                                're-throw the exception with same stack
                            End If
                        Finally
                            If buslocked Then
                                Monitor.[Exit](IODevice.lockbus(parent.interfacelockid))
                            End If
                        End Try


                        If (result = 0) And (arr IsNot Nothing) Then
                            copyorappend_array(arr, query.resparr)
                            'redefine timeout condition after successful reading

                            startread = DateTime.Now
                        End If

                        Dim part As Boolean = Not EOI And parent.checkEOI  'data partially read

                        'reread: append (works if EOI false (buffer too small) but usually causes errors if aborted during IO)
                        If (result = 0) And part Then
                            rereadflag = True
                        End If
                        'no delay if reread because no EOI
                        readcount += 1


                        'so status=1 tmo on send;
                        ' =3 tmo on rcv, =4 other err on send, =6 other err on rcv

                        If (result = 0) And Not part Then

                            query.status = 0
                        Else
                            'check timeout condition
                            If (result = 1) Or ((result = 0) And part) Then

                                'first check if reread:
                                'if cumulated time < readtmo then just repeat 
                                Dim d As Double = DateTime.Now.Subtract(startread).TotalMilliseconds

                                If d < parent.readtimeout Then
                                    rereadflag = True
                                    'dont set timeout status yet

                                    readcount += 1
                                Else
                                    'status=1 tmo on send; =3 tmo on rcv, =4 other err on send, =6 other err on rcv
                                    rereadflag = False
                                    query.status = 3
                                    'definitive read timout
                                    query.errmsg = "read timeout"
                                    'ex.Message
                                    If part Then
                                        query.errmsg += "  (EOI not set)"
                                    End If
                                    'other error
                                End If
                            Else


                                query.status = 6
                            End If
                        End If

                        If rereadflag Then
                            waitdelay(DateTime.Now, parent.delayrereadontimeout, True)
                        End If

                    End If
                Loop While Not (Not rereadflag Or abort Or parent.disposing)
                '(though no way to abort from msg window)


            End Sub


            Private Sub cleardevice()
                Dim result As Integer = 0

                Dim errmsg As String = ""
                Dim lasterrmsg As String = query.errmsg
                Dim count = 0

                '--------- retry clearing device
                Do
                    count += 1
                    Dim buslocked As Boolean = False


                    'set <0 if no bus lock needed (eg. serial)
                    If parent.interfacelockid >= 0 Then
                        buslocked = False
                        'if already locked by another device
                        While Not buslocked And Not abort And Not parent.disposing
                            '******************************** critical section for the bus: locked at global class level
                            ' use Monitor.TryEnter  instead of synclock to avoid deadlocks or freezing GUI (when a blocking query waits for lock, if it was freezing GUI  then "Invoke" would be blocked too) 
                            buslocked = Monitor.TryEnter(IODevice.lockbus(parent.interfacelockid), 5)
                            If Not buslocked Then
                                Thread.Sleep(1)
                                'true only when the call belongs to a task performing a "blocking" routine from gui thread
                                If blocking Then
                                    'must give a chance to process events here, to avoid freezing GUI, when blocking call waits until unlocked by asyncthread
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                            End If
                        End While
                    End If


                    'may abort/dispose on waiting for lock 
                    If buslocked Or parent.interfacelockid < 0 Then
                        Try
                            result = parent.ClearDevice(query.errcode, errmsg)
                        Catch ex As Exception
                            query.errmsg = "exception in 'ClearDevice': " & CrLf & ex.Message
                            If ex.InnerException IsNot Nothing Then
                                query.errmsg += CrLf & ex.InnerException.Message
                            End If
                            result = -1
                            query.errcode = -1
                            If Not parent.catchinterfaceexceptions Then
                                Throw
                                're-throw the exception with same stack
                            End If
                        Finally
                            If buslocked Then
                                Monitor.[Exit](IODevice.lockbus(parent.interfacelockid))
                            End If
                        End Try


                        'query.status = 0  'don't reinitialize status if cleared (needed to retry!)
                        If result <> 0 Then
                            query.status = 4
                            If count = 1 Then query.errmsg += CrLf & "cannot clear device:" + errmsg
                            runsub(AddressOf msg_on, query, False)

                            waitdelay(DateTime.Now, parent.delayretry, True)
                        Else
                            query.errmsg = lasterrmsg & CrLf & "device cleared"
                        End If
                        'retry clearing device
                    End If
                Loop While Not (result = 0 Or Not retry Or abort Or parent.disposing)

            End Sub

            'main IO task


            Protected Friend Sub TaskProc()

                query.resparr = Nothing
                query.status = 0
                query.errmsg = ""
                query.timestart = DateTime.MinValue

                parent.liberrmsg = ""
                'reset when task starts


                Dim locked As Boolean = False
                'device lock




                If Not parent.disposing And Not abort Then
                    '-----------------   query retry loop
                    Do
                        ' try to lock device:
                        locked = False
                        'if already locked by async call
                        While Not locked And Not abort And Not parent.disposing
                            '******************************** critical section for the device: locked at device instance level
                            ' use Monitor.TryEnter  instead of synclock to avoid deadlocks or freezing GUI (when a blocking query waits for lock, if it was freezing GUI  then "Invoke" would be blocked too) 
                            locked = Monitor.TryEnter(parent.lockdev, 5)
                            If Not locked Then
                                Thread.Sleep(1)
                                'true only when the call belongs to a task performing a "blocking" routine from gui thread
                                If blocking Then
                                    'must give a chance to process events here, to avoid freezing GUI, when blocking call waits until unlocked by asyncthread
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                            End If
                        End While

                        'may abort/dispose on waiting for lock  
                        If locked Then

                            Try
                                parent.currentactivetask = Me
                                updated = False
                                'repeat in case another thread changes it during retry

                                If query.timestart = DateTime.MinValue Then
                                    query.timestart = DateTime.Now
                                End If
                                'set time start when device available and gpib op starts

                                'make sure there is minimum delay counted from end of last operation (for slow devices)
                                'wait when device locked so cannot be preempted by other commands 
                                waitdelay(parent.lastoptime, parent.delayop, False)

                                query.status = 0
                                query.errmsg = ""


                                '-----------send command
                                If Not String.IsNullOrEmpty(query.cmd) Then
                                    rearmforwakeup() 'rearming before sending a new command rather than before starting to wait to avoid race condition
                                    'if no command sent then only auto-rearming after a delay, so that driver callbacks can anticipate delays 
                                    sendcmd()
                                End If


                                '------------- if send succeeded then  reading (otherwise will retry)


                                If Not abort And Not parent.disposing And (query.status = 0) And (query.type = 2) Then

                                    waitdelay(DateTime.Now, delayread, True)

                                    If parent.enablepoll Then
                                        'serial poll periodically to check when data available
                                        waitMAV()
                                    End If

                                    'If polling ok 
                                    If (query.status = 0) And Not abort And Not parent.disposing Then
                                        getresponse()

                                    End If
                                End If
                                '  end of read


                                If query.status <> 0 And Not abort And Not parent.disposing Then
                                    waitdelay(DateTime.Now, parent.delayretry, False)
                                    'wait before retry (necessary in blocking cmd)
                                    cleardevice()
                                End If
                            Finally
                                parent.lastoptime = DateTime.Now
                                'mark timestamp of end of last op
                                Monitor.Exit(parent.lockdev)
                                'release lock on device
                                parent.currentactivetask = Nothing
                                updated = False
                            End Try
                        End If



                        'update before retry loop:
                        If abort Then
                            query.status += 8
                        End If
                        'will be visible only in lastquery

                        If query.status > 0 And Not abort And Not parent.disposing Then
                            runsub(AddressOf msg_on, query, False)
                            'call before retry
                            If retry And parent.callbackonretry Then

                                runsub(callback, query, cbwait)

                            End If
                        End If


                        If query.status = 0 Or abort Or parent.disposing Then
                            runsub(AddressOf msg_off, query, False)
                            'close msg if ok or aborted by user

                            If query.status = 0 Then
                                query.errmsg = ""
                            Else
                                query.errmsg = "task aborted by user"
                            End If

                            '       ------------  main retry loop 
                        End If

                    Loop While Not (query.status = 0 Or Not retry Or abort Or parent.disposing)
                End If

                query.timeend = DateTime.Now




                If query.status = 0 And query.type = 2 And txt IsNot Nothing And Not abort Then

                    settext_TS(txt, parent.ByteArrayToString(query.resparr))
                End If

                If Not blocking Then
                    Interlocked.Exchange(parent.lastasyncquery, query)
                End If


                If Not (retry And parent.callbackonretry And query.status > 0) And Not abort And Not parent.disposing Then
                    'will  set status to -1 on exception in callback
                    runsub(callback, query, cbwait)
                End If




            End Sub

            'append front to back
            Private Sub copyorappend_array(ByVal front As Byte(), ByRef back As Byte())
                If front Is Nothing Then
                    Return
                End If
                'just leave as is


                If back Is Nothing Then
                    'just copy array

                    back = front
                Else
                    Dim backOriginalLength = back.Length
                    Array.Resize(back, backOriginalLength + front.Length)

                    Array.Copy(front, 0, back, backOriginalLength, front.Length)
                End If

            End Sub




            'must be invoked from gui thread
            Private Sub msg_off(ByVal q As IOQuery)


                If parent.showmessages Then
                    If parent.msgfrm IsNot Nothing AndAlso Not parent.msgfrm.IsDisposed Then
                        parent.msgfrm.Close()
                    End If
                End If

            End Sub
            'param q to init msgform
            'must be invoked from gui thread

            'make/show error message form, 
            Private Sub msg_on(ByVal q As IOQuery)
                If query.status = 0 Then
                    Return
                End If


                If parent.showmessages Then
                    If parent.msgfrm Is Nothing OrElse parent.msgfrm.IsDisposed Then
                        parent.msgfrm = New IODeviceForms.IOmsgForm()
                    End If

                    parent.msgfrm.showit(q)
                End If


            End Sub


            'thread-safe version of  txt.Text = s 

            Private Shared Sub settext_TS(ByVal txt As TextBox, ByVal s As String)


                If txt IsNot Nothing AndAlso Not (txt.IsDisposed Or txt.Disposing) Then
                    If txt.InvokeRequired Then

                        'use action(of ...) to make delegate
                        'Invoke immediately: will wait until done (so that sub will always be executed after)
                        'invokes itself: will fall on "else" when in GUI
                        'should be immediate in case a textbox is used to retrieve data

                        txt.Invoke(New Action(Of TextBox, String)(AddressOf settext_TS), New Object() {txt, s})
                    Else
                        txt.Text = s
                    End If
                End If


            End Sub



            Private Sub runsub(ByVal dsub As IOCallback, ByVal q As IOQuery, ByVal immediate As Boolean)


                If dsub Is Nothing Then
                    Return
                End If

                Dim mon As IOCallback = AddressOf msg_on
                Dim moff As IOCallback = AddressOf msg_off

                Dim cb As Boolean = (dsub IsNot mon And dsub IsNot moff)
                'callback


                If IODevice.frm Is Nothing Then
                    'form not defined eg console app?(didnt test with console app) - but then callback has to be thread-safe!
                    dsub.Invoke(q)
                Else
                    If Not (IODevice.frm.IsDisposed Or IODevice.frm.Disposing) Then
                        If IODevice.frm.InvokeRequired Then
                            Try
                                If immediate And Not parent.disposing Then
                                    'deferred call("post message")
                                    IODevice.frm.Invoke(New Action(Of IOCallback, IOQuery, Boolean)(AddressOf runsub), New Object() {dsub, q, True})
                                Else
                                    IODevice.frm.BeginInvoke(New Action(Of IOCallback, IOQuery, Boolean)(AddressOf runsub), New Object() {dsub, q, False})
                                End If
                            Catch generatedExceptionName As TargetInvocationException
                                ' throw new Exception("cannot invoke callback: called with main form disposed ?"); //may happen on closing app

                                'unhandled exceptions in callback function get here 
                                Debug.Print("TargetInvocationException when invoking callback:  main form disposed?")
                            Catch ex As Exception
                                'if not catched the debugger will stop on "invoke" line, it is more convenient to catch them here and display a message

                                'check in case an error in msg_on
                                If cb And parent.catchcallbackexceptions Then
                                    If q.status = 0 Then
                                        'additional error
                                        q.errmsg = "Exception message: " & ex.Message
                                    Else
                                        q.errmsg = " also: unhandled exception in user callback function: " & ex.Message
                                    End If

                                    q.status += 256
                                    '    'internally using bit 9 to signal callback exception, will be removed by msgon


                                    runsub(AddressOf msg_on, q, False)
                                Else
                                    'rethrow callback exception
                                    Throw

                                End If
                            End Try
                        Else
                            dsub.Invoke(q)
                        End If
                    Else
                        'frm disposed: should not happen unless a device is created after main form is disposed but then we are in trouble!  (fatal error)
                        ' throw new Exception("callback called with main form disposed"); //may happen on closing app
                        Debug.Print("async query error : main form disposed, cannot invoke callback")
                    End If
                End If



            End Sub
            '*****************************************end of inner IOTask class





        End Class



    End Class





End Namespace

'

