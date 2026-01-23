Imports System
Imports System.Runtime.InteropServices
Imports System.Text

Namespace IODevices

    ' Native USBTMC (no VISA). Uses WinUSB + USBTMC framing.
    ' Address format supported:
    '   - "VID=03EB;PID=2065"   (recommended)
    '   - "03EB:2065"
    '   - "USB0::0x03EB::0x2065::....::INSTR"  (VID/PID are parsed, rest ignored)
    '
    ' NOTE: This talks to the USBGPIB adapter (VID/PID), NOT the instrument behind it.

    ' This is not currently used, not currently working

    Public Class XyphroUsbGpibDevice
        Inherits IODevice

        Public Sub New(ByVal name As String, ByVal addr As String)
            MyBase.New(name, addr)
            create(name, addr, False)
        End Sub

        Public Sub New(ByVal name As String, ByVal addr As String, ByVal interlocked As Boolean)
            MyBase.New(name, addr)
            create(name, addr, interlocked)
        End Sub

        '========================
        ' New private fields (you MUST add these)
        '========================
        Private hFile As IntPtr = IntPtr.Zero
        Private hWinUsb As IntPtr = IntPtr.Zero
        Private bulkInPipe As Byte = 0
        Private bulkOutPipe As Byte = 0
        Private bTag As Byte = 1
        Private vid As UShort = &H0
        Private pid As UShort = &H0

        '------------------------
        ' USBTMC constants
        '------------------------
        Private Const USBTMC_MSGID_DEV_DEP_MSG_OUT As Byte = 1
        Private Const USBTMC_MSGID_DEV_DEP_MSG_IN As Byte = 2

        ' USBTMC class-specific requests (USBTMC spec)
        Private Const REQ_INITIATE_CLEAR As Byte = 5
        Private Const REQ_CHECK_CLEAR_STATUS As Byte = 6

        ' USB488 subclass request (USB488 spec) - READ_STATUS_BYTE
        ' Many USBGPIB adapters implement this to support serial poll/MAV.
        Private Const REQ_READ_STATUS_BYTE As Byte = &H80

        Private Const USBTMC_STATUS_SUCCESS As Byte = 1
        Private Const USBTMC_STATUS_PENDING As Byte = 2

        ' bmRequestType bits
        Private Const USB_DIR_IN As Byte = &H80
        Private Const USB_TYPE_CLASS As Byte = &H20
        Private Const USB_RECIP_INTERFACE As Byte = &H1

        ' Win32 CreateFile
        Private Const GENERIC_READ As UInteger = &H80000000UI
        Private Const GENERIC_WRITE As UInteger = &H40000000UI
        Private Const FILE_SHARE_READ As UInteger = &H1UI
        Private Const FILE_SHARE_WRITE As UInteger = &H2UI
        Private Const OPEN_EXISTING As UInteger = 3UI
        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI
        Private Shared ReadOnly INVALID_HANDLE_VALUE As New IntPtr(-1)

        Private Sub create(ByVal name As String, ByVal addr As String, ByVal interlocked As Boolean)
            interfacename = "XYPHRO UsbGpib (Native USBTMC)"
            'syncinterface = interlocked

            ParseVidPid(addr, vid, pid)

            Dim devicePath As String = FindUsbDevicePath(vid, pid)
            If String.IsNullOrEmpty(devicePath) Then
                Throw New Exception($"USBTMC device not found (VID={vid:X4}, PID={pid:X4}).")
            End If

            hFile = CreateFile(devicePath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE,
                               IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero)

            If hFile = INVALID_HANDLE_VALUE OrElse hFile = IntPtr.Zero Then
                Throw New Exception("CreateFile failed for USBTMC device path.")
            End If

            If Not WinUsb_Initialize(hFile, hWinUsb) Then
                Throw New Exception("WinUsb_Initialize failed (is WinUSB driver bound to this interface?).")
            End If

            ' Locate bulk in/out pipes
            DiscoverBulkPipes()

            If bulkInPipe = 0 OrElse bulkOutPipe = 0 Then
                Throw New Exception("Could not locate USBTMC bulk in/out endpoints.")
            End If
        End Sub

        '========================
        ' Required IODevice overrides
        '========================

        Protected Overrides Function Send(ByVal cmd As String, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                Dim data As Byte() = Encoding.ASCII.GetBytes(cmd)
                Return UsbTmcWrite(data, True, errcode, errmsg)
            Catch ex As Exception
                errcode = Marshal.GetHRForException(ex)
                errmsg = ex.Message
                Return -1
            End Try
        End Function

        Protected Overrides Function ReceiveByteArray(ByRef arr As Byte(), ByRef EOI As Boolean, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                Dim data As Byte() = Nothing
                Dim eom As Boolean = False
                Dim rc As Integer = UsbTmcRead(data, eom, errcode, errmsg)
                arr = data
                EOI = eom
                Return rc
            Catch ex As Exception
                errcode = Marshal.GetHRForException(ex)
                errmsg = ex.Message
                arr = New Byte() {}
                EOI = False
                Return -1
            End Try
        End Function

        Protected Overrides Function ClearDevice(ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                ' USBTMC INITIATE_CLEAR + CHECK_CLEAR_STATUS loop (USBTMC spec)
                Dim buf1(0) As Byte
                Dim transferred As UInteger = 0

                ' INITIATE_CLEAR (IN, CLASS, INTERFACE)
                If Not WinUsb_ControlTransfer(hWinUsb, MakeSetupPacket(USB_DIR_IN Or USB_TYPE_CLASS Or USB_RECIP_INTERFACE,
                                                                      REQ_INITIATE_CLEAR, 0US, 0US, 1US),
                                              buf1, CUInt(buf1.Length), transferred, IntPtr.Zero) Then
                    errcode = Marshal.GetLastWin32Error()
                    errmsg = "INITIATE_CLEAR failed."
                    Return -1
                End If

                If buf1(0) <> USBTMC_STATUS_SUCCESS Then
                    errcode = buf1(0)
                    errmsg = "INITIATE_CLEAR returned status " & buf1(0).ToString()
                    Return -1
                End If

                ' CHECK_CLEAR_STATUS until SUCCESS (some devices return PENDING briefly)
                Dim buf2(1) As Byte
                For i As Integer = 0 To 40 ' ~2s total at 50ms
                    transferred = 0
                    If Not WinUsb_ControlTransfer(hWinUsb, MakeSetupPacket(USB_DIR_IN Or USB_TYPE_CLASS Or USB_RECIP_INTERFACE,
                                                                          REQ_CHECK_CLEAR_STATUS, 0US, 0US, 2US),
                                                  buf2, CUInt(buf2.Length), transferred, IntPtr.Zero) Then
                        errcode = Marshal.GetLastWin32Error()
                        errmsg = "CHECK_CLEAR_STATUS failed."
                        Return -1
                    End If

                    If buf2(0) = USBTMC_STATUS_SUCCESS Then
                        Return 0
                    End If

                    If buf2(0) <> USBTMC_STATUS_PENDING Then
                        errcode = buf2(0)
                        errmsg = "CHECK_CLEAR_STATUS returned status " & buf2(0).ToString()
                        Return -1
                    End If

                    Threading.Thread.Sleep(50)
                Next

                errcode = -1
                errmsg = "CLEAR timeout."
                Return 1 'timeout style
            Catch ex As Exception
                errcode = Marshal.GetHRForException(ex)
                errmsg = ex.Message
                Return -1
            End Try
        End Function

        Protected Overrides Function PollMAV(ByRef mav As Boolean, ByRef statusbyte As Byte, ByRef errcode As Integer, ByRef errmsg As String) As Integer
            Try
                ' USB488 READ_STATUS_BYTE (class request 0x80) if implemented by the adapter.
                ' Returns 2 bytes: [0]=USBTMC status, [1]=IEEE488 status byte (typical behavior).
                Dim resp(1) As Byte
                Dim transferred As UInteger = 0

                Dim tag As Byte = NextTag()
                Dim wValue As UShort = tag ' many implementations use bTag in wValue

                If Not WinUsb_ControlTransfer(hWinUsb, MakeSetupPacket(USB_DIR_IN Or USB_TYPE_CLASS Or USB_RECIP_INTERFACE,
                                                                      REQ_READ_STATUS_BYTE, wValue, 0US, 2US),
                                              resp, CUInt(resp.Length), transferred, IntPtr.Zero) Then
                    ' If not supported, don’t hard-fail the whole app.
                    statusbyte = 0
                    mav = False
                    errcode = 0
                    errmsg = ""
                    Return 0
                End If

                ' If device returns a status code in resp(0), SUCCESS usually = 1
                If resp(0) <> USBTMC_STATUS_SUCCESS Then
                    statusbyte = 0
                    mav = False
                    errcode = resp(0)
                    errmsg = "READ_STATUS_BYTE returned status " & resp(0).ToString()
                    Return -1
                End If

                statusbyte = resp(1)
                ' MAV bit is bit4 (0x10) in IEEE 488.2 status byte
                mav = ((statusbyte And &H10) <> 0)
                Return 0

            Catch ex As Exception
                errcode = Marshal.GetHRForException(ex)
                errmsg = ex.Message
                statusbyte = 0
                mav = False
                Return -1
            End Try
        End Function

        Protected Overrides Sub DisposeDevice()
            Try
                If hWinUsb <> IntPtr.Zero Then
                    WinUsb_Free(hWinUsb)
                    hWinUsb = IntPtr.Zero
                End If
            Catch
            End Try

            Try
                If hFile <> IntPtr.Zero AndAlso hFile <> INVALID_HANDLE_VALUE Then
                    CloseHandle(hFile)
                    hFile = IntPtr.Zero
                End If
            Catch
            End Try
        End Sub


        '========================
        ' USBTMC bulk transfer helpers
        '========================

        Private Function UsbTmcWrite(ByVal payload As Byte(), ByVal eoi As Boolean,
                                     ByRef errcode As Integer, ByRef errmsg As String) As Integer
            errcode = 0 : errmsg = ""

            Dim tag As Byte = NextTag()

            Dim header(11) As Byte
            header(0) = USBTMC_MSGID_DEV_DEP_MSG_OUT
            header(1) = tag
            header(2) = CByte(&HFF - tag)
            header(3) = 0

            Dim n As UInteger = CUInt(payload.Length)
            WriteUInt32LE(header, 4, n)

            header(8) = If(eoi, CByte(1), CByte(0)) ' bmTransferAttributes: EOM/EOI
            header(9) = 0
            header(10) = 0
            header(11) = 0

            ' Pad to 4-byte multiple (USBTMC rule)
            Dim totalLen As Integer = header.Length + payload.Length
            Dim pad As Integer = (4 - (totalLen Mod 4)) Mod 4

            Dim outBuf(totalLen + pad - 1) As Byte
            Buffer.BlockCopy(header, 0, outBuf, 0, header.Length)
            If payload.Length > 0 Then Buffer.BlockCopy(payload, 0, outBuf, header.Length, payload.Length)
            ' pad bytes are already 0

            Dim written As UInteger = 0
            If Not WinUsb_WritePipe(hWinUsb, bulkOutPipe, outBuf, CUInt(outBuf.Length), written, IntPtr.Zero) Then
                errcode = Marshal.GetLastWin32Error()
                errmsg = "WinUsb_WritePipe failed."
                Return -1
            End If

            Return 0
        End Function

        Private Function UsbTmcRead(ByRef data As Byte(), ByRef eom As Boolean,
                                    ByRef errcode As Integer, ByRef errmsg As String) As Integer
            errcode = 0 : errmsg = ""
            eom = False

            ' request size: your IODevice base typically controls actual read size by repeated reads.
            ' We'll request a generous block; caller can re-read if needed.
            Dim requestSize As UInteger = 1024UI * 64UI

            Dim tag As Byte = NextTag()

            ' Send DEV_DEP_MSG_IN request header on Bulk-OUT
            Dim inHdr(11) As Byte
            inHdr(0) = USBTMC_MSGID_DEV_DEP_MSG_IN
            inHdr(1) = tag
            inHdr(2) = CByte(&HFF - tag)
            inHdr(3) = 0
            WriteUInt32LE(inHdr, 4, requestSize)
            inHdr(8) = 0 ' bmTransferAttributes (no TERM char)
            inHdr(9) = 0
            inHdr(10) = 0
            inHdr(11) = 0

            Dim written As UInteger = 0
            If Not WinUsb_WritePipe(hWinUsb, bulkOutPipe, inHdr, CUInt(inHdr.Length), written, IntPtr.Zero) Then
                errcode = Marshal.GetLastWin32Error()
                errmsg = "WinUsb_WritePipe (DEV_DEP_MSG_IN) failed."
                data = New Byte() {}
                Return -1
            End If

            ' Now read response from Bulk-IN (header + payload)
            ' Read a big chunk; device will return as much as it has.
            Dim inBuf(12 + CInt(requestSize) - 1) As Byte
            Dim read As UInteger = 0

            If Not WinUsb_ReadPipe(hWinUsb, bulkInPipe, inBuf, CUInt(inBuf.Length), read, IntPtr.Zero) Then
                errcode = Marshal.GetLastWin32Error()
                errmsg = "WinUsb_ReadPipe failed."
                data = New Byte() {}
                Return -1
            End If

            If read < 12UI Then
                errcode = -1
                errmsg = "Short USBTMC header."
                data = New Byte() {}
                Return -1
            End If

            ' Parse header
            ' inBuf(0)=MsgID, inBuf(1)=bTag, etc.
            Dim transferSize As UInteger = ReadUInt32LE(inBuf, 4)
            Dim attrs As Byte = inBuf(8)

            eom = ((attrs And 1) <> 0) ' EOM bit

            Dim available As Integer = CInt(Math.Min(transferSize, read - 12UI))
            If available < 0 Then available = 0

            data = New Byte(available - 1) {}
            If available > 0 Then Buffer.BlockCopy(inBuf, 12, data, 0, available)

            Return 0
        End Function

        Private Sub DiscoverBulkPipes()
            Dim iface As WINUSB_INTERFACE_DESCRIPTOR
            If Not WinUsb_QueryInterfaceSettings(hWinUsb, 0, iface) Then
                Throw New Exception("WinUsb_QueryInterfaceSettings failed.")
            End If

            For i As Integer = 0 To iface.bNumEndpoints - 1
                Dim pipe As WINUSB_PIPE_INFORMATION
                If WinUsb_QueryPipe(hWinUsb, 0, CByte(i), pipe) Then
                    If pipe.PipeType = USBD_PIPE_TYPE.UsbdPipeTypeBulk Then
                        If (pipe.PipeId And &H80) <> 0 Then
                            bulkInPipe = pipe.PipeId
                        Else
                            bulkOutPipe = pipe.PipeId
                        End If
                    End If
                End If
            Next
        End Sub

        Private Function NextTag() As Byte
            Dim t As Byte = bTag
            bTag = CByte((bTag Mod 255) + 1) ' 1..255
            Return t
        End Function

        Private Shared Sub WriteUInt32LE(ByVal buf As Byte(), ByVal offset As Integer, ByVal value As UInteger)
            buf(offset + 0) = CByte(value And &HFFUI)
            buf(offset + 1) = CByte((value >> 8) And &HFFUI)
            buf(offset + 2) = CByte((value >> 16) And &HFFUI)
            buf(offset + 3) = CByte((value >> 24) And &HFFUI)
        End Sub

        Private Shared Function ReadUInt32LE(ByVal buf As Byte(), ByVal offset As Integer) As UInteger
            Return CUInt(buf(offset)) Or (CUInt(buf(offset + 1)) << 8) Or (CUInt(buf(offset + 2)) << 16) Or (CUInt(buf(offset + 3)) << 24)
        End Function

        '========================
        ' Address parsing
        '========================
        Private Shared Sub ParseVidPid(ByVal addr As String, ByRef vid As UShort, ByRef pid As UShort)
            Dim s As String = addr.Trim()

            Dim v As Integer = -1, p As Integer = -1

            ' VISA-like: USB0::0x03EB::0x2065::...
            Dim m1 = Text.RegularExpressions.Regex.Match(s, "0x([0-9A-Fa-f]{4})::0x([0-9A-Fa-f]{4})")
            If m1.Success Then
                v = Convert.ToInt32(m1.Groups(1).Value, 16)
                p = Convert.ToInt32(m1.Groups(2).Value, 16)
            End If

            ' VID=03EB;PID=2065
            If v < 0 Then
                Dim m2 = Text.RegularExpressions.Regex.Match(s, "VID\s*=\s*([0-9A-Fa-f]{4}).*PID\s*=\s*([0-9A-Fa-f]{4})")
                If m2.Success Then
                    v = Convert.ToInt32(m2.Groups(1).Value, 16)
                    p = Convert.ToInt32(m2.Groups(2).Value, 16)
                End If
            End If

            ' 03EB:2065
            If v < 0 Then
                Dim m3 = Text.RegularExpressions.Regex.Match(s, "([0-9A-Fa-f]{4})\s*:\s*([0-9A-Fa-f]{4})")
                If m3.Success Then
                    v = Convert.ToInt32(m3.Groups(1).Value, 16)
                    p = Convert.ToInt32(m3.Groups(2).Value, 16)
                End If
            End If

            If v < 0 OrElse p < 0 Then
                Throw New Exception("Invalid address format for native USBTMC. Use e.g. VID=03EB;PID=2065")
            End If

            vid = CUShort(v)
            pid = CUShort(p)
        End Sub

        '========================
        ' SetupAPI device enumeration by VID/PID
        '========================
        Private Shared Function FindUsbDevicePath(ByVal vid As UShort, ByVal pid As UShort) As String
            ' Enumerate all present USB device interfaces and match by VID/PID in the device path string.
            Dim guid As Guid = GUID_DEVINTERFACE_USB_DEVICE

            Dim hInfo = SetupDiGetClassDevs(guid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)
            If hInfo = IntPtr.Zero OrElse hInfo = New IntPtr(-1) Then Return Nothing

            Try
                Dim index As UInteger = 0
                Dim did As New SP_DEVICE_INTERFACE_DATA()
                did.cbSize = CUInt(Marshal.SizeOf(did))

                Dim match As String = $"vid_{vid:X4}&pid_{pid:X4}"

                While SetupDiEnumDeviceInterfaces(hInfo, IntPtr.Zero, guid, index, did)
                    ' get required size
                    Dim required As UInteger = 0
                    SetupDiGetDeviceInterfaceDetail(hInfo, did, IntPtr.Zero, 0, required, IntPtr.Zero)

                    Dim detailBuffer As IntPtr = Marshal.AllocHGlobal(CInt(required))
                    Try
                        ' cbSize differs x86/x64
                        Dim cbSize As Integer = If(IntPtr.Size = 8, 8, 6)
                        Marshal.WriteInt32(detailBuffer, cbSize)

                        If SetupDiGetDeviceInterfaceDetail(hInfo, did, detailBuffer, required, required, IntPtr.Zero) Then
                            Dim pDevicePath As IntPtr = IntPtr.Add(detailBuffer, 4)
                            If IntPtr.Size = 8 Then pDevicePath = IntPtr.Add(detailBuffer, 8)

                            Dim path As String = Marshal.PtrToStringAuto(pDevicePath)
                            If path IsNot Nothing AndAlso path.ToLower().Contains(match.ToLower()) Then
                                Return path
                            End If
                        End If
                    Finally
                        Marshal.FreeHGlobal(detailBuffer)
                    End Try

                    index += 1UI
                End While

                Return Nothing
            Finally
                SetupDiDestroyDeviceInfoList(hInfo)
            End Try
        End Function

        '========================
        ' WinUSB + SetupAPI P/Invokes
        '========================

        <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Shared Function CreateFile(ByVal lpFileName As String,
                                           ByVal dwDesiredAccess As UInteger,
                                           ByVal dwShareMode As UInteger,
                                           ByVal lpSecurityAttributes As IntPtr,
                                           ByVal dwCreationDisposition As UInteger,
                                           ByVal dwFlagsAndAttributes As UInteger,
                                           ByVal hTemplateFile As IntPtr) As IntPtr
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function CloseHandle(ByVal hObject As IntPtr) As Boolean
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINUSB_INTERFACE_DESCRIPTOR
            Public bLength As Byte
            Public bDescriptorType As Byte
            Public bInterfaceNumber As Byte
            Public bAlternateSetting As Byte
            Public bNumEndpoints As Byte
            Public bInterfaceClass As Byte
            Public bInterfaceSubClass As Byte
            Public bInterfaceProtocol As Byte
            Public iInterface As Byte
        End Structure

        Private Enum USBD_PIPE_TYPE As Integer
            UsbdPipeTypeControl = 0
            UsbdPipeTypeIsochronous = 1
            UsbdPipeTypeBulk = 2
            UsbdPipeTypeInterrupt = 3
        End Enum

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINUSB_PIPE_INFORMATION
            Public PipeType As USBD_PIPE_TYPE
            Public PipeId As Byte
            Public MaximumPacketSize As UShort
            Public Interval As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINUSB_SETUP_PACKET
            Public RequestType As Byte
            Public Request As Byte
            Public Value As UShort
            Public Index As UShort
            Public Length As UShort
        End Structure

        Private Shared Function MakeSetupPacket(ByVal bmRequestType As Byte, ByVal bRequest As Byte,
                                                ByVal wValue As UShort, ByVal wIndex As UShort, ByVal wLength As UShort) As WINUSB_SETUP_PACKET
            Dim s As New WINUSB_SETUP_PACKET()
            s.RequestType = bmRequestType
            s.Request = bRequest
            s.Value = wValue
            s.Index = wIndex
            s.Length = wLength
            Return s
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_Initialize(ByVal DeviceHandle As IntPtr, ByRef InterfaceHandle As IntPtr) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_Free(ByVal InterfaceHandle As IntPtr) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_QueryInterfaceSettings(ByVal InterfaceHandle As IntPtr,
                                                              ByVal AlternateInterfaceNumber As Byte,
                                                              ByRef UsbAltInterfaceDescriptor As WINUSB_INTERFACE_DESCRIPTOR) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_QueryPipe(ByVal InterfaceHandle As IntPtr,
                                                 ByVal AlternateInterfaceNumber As Byte,
                                                 ByVal PipeIndex As Byte,
                                                 ByRef PipeInformation As WINUSB_PIPE_INFORMATION) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_ReadPipe(ByVal InterfaceHandle As IntPtr,
                                                ByVal PipeID As Byte,
                                                <Out> ByVal Buffer As Byte(),
                                                ByVal BufferLength As UInteger,
                                                ByRef LengthTransferred As UInteger,
                                                ByVal Overlapped As IntPtr) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_WritePipe(ByVal InterfaceHandle As IntPtr,
                                                 ByVal PipeID As Byte,
                                                 ByVal Buffer As Byte(),
                                                 ByVal BufferLength As UInteger,
                                                 ByRef LengthTransferred As UInteger,
                                                 ByVal Overlapped As IntPtr) As Boolean
        End Function

        <DllImport("winusb.dll", SetLastError:=True)>
        Private Shared Function WinUsb_ControlTransfer(ByVal InterfaceHandle As IntPtr,
                                                       ByVal SetupPacket As WINUSB_SETUP_PACKET,
                                                       ByVal Buffer As Byte(),
                                                       ByVal BufferLength As UInteger,
                                                       ByRef LengthTransferred As UInteger,
                                                       ByVal Overlapped As IntPtr) As Boolean
        End Function

        ' SetupAPI
        Private Const DIGCF_PRESENT As UInteger = &H2UI
        Private Const DIGCF_DEVICEINTERFACE As UInteger = &H10UI

        Private Shared ReadOnly GUID_DEVINTERFACE_USB_DEVICE As New Guid("A5DCBF10-6530-11D2-901F-00C04FB951ED")

        <StructLayout(LayoutKind.Sequential)>
        Private Structure SP_DEVICE_INTERFACE_DATA
            Public cbSize As UInteger
            Public InterfaceClassGuid As Guid
            Public Flags As UInteger
            Public Reserved As IntPtr
        End Structure

        <DllImport("setupapi.dll", SetLastError:=True)>
        Private Shared Function SetupDiGetClassDevs(ByRef ClassGuid As Guid,
                                                    ByVal Enumerator As IntPtr,
                                                    ByVal hwndParent As IntPtr,
                                                    ByVal Flags As UInteger) As IntPtr
        End Function

        <DllImport("setupapi.dll", SetLastError:=True)>
        Private Shared Function SetupDiEnumDeviceInterfaces(ByVal DeviceInfoSet As IntPtr,
                                                            ByVal DeviceInfoData As IntPtr,
                                                            ByRef InterfaceClassGuid As Guid,
                                                            ByVal MemberIndex As UInteger,
                                                            ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Boolean
        End Function

        <DllImport("setupapi.dll", SetLastError:=True)>
        Private Shared Function SetupDiGetDeviceInterfaceDetail(ByVal DeviceInfoSet As IntPtr,
                                                                ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA,
                                                                ByVal DeviceInterfaceDetailData As IntPtr,
                                                                ByVal DeviceInterfaceDetailDataSize As UInteger,
                                                                ByRef RequiredSize As UInteger,
                                                                ByVal DeviceInfoData As IntPtr) As Boolean
        End Function

        <DllImport("setupapi.dll", SetLastError:=True)>
        Private Shared Function SetupDiDestroyDeviceInfoList(ByVal DeviceInfoSet As IntPtr) As Boolean
        End Function

    End Class

End Namespace
