Imports System
Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>
''' Contains common function from the User32 library.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class User32

    ''' <summary>
    ''' The EnumDisplayDevices function lets you obtain information about the display devices in the current session.
    ''' </summary>
    ''' <param name="lpDevice">A pointer to the device name. If NULL, function returns information for the display adapter(s) on the machine, based on iDevNum.</param>
    ''' <param name="iDevNum">An index value that specifies the display device of interest.</param>
    ''' <param name="lpDisplayDevice">A pointer to a DISPLAY_DEVICE structure that receives information about the display device specified by iDevNum.  Before calling EnumDisplayDevices, you must initialize the cb member of DISPLAY_DEVICE to the size, in bytes, of DISPLAY_DEVICE.</param>
    ''' <param name="dwFlags">Set this flag to EDD_GET_DEVICE_INTERFACE_NAME (0x00000001) to retrieve the device interface name for GUID_DEVINTERFACE_MONITOR, which is registered by the operating system on a per monitor basis. The value is placed in the DeviceID member of the DISPLAY_DEVICE structure returned in lpDisplayDevice. The resulting device interface name can be used with SetupAPI functions and serves as a link between GDI monitor devices and SetupAPI monitor devices.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. The function fails if iDevNum is greater than the largest device index.</returns>
    ''' <remarks>To query all display devices in the current session, call this function in a loop, starting with iDevNum set to 0, and incrementing iDevNum until the function fails. To select all display devices in the desktop, use only the display devices that have the DISPLAY_DEVICE_ATTACHED_TO_DESKTOP flag in the DISPLAY_DEVICE structure.</remarks>
    <DllImport("user32.dll")> _
    Public Shared Function EnumDisplayDevices(lpDevice As String, iDevNum As UInteger, <Out()> lpDisplayDevice As Structures.DisplayDevice, dwFlags As Constants.DisplayDevicesFlags) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function

    ''' <summary>
    ''' Retrieves information about one of the graphics modes for a display device. To retrieve information for all the graphics modes of a display device, make a series of calls to this function.
    ''' </summary>
    ''' <param name="lpszDeviceName">A pointer to a null-terminated string that specifies the display device about whose graphics mode the function will obtain information. This parameter is either NULL or a DISPLAY_DEVICE.DeviceName returned from EnumDisplayDevices. A NULL value specifies the current display device on the computer on which the calling thread is running.</param>
    ''' <param name="iModeNum">The type of information to be retrieved.</param>
    ''' <param name="lpDevMode">A pointer to a DEVMODE structure into which the function stores information about the specified graphics mode. Before calling EnumDisplaySettings, set the dmSize member to sizeof(DEVMODE), and set the dmDriverExtra member to indicate the size, in bytes, of the additional space available to receive private driver data.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.</returns>
    ''' <remarks>The function fails if iModeNum is greater than the index of the display device's last graphics mode. As noted in the description of the iModeNum parameter, you can use this behavior to enumerate all of a display device's graphics modes.</remarks>
    <DllImport("user32")> _
    Public Shared Function EnumDisplaySettings(lpszDeviceName As String, iModeNum As Integer, <Out()> lpDevMode As Structures.DeviceMode) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function

    ''' <summary>
    ''' Retrieves information about one of the graphics modes for a display device. To retrieve information for all the graphics modes of a display device, make a series of calls to this function.
    ''' </summary>
    ''' <param name="lpszDeviceName">A pointer to a null-terminated string that specifies the display device about whose graphics mode the function will obtain information. This parameter is either NULL or a DISPLAY_DEVICE.DeviceName returned from EnumDisplayDevices. A NULL value specifies the current display device on the computer on which the calling thread is running.</param>
    ''' <param name="iModeNum">The type of information to be retrieved.</param>
    ''' <param name="lpDevMode">A pointer to a DEVMODE structure into which the function stores information about the specified graphics mode. Before calling EnumDisplaySettings, set the dmSize member to sizeof(DEVMODE), and set the dmDriverExtra member to indicate the size, in bytes, of the additional space available to receive private driver data.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.</returns>
    ''' <remarks>The function fails if iModeNum is greater than the index of the display device's last graphics mode. As noted in the description of the iModeNum parameter, you can use this behavior to enumerate all of a display device's graphics modes.</remarks>
    Public Shared Function EnumDisplaySettings(lpszDeviceName As String, iModeNum As Constants.GraphicsMode, <Out()> lpDevMode As Structures.DeviceMode) As <MarshalAs(UnmanagedType.Bool)> Boolean
        Return User32.EnumDisplaySettings(lpszDeviceName, iModeNum, lpDevMode)
    End Function

    ''' <summary>
    ''' Retrieves a handle to the top-level window whose class name and window name match the specified strings. This function does not search child windows. This function does not perform a case-sensitive search.
    ''' </summary>
    ''' <param name="lpClassName">The class name or a class atom created by a previous call to the RegisterClass or RegisterClassEx function. The atom must be in the low-order word of lpClassName; the high-order word must be zero. If lpClassName points to a string, it specifies the window class name. The class name can be any name registered with RegisterClass or RegisterClassEx, or any of the predefined control-class names. If lpClassName is NULL, it finds any window whose title matches the lpWindowName parameter.</param>
    ''' <param name="lpWindowName">The window name (the window's title). If this parameter is NULL, all window names match.</param>
    ''' <returns>If the function succeeds, the return value is a handle to the window that has the specified class name and window name. If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
    ''' <remarks>If the lpWindowName parameter is not NULL, FindWindow calls the GetWindowText function to retrieve the window name for comparison. For a description of a potential problem that can arise, see the Remarks for GetWindowText.</remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr

    End Function

    ''' <summary>
    ''' Retrieves the name of the class to which the specified window belongs.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window and, indirectly, the class to which the window belongs.</param>
    ''' <returns>If the function succeeds, the return value is the class name of the specified window. If the function fails, the return value is String.Empty.</returns>
    ''' <remarks>This is a helper function for the actual GetClassName function.</remarks>
    Public Shared Function GetClassName(hWnd As IntPtr) As String
        Dim className As New StringBuilder("", 256)

        GetClassName(hWnd, className, className.Capacity)

        Return className.ToString()
    End Function

    ''' <summary>
    ''' Retrieves the name of the class to which the specified window belongs.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window and, indirectly, the class to which the window belongs.</param>
    ''' <param name="lpClassName">The class name string.</param>
    ''' <param name="nMaxCount">The length of the lpClassName buffer, in characters. The buffer must be large enough to include the terminating null character; otherwise, the class name string is truncated to nMaxCount-1 characters.</param>
    ''' <remarks>If the function succeeds, the return value is the number of characters copied to the buffer, not including the terminating null character. If the function fails, the return value is zero. To get extended error information, call GetLastError.</remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Public Shared Sub GetClassName(hWnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer)

    End Sub

    ''' <summary>
    ''' Copies the text of the specified window's title bar (if it has one) into a buffer. If the specified window is a control, the text of the control is copied. However, GetWindowText cannot retrieve the text of a control in another application.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window or control containing the text.</param>
    ''' <returns>If the function succeeds, the return value is the text of the specified window or control. If the window has no title bar or text, if the title bar is empty, or if the window or control handle is invalid, the return value is an empty String. To get extended error information, call GetLastError.</returns>
    ''' <remarks>If the target window is owned by the current process, GetWindowText causes a WM_GETTEXT message to be sent to the specified window or control. If the target window is owned by another process and has a caption, GetWindowText retrieves the window caption text. If the window does not have a caption, the return value is a null string. This behavior is by design. It allows applications to call GetWindowText without becoming unresponsive if the process that owns the target window is not responding. However, if the target window is not responding and it belongs to the calling application, GetWindowText will cause the calling application to become unresponsive. To retrieve the text of a control in another process, send a WM_GETTEXT message directly instead of calling GetWindowText.</remarks>
    Public Shared Function GetWindowText(ByVal hWnd As IntPtr) As String
        Dim buffer As New StringBuilder(1024)
        Dim length As Integer = User32.GetWindowText(hWnd, buffer, buffer.Length)

        Return buffer.ToString(0, length)
    End Function

    ''' <summary>
    ''' Copies the text of the specified window's title bar (if it has one) into a buffer. If the specified window is a control, the text of the control is copied. However, GetWindowText cannot retrieve the text of a control in another application.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window or control containing the text.</param>
    ''' <param name="lpString">The buffer that will receive the text. If the string is as long or longer than the buffer, the string is truncated and terminated with a null character.</param>
    ''' <param name="nMaxCount">The maximum number of characters to copy to the buffer, including the null character. If the text exceeds this limit, it is truncated.</param>
    ''' <returns>If the function succeeds, the return value is the length, in characters, of the copied string, not including the terminating null character. If the window has no title bar or text, if the title bar is empty, or if the window or control handle is invalid, the return value is zero. To get extended error information, call GetLastError.</returns>
    ''' <remarks>If the target window is owned by the current process, GetWindowText causes a WM_GETTEXT message to be sent to the specified window or control. If the target window is owned by another process and has a caption, GetWindowText retrieves the window caption text. If the window does not have a caption, the return value is a null string. This behavior is by design. It allows applications to call GetWindowText without becoming unresponsive if the process that owns the target window is not responding. However, if the target window is not responding and it belongs to the calling application, GetWindowText will cause the calling application to become unresponsive. To retrieve the text of a control in another process, send a WM_GETTEXT message directly instead of calling GetWindowText.</remarks>
    <DllImport("user32", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetWindowText(ByVal hWnd As IntPtr, <Out(), MarshalAs(UnmanagedType.LPTStr)> ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer

    End Function

    ''' <summary>
    ''' Retrieves the identifier of the thread that created the specified window and, optionally, the identifier of the process that created the window.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window.</param>
    ''' <param name="lpdwProcessId">A pointer to a variable that receives the process identifier. If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process to the variable; otherwise, it does not.</param>
    ''' <returns>The return value is the identifier of the thread that created the window.</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetWindowThreadProcessId(hWnd As IntPtr, <Out()> ByRef lpdwProcessId As Integer) As Int32

    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows. The SendMessage function calls the window procedure for the specified window and does not return until the window procedure has processed the message.
    ''' </summary>
    ''' <param name="hWnd">A handle to the window whose window procedure will receive the message. If this parameter is HWND_BROADCAST ((HWND)0xffff), the message is sent to all top-level windows in the system, including disabled or invisible unowned windows, overlapped windows, and pop-up windows; but the message is not sent to child windows. Message sending is subject to UIPI. The thread of a process can send messages only to message queues of threads in processes of lesser or equal integrity level.</param>
    ''' <param name="Msg">The message to be sent.</param>
    ''' <param name="wParam">Additional message-specific information.</param>
    ''' <param name="lParam">Additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    ''' <remarks>When a message is blocked by UIPI the last error, retrieved with GetLastError, is set to 5 (access denied). Applications that need to communicate using HWND_BROADCAST should use the RegisterWindowMessage function to obtain a unique message for inter-application communication. The system only does marshalling for system messages (those in the range 0 to (WM_USER-1)). To send other messages (those >= WM_USER) to another process, you must do custom marshalling. If the specified window was created by the calling thread, the window procedure is called immediately as a subroutine. If the specified window was created by a different thread, the system switches to that thread and calls the appropriate window procedure. Messages sent between threads are processed only when the receiving thread executes message retrieval code. The sending thread is blocked until the receiving thread processes the message. However, the sending thread will process incoming nonqueued messages while waiting for its message to be processed. To prevent this, use SendMessageTimeout with SMTO_BLOCK set. For more information on nonqueued messages, see Nonqueued Messages.</remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    End Function

End Class