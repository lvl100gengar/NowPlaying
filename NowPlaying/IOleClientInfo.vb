'Imports System.Runtime.InteropServices

'Namespace NowPlaying

'    '''<summary>Interface used by Media Player to determine WMP Remoting status.</summary>
'    <ComImport(), ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("CBB92747-741F-44fe-AB5B-F1A48F3B2A59")> _
'    Public Interface IOleClientInfo

'        '''<summary>
'        '''Service type.
'        '''</summary>
'        '''<returns><code>Remote</code> if the control is to be remoted (attached to WMP.) 
'        '''<code>Local</code>if this is an independent WMP instance not connected to WMP application.  If you want local, you shouldn't bother
'        '''using this control!
'        '''</returns>
'        Function GetServiceType() As String

'        '''<summary>
'        '''Value to display in Windows Media Player Switch To Application menu option (under View.)
'        '''</summary>
'        '''<returns>[return: MarshalAs(UnmanagedType.BStr)] </returns>
'        Function GetApplicationName() As <MarshalAs(UnmanagedType.BStr)> String

'        '''<summary>
'        '''Not in use, see MSDN for details.
'        '''</summary>
'        '''<param name="name"></param>
'        '''<param name="dispatch"></param>
'        '''<returns></returns>
'        <PreserveSig()> Function GetScriptableObject(<MarshalAs(UnmanagedType.BStr)> ByRef name As String, <MarshalAs(UnmanagedType.IDispatch)> ByRef dispatch As Object) As <MarshalAs(UnmanagedType.U4)> Integer

'        '''<summary>
'        '''Not in use, see MSDN for details.
'        '''</summary>
'        '''<param name="file"></param>
'        '''<returns></returns>
'        <PreserveSig()> Function GetCustomUIMode(<MarshalAs(UnmanagedType.BStr)> ByRef file As String) As <MarshalAs(UnmanagedType.U4)> Integer

'    End Interface

'End Namespace