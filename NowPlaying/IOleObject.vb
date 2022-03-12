Imports System.Runtime.InteropServices

'''<summary>This interface is implemented by WMP ActiveX/COM control. The only function we need is "SetClientSite".</summary>
<ComImport(), ComVisible(True), Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IOleObject

    '''<summary>
    '''Used to pass our custom <see cref="IOleClientSite"/> object to WMP.  The object we pass must also
    '''implement <see cref="IOleServiceProvider"/> to work right.
    '''</summary>
    '''<param name="pClientSite">The <see cref="IOleClientSite"/> to pass.</param>
    Sub SetClientSite(ByVal pClientSite As IOleClientSite)
    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetClientSite() As <MarshalAs(UnmanagedType.Interface)> IOleClientSite

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub SetHostName(<MarshalAs(UnmanagedType.LPWStr)> ByVal szContainerApp As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal szContainerObj As String)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub Close(ByVal dwSaveOption As UInt32)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub SetMoniker(ByVal dwWhichMoniker As UInt32, ByVal pmk As Object)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetMoniker(ByVal dwAssign As UInt32, ByVal dwWhichMoniker As UInt32) As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub InitFromData(ByVal pDataObject As Object, ByVal fCreation As Boolean, ByVal dwReserved As UInt32)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetClipboardDate(ByVal dwReserved As UInt32) As Object


    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub DoVerb(ByVal iVerb As UInt32, ByVal lpmsg As UInt32, <MarshalAs(UnmanagedType.Interface)> ByVal pActiveSite As Object, ByVal lindex As UInt32, ByVal hwndparent As UInt32, ByVal lprcPosRect As UInt32)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function EnumVerbs() As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub Update()

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    <PreserveSig()> _
    Function IsUpToDate() As <MarshalAs(UnmanagedType.U4)> Integer

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetUserClassID() As Guid

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetUserType(ByVal dwFormOfType As UInt32) As <MarshalAs(UnmanagedType.LPWStr)> String

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub SetExtent(ByVal dwDrawAspect As UInt32, <MarshalAs(UnmanagedType.Interface)> ByVal psizel As Object)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetExtent(ByVal dwDrawAspect As UInt32) As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function Advise(<MarshalAs(UnmanagedType.Interface)> ByVal pAdvSink As Object) As UInt32

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub Unadvise(ByVal dwConnection As UInt32)

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function EnumAdvise() As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Function GetMiscStatus(<MarshalAs(UnmanagedType.U4)> ByVal dwAspect As Integer) As UInt32

    '''<summary>
    '''Implemented by Windows Media Player ActiveX control.
    '''See MSDN for more information.
    '''</summary>
    Sub SetColorScheme(<MarshalAs(UnmanagedType.Interface)> ByVal pLogpal As Object)

End Interface