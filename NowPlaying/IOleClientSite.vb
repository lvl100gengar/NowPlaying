Imports System.Runtime.InteropServices

'''<summary>'Need to implement this interface so we can pass it to <see cref="IOleObject.SetClientSite"/>.
'''All functions return E_NOTIMPL.  We don't need to actually implement anything to get
'''the remoting to work.</summary>
<ComImport(), ComVisible(True), Guid("00000118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IOleClientSite
    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Sub SaveObject()

    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Function GetMoniker(ByVal dwAssign As UInt32, ByVal dwWhichMoniker As UInt32) As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Function GetContainer() As <MarshalAs(UnmanagedType.Interface)> Object

    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Sub ShowObject()

    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Sub OnShowWindow(ByVal fShow As Boolean)

    '''<summary>
    '''See MSDN for more information.  Throws <see cref="COMException"/> with id of E_NOTIMPL.
    '''</summary>
    '''<exception cref="COMException">E_NOTIMPL</exception>
    Sub RequestNewObjectLayout()

End Interface