Imports System.Runtime.InteropServices

'''<summary>
'''Interface used by Windows Media Player to return an instance of IWMPRemoteMediaServices.
'''</summary>
<ComImport(), GuidAttribute("6d5140c1-7436-11ce-8034-00aa006009fa"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown), ComVisible(True)> _
Public Interface IOleServiceProvider

    '''<summary>
    '''Similar to QueryInterface, riid will contain the Guid of an object to return.
    '''In our project we will look for IWMPRemoteMediaServices Guid and return the object
    '''that implements that interface.
    '''</summary>
    '''<param name="guidService"></param>
    '''<param name="riid">The Guid of the desired Service to provide.</param>
    '''<returns>A pointer to the interface requested by the Guid.</returns>
    Function QueryService(ByRef guidservice As Guid, ByRef riid As Guid) As IntPtr

End Interface