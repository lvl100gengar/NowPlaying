Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports WMPLib

<ComVisible(True), ClassInterface(ClassInterfaceType.None)> _
Public Class WindowsMediaRemoteHostInfo
    Implements IWMPRemoteMediaServices

    'Public Function GetServiceType() As String Implements IOleClientInfo.GetServiceType
    '    Return "Remote"
    'End Function

    'Public Function GetApplicationName() As String Implements IOleClientInfo.GetApplicationName
    '    Return Process.GetCurrentProcess().ProcessName
    'End Function

    'Public Function GetScritableObject(ByRef name As String, ByRef dispatch As Object) As Integer Implements IOleClientInfo.GetScriptableObject
    '    name = Nothing
    '    dispatch = Nothing

    '    Return -2147467263
    'End Function

    'Public Function GetCustomUIMode(ByRef file As String) As Integer Implements IOleClientInfo.GetCustomUIMode
    '    file = Nothing

    '    Return -2147467263
    'End Function

    Public Sub GetApplicationName(ByRef pbstrName As String) Implements WMPLib.IWMPRemoteMediaServices.GetApplicationName
        pbstrName = Process.GetCurrentProcess().ProcessName
    End Sub

    Public Sub GetCustomUIMode(ByRef pbstrFile As String) Implements WMPLib.IWMPRemoteMediaServices.GetCustomUIMode
        pbstrFile = Nothing
    End Sub

    Public Sub GetScriptableObject(ByRef pbstrName As String, ByRef ppDispatch As Object) Implements WMPLib.IWMPRemoteMediaServices.GetScriptableObject
        pbstrName = Nothing
        ppDispatch = Nothing
    End Sub

    Public Sub GetServiceType(ByRef pbstrType As String) Implements WMPLib.IWMPRemoteMediaServices.GetServiceType
        pbstrType = "Remote"
    End Sub

End Class