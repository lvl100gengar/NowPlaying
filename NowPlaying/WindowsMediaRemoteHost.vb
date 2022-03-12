Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports WMPLib

<AxHost.ClsidAttribute("{6bf52a52-394a-11d3-b153-00c04f79faa6}"), ComVisible(True), ClassInterface(ClassInterfaceType.AutoDispatch)> _
Public Class WindowsMediaRemoteHost
    Inherits AxHost
    Implements IOleClientSite
    Implements IOleServiceProvider

    Private myPlayer As WindowsMediaPlayer

    Public ReadOnly Property Player() As WindowsMediaPlayer
        Get
            Return myPlayer
        End Get
    End Property

    Public Sub New()
        MyBase.New("6bf52a52-394a-11d3-b153-00c04f79faa6")
    End Sub

    '''<summary>Used to attach the appropriate interface to Windows Media Player. We call SetClientSite on the WMP Control, passing it this instance.</summary>
    Protected Overrides Sub AttachInterfaces()
        Try
            'Get the IOleObject for Windows Media Player.
            Dim oleObject As IOleObject = DirectCast(Me.GetOcx, IOleObject)

            'Set the Client Site for the WMP control.
            oleObject.SetClientSite(Me)

            myPlayer = DirectCast(Me.GetOcx(), WMPLib.WindowsMediaPlayer)
        Catch avx As AccessViolationException
            'Do nothing.
        Catch ex As Exception
            'Do nothing.
        End Try
    End Sub

    '''<summary>
    '''During SetClientSite, WMP calls this function to get the pointer to RemoteHostInfo.
    '''</summary>
    '''<param name="guidService">See MSDN for more information - we do not use this parameter.</param>
    '''<param name="riid">The Guid of the desired service to be returned.  For this application it will always match
    '''the Guid of <see cref="IWMPRemoteMediaServices"/>.</param>
    '''<returns></returns>
    Public Function QueryService(ByRef guidservice As Guid, ByRef riid As Guid) As IntPtr Implements IOleServiceProvider.QueryService
        'If we get to here, it means Media Player is requesting our IWMPRemoteMediaServices interface
        If (riid.Equals(New Guid("cbb92747-741f-44fe-ab5b-f1a48f3b2a59"))) Then
            Dim iwmp As IWMPRemoteMediaServices = New WindowsMediaRemoteHostInfo()
            Return Marshal.GetComInterfaceForObject(iwmp, GetType(IWMPRemoteMediaServices))
        End If

        Throw New COMException("No Interface", -2147467262)
    End Function

    Public Sub SaveObject() Implements IOleClientSite.SaveObject
        Throw New COMException("Not Implemented", -2147467263)
    End Sub

    Public Function GetMoniker(ByVal dwAssign As UInt32, ByVal dwWhichMoniker As UInt32) As Object Implements IOleClientSite.GetMoniker
        Throw New COMException("Not Implemented", -2147467263)
    End Function

    Public Function GetContainer() As Object Implements IOleClientSite.GetContainer
        'Throw New COMException("Not Implemented", -2147467263)
        Return Me
    End Function

    Public Sub ShowObject() Implements IOleClientSite.ShowObject
        Throw New COMException("Not Implemented", -2147467263)
    End Sub

    Public Sub OnShowWindow(ByVal fShow As Boolean) Implements IOleClientSite.OnShowWindow
        Throw New COMException("Not Implemented", -2147467263)
    End Sub

    Public Sub RequestNewObjectLayout() Implements IOleClientSite.RequestNewObjectLayout
        Throw New COMException("Not Implemented", -2147467263)
    End Sub

End Class