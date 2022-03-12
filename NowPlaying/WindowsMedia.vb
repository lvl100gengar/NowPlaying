Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.IO

Public Class WindowsMedia
    Implements IMetaProvider

    Private hostControl As Control
    Private playerControl As WindowsMediaRemoteHost
    Private Shared myInstance As New WindowsMedia()

    Public Shared ReadOnly Property Instance As WindowsMedia
        Get
            Return WindowsMedia.myInstance
        End Get
    End Property

    Private Sub New()
        hostControl = Application.OpenForms(0)
    End Sub

    Private Sub InitializeRemoteWMP()
        Try
            playerControl = New NowPlaying.WindowsMediaRemoteHost()
            playerControl.Parent = hostControl
            playerControl.Visible = False

            AddHandler playerControl.Player.SwitchedToControl, AddressOf PlayerControl_SwitchedToControl
        Catch ax As ArgumentException
            'For some reason this exception was throw on the first line above...
            'Setting a breakpoint to maybe find out conditions that cause it.

            '1. This may be a way to tell if a "bind" has failed.
            '2. Not sure why this code is running when wmplayer is not.
            'Log.Add(ax, "InitializeRemoteWMP")
            Exit Sub

        Catch nrx As NullReferenceException
            'This was thrown when closing the WMP application during playback.
            Exit Sub
        End Try
    End Sub

    Private Sub PlayerControl_SwitchedToControl()
        'User has closed Windows Media Player. This Control must be destroyed or it will keep wmplayer.exe running.
        playerControl.Dispose()
        playerControl = Nothing
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not Me.playerControl Is Nothing Then
            Me.playerControl.Dispose()
            Me.playerControl = Nothing
        End If

        GC.SuppressFinalize(Me)
    End Sub

    Public Sub DisplayAttributes()
        Dim media As WMPLib.IWMPMedia = playerControl.Player.currentMedia
        Dim str As String = ""

        For i As Integer = 0 To media.attributeCount - 1
            str += media.getAttributeName(i) + Environment.NewLine
        Next

        Clipboard.SetText(str)
        MessageBox.Show("Done!")
    End Sub

    Public Function GenerateMeta(field As MetaField) As String Implements IMetaProvider.GenerateMeta
        On Error Resume Next

        Dim media As WMPLib.IWMPMedia = playerControl.Player.currentMedia

        Select Case field

            Case MetaField.Album
                Return media.getItemInfo("WM/AlbumTitle")

            Case MetaField.AlbumArtist
                Return media.getItemInfo("WM/AlbumArtist")

            Case MetaField.AlbumLength
                'Return playerControl.Player.mediaCollection.getByAlbum(media.getItemInfo("WM/AlbumTitle")).count.ToString()

            Case MetaField.AlbumPosition
                Return media.getItemInfo("WM/Track")

            Case MetaField.Artist
                Return media.getItemInfo("Author")

            Case MetaField.Bitrate
                Return media.getItemInfo("Bitrate")

            Case MetaField.Channels


            Case MetaField.FIleExtension
                Return Path.GetExtension(media.sourceURL)

            Case MetaField.FileName
                Return Path.GetFileNameWithoutExtension(media.sourceURL)

            Case MetaField.FilePath
                Return media.sourceURL

            Case MetaField.FileSize
                Return New FileInfo(media.sourceURL).Length.ToString()

            Case MetaField.Genre
                Return media.getItemInfo("WM/Genre")

            Case MetaField.PlayCount
                Return media.getItemInfo("UserPlayCount")

            Case MetaField.PlaylistLength
                Return playerControl.Player.currentPlaylist.count.ToString()

            Case MetaField.PlaylistPosition


            Case MetaField.Rating
                Return media.getItemInfo("UserEffectiveRating")

            Case MetaField.SampleRate


            Case MetaField.Title
                Return media.getItemInfo("Title")

            Case MetaField.TrackLength
                Return media.duration.ToString()

            Case MetaField.TrackPosition


            Case MetaField.VideoHeight
                Return media.getItemInfo("WM/VideoHeight")

            Case MetaField.VideoWidth
                Return media.getItemInfo("WM/VideoWidth")

            Case MetaField.Year
                Return media.getItemInfo("WM/Year")

        End Select

        Return "N/A"
    End Function

    Public Function GetState() As ProviderState Implements IMetaProvider.GetState
        If playerControl Is Nothing Then
            'Check to see if Windows Media Player is running, and begin remoting if it is.
            If Diagnostics.Process.GetProcessesByName("wmplayer").Length = 1 Then
                hostControl.BeginInvoke(New Action(AddressOf InitializeRemoteWMP))
            End If

            Return ProviderState.Unavailable
        Else
            Return GetCurrentMediaState()
        End If
    End Function

    Public Function SupportsMeta(field As MetaField) As Boolean Implements IMetaProvider.SupportsMeta
        Return True
    End Function

    Private Function GetCurrentMediaState() As ProviderState
        Try
            Dim paused As Boolean = playerControl.Player.playState = WMPLib.WMPPlayState.wmppsPaused

            'Check for the existance of the currentMedia object.
            If playerControl.Player.currentMedia Is Nothing OrElse playerControl.Player.playState = WMPLib.WMPPlayState.wmppsStopped Then
                Return ProviderState.Stopped
            End If

            'If the media URL is HTTP based, this media is being streamed.
            If playerControl.Player.currentMedia.sourceURL.StartsWith("http://") Then
                Return If(paused, ProviderState.PausedStream, ProviderState.PlayingStream)
            End If

            'Check to see if the player reports that a video is playing.
            If playerControl.Player.currentMedia.getItemInfo("MediaType").Equals("video") Then
                Return If(paused, ProviderState.PausedVideo, ProviderState.PlayingVideo)
            End If

            'If all of the above are false, this must be a standard audio file.
            Return If(paused, ProviderState.PausedAudio, ProviderState.PlayingAudio)
        Catch sx As SystemException
            playerControl.Dispose()
            playerControl = Nothing
            Return ProviderState.Unavailable
        End Try
    End Function

End Class