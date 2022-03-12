Imports System.Net
Imports System.IO
Imports System.Xml.Serialization

Public Class Videolan
    Implements IMetaProvider

    Private mediaInfo As VideolanStatus
    Private mediaList As VideolanNode
    Private Shared myInstance As New Videolan

    'Public ReadOnly Property Channels As String Implements IMediaInfoProvider.Channels
    '    Get
    '        Return Me.GetAttribute("Audio", "Channels")
    '    End Get
    'End Property

    'Public ReadOnly Property Codec As String Implements IMediaInfoProvider.Codec
    '    Get
    '        Return Me.GetAttribute("Audio", "Codec")
    '    End Get
    'End Property

    'Public ReadOnly Property Position As String Implements IMediaInfoProvider.Position
    '    Get
    '        Return (Me.mediaInfo.Position * Me.mediaInfo.Length) / 100
    '    End Get
    'End Property

    'Public ReadOnly Property SampleRate As String Implements IMediaInfoProvider.SampleRate
    '    Get
    '        Return Me.GetAttribute("Audio", "Sample rate")
    '    End Get
    'End Property

    Public Shared ReadOnly Property Instance As Videolan
        Get
            Return myInstance
        End Get
    End Property

    Private Sub New()

    End Sub

    Private ReadOnly Property FileName As String
        Get
            Dim currentItem As VideolanItem = Nothing

            If Me.mediaList.FindCurrent(Nothing, currentItem) Then
                Return Uri.UnescapeDataString(currentItem.URI.Replace("file:///", String.Empty))
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Private ReadOnly Property PlaylistLength As Integer
        Get
            Dim currentPlaylist As VideolanNode = Nothing

            If Me.mediaList.FindCurrent(currentPlaylist, Nothing) Then
                Return currentPlaylist.Items.Length
            Else
                Return 0
            End If
        End Get
    End Property

    Private ReadOnly Property PlaylistPosition As Integer
        Get
            Dim currentItem As VideolanItem = Nothing
            Dim currentNode As VideolanNode = Nothing

            If Me.mediaList.FindCurrent(currentNode, currentItem) Then
                Return Array.IndexOf(Of VideolanItem)(currentNode.Items, currentItem) + 1
            Else
                Return 0
            End If
        End Get
    End Property

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Private Function GetAttribute(ByVal streamType As String, ByVal attributeName As String) As String
        Dim streamInfo As VideolanStreamInfo = Me.mediaInfo.Information.Streams(streamType)

        If streamInfo Is Nothing Then
            Return String.Empty
        Else
            Return streamInfo.Attributes(attributeName)
        End If
    End Function

    Public Function GenerateMeta(field As MetaField) As String Implements IMetaProvider.GenerateMeta
        'In the event of an error getting any of the meta data, let the function return the default value.
        On Error Resume Next

        Select Case field

            Case MetaField.Album
                Return Me.mediaInfo.Information.MetaInfo.Album

            Case MetaField.AlbumArtist


            Case MetaField.AlbumLength
                Return Me.PlaylistLength.ToString()

            Case MetaField.AlbumPosition
                Return Me.mediaInfo.Information.MetaInfo.Track

            Case MetaField.Artist
                Return Me.mediaInfo.Information.MetaInfo.Artist

            Case MetaField.Bitrate
                Return Me.GetAttribute("Audio", "Bitrate")

            Case MetaField.Channels
                Return Me.GetAttribute("Audio", "Channels")

            Case MetaField.FIleExtension
                Return Path.GetExtension(Me.FileName)

            Case MetaField.FileName
                Return Path.GetFileNameWithoutExtension(Me.FileName)

            Case MetaField.FilePath
                Return Me.FileName

            Case MetaField.FileSize
                Return New FileInfo(Me.FileName).Length.ToString()

            Case MetaField.Genre
                Return Me.mediaInfo.Information.MetaInfo.Genre

            Case MetaField.PlayCount


            Case MetaField.PlaylistLength
                Return Me.PlaylistLength

            Case MetaField.PlaylistPosition
                Return Me.PlaylistPosition

            Case MetaField.RadioTitle
                Return Me.ReplaceHTTPEntities(Me.mediaInfo.Information.MetaInfo.NowPlaying)

            Case MetaField.Rating
                Return Me.mediaInfo.Information.MetaInfo.Rating

            Case MetaField.SampleRate
                Return Me.GetAttribute("Audio", "SampleRate")

            Case MetaField.Title
                Return Me.mediaInfo.Information.MetaInfo.Title

            Case MetaField.TrackLength
                Return Me.mediaInfo.Length.ToString()

            Case MetaField.TrackPosition
                Return Me.mediaInfo.Time.ToString()

            Case MetaField.VideoHeight
                Return Me.GetAttribute("Video", "Resolution")

            Case MetaField.VideoWidth
                Return Me.GetAttribute("Video", "Resolution")

            Case MetaField.Year
                Return Me.mediaInfo.Information.MetaInfo.Date

        End Select

        Return "N/A"
    End Function

    Private Function ReplaceHTTPEntities(value As String) As String
        value = value.Replace("&amp;", "&")
        value = value.Replace("&#39;", "'")

        Return value
    End Function

    Public Function GetState() As ProviderState Implements IMetaProvider.GetState
        Dim request1 As WebRequest = HttpWebRequest.Create("http://localhost:8080/requests/status.xml")
        Dim request2 As WebRequest = HttpWebRequest.Create("http://localhost:8080/requests/playlist.xml")

        Try
            Using response As WebResponse = request1.GetResponse()
                Using data As Stream = response.GetResponseStream()
                    Me.mediaInfo = Markup(Of VideolanStatus).Read(data)
                End Using
            End Using

            Using response As WebResponse = request2.GetResponse()
                Using data As Stream = response.GetResponseStream()
                    Me.mediaList = Markup(Of VideolanNode).Read(data)
                End Using
            End Using
        Catch wx As WebException
            'The player is closed or the HTTP server is disabled.
            Return ProviderState.Unavailable
        End Try

        If Me.mediaInfo Is Nothing OrElse Me.mediaList Is Nothing OrElse Me.mediaInfo.State.Equals("stop") Then
            Return ProviderState.Stopped
        Else
            Dim paused As Boolean = Me.mediaInfo.State.Equals("paused")

            If Me.FileName.StartsWith("http://") Then
                Return If(paused, ProviderState.PausedStream, ProviderState.PlayingStream)
            End If

            If Me.mediaInfo.Information.Streams("Video") IsNot Nothing Then
                Return If(paused, ProviderState.PausedVideo, ProviderState.PlayingVideo)
            End If

            Return If(paused, ProviderState.PausedAudio, ProviderState.PlayingAudio)
        End If
    End Function

    Public Function SupportsMeta(field As MetaField) As Boolean Implements IMetaProvider.SupportsMeta
        Return True
    End Function

End Class

Public NotInheritable Class Markup(Of T)

    Public Shared Function Read(ByVal source As Stream) As T
        Dim parser As New XmlSerializer(GetType(T))
        Dim result As Object

        Try
            result = parser.Deserialize(source)
        Catch iox As InvalidOperationException
            Return Nothing
        End Try

        Return DirectCast(result, T)
    End Function

    Public Shared Function Write(ByVal destination As Stream, ByVal data As T) As Boolean
        Dim parser As New XmlSerializer(GetType(T))

        Try
            parser.Serialize(destination, data)
        Catch iox As InvalidOperationException
            Return False
        End Try

        Return True
    End Function

    Public Shared Function IsCurrent(ByVal source As String, ByVal updateInterval As Integer) As Boolean
        If File.Exists(source) Then
            Dim writeDate As Date = New FileInfo(source).LastWriteTime
            Dim elapsed As TimeSpan = Date.Now - writeDate

            Return elapsed.TotalMinutes < CDbl(updateInterval)
        End If

        Return False
    End Function

End Class