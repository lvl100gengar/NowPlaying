''' <summary>
''' The list of metadata fields that can be pulled from NowPlaying media sources.
''' </summary>
''' <remarks></remarks>
Public Enum MetaField

    ''' <summary>
    ''' The title of the current album.
    ''' </summary>
    ''' <remarks></remarks>
    Album

    ''' <summary>
    ''' The artist(s) of the current album.
    ''' </summary>
    ''' <remarks></remarks>
    AlbumArtist

    ''' <summary>
    ''' The number of tracks on the current album. PlaylistLength may be returned if this value is not available.
    ''' </summary>
    ''' <remarks></remarks>
    AlbumLength

    ''' <summary>
    ''' The position of the current track on the current album. PlaylistPosition may be returned if this value is not available.
    ''' </summary>
    ''' <remarks></remarks>
    AlbumPosition

    ''' <summary>
    ''' The artist(s) of the current track.
    ''' </summary>
    ''' <remarks></remarks>
    Artist

    ''' <summary>
    ''' The bitrate of the current track in bytes.
    ''' </summary>
    ''' <remarks></remarks>
    Bitrate

    ''' <summary>
    ''' The number of channels available in the current audio track.
    ''' </summary>
    ''' <remarks></remarks>
    Channels

    ''' <summary>
    ''' The file extension of the current track.
    ''' </summary>
    ''' <remarks></remarks>
    FIleExtension

    ''' <summary>
    ''' The file name (not including path or extension) of the current track.
    ''' </summary>
    ''' <remarks></remarks>
    FileName

    ''' <summary>
    ''' The absolute path of the current track.
    ''' </summary>
    ''' <remarks></remarks>
    FilePath

    ''' <summary>
    ''' The file size of the current track in bytes.
    ''' </summary>
    ''' <remarks></remarks>
    FileSize

    ''' <summary>
    ''' The genre of the current track.
    ''' </summary>
    ''' <remarks></remarks>
    Genre

    ''' <summary>
    ''' The length of the current playlist.
    ''' </summary>
    ''' <remarks></remarks>
    PlaylistLength

    ''' <summary>
    ''' The position of the current track within the current playlist.
    ''' </summary>
    ''' <remarks></remarks>
    PlaylistPosition

    ''' <summary>
    ''' The number of times this track has been played by this media player.
    ''' </summary>
    ''' <remarks></remarks>
    PlayCount

    ''' <summary>
    ''' When a stream is playing, RadioTitle is the title reported by the server.
    ''' </summary>
    ''' <remarks></remarks>
    RadioTitle

    ''' <summary>
    ''' The rating of the current track as a percentage.
    ''' </summary>
    ''' <remarks></remarks>
    Rating

    ''' <summary>
    ''' The sample rate of the current audio track in hertz.
    ''' </summary>
    ''' <remarks></remarks>
    SampleRate

    ''' <summary>
    ''' The title of the current track. FileName may be returned if the title is not available.
    ''' </summary>
    ''' <remarks></remarks>
    Title

    ''' <summary>
    ''' The length of the current track in seconds.
    ''' </summary>
    ''' <remarks></remarks>
    TrackLength

    ''' <summary>
    ''' The position of playback on the current track in seconds.
    ''' </summary>
    ''' <remarks></remarks>
    TrackPosition

    ''' <summary>
    ''' The height in pixels of the current video. Nothing is returned when no video is playing or for players the do no support video playback.
    ''' </summary>
    ''' <remarks></remarks>
    VideoHeight

    ''' <summary>
    ''' The width in pixels of the current video. Nothing is returned when no video is playing or for players the do no support video playback.
    ''' </summary>
    ''' <remarks></remarks>
    VideoWidth

    ''' <summary>
    ''' The year the current track was released. Always expressed using four digits.
    ''' </summary>
    ''' <remarks></remarks>
    Year

End Enum