''' <summary>
''' Possible states for an IMetaProvider object.
''' </summary>
''' <remarks></remarks>
Public Enum ProviderState

    ''' <summary>
    ''' The provider is playing an audio file.
    ''' </summary>
    ''' <remarks></remarks>
    PlayingAudio

    ''' <summary>
    ''' The provider is playing a stream. The stream may contain audio and/or video.
    ''' </summary>
    ''' <remarks></remarks>
    PlayingStream

    ''' <summary>
    ''' The provider is playing a video file. The video may be accompanied by audio.
    ''' </summary>
    ''' <remarks></remarks>
    PlayingVideo

    ''' <summary>
    ''' The provider is paused. Meta is not guaranteed to be present.
    ''' </summary>
    ''' <remarks></remarks>
    PausedAudio

    PausedStream

    PausedVideo

    ''' <summary>
    ''' The provider is open but not playing anything, or has no information.
    ''' </summary>
    ''' <remarks></remarks>
    Stopped

    ''' <summary>
    ''' The provider is not running or an error occoured setting up a remote connection.
    ''' </summary>
    ''' <remarks></remarks>
    Unavailable

    Connecting

    Searching

    Disabled

End Enum