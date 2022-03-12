Imports iTunesLib
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.IO

Public NotInheritable Class ITunes
    Implements IMetaProvider

    Private base As iTunesApp
    Private Shared myInstance As New ITunes()

    Private Sub New()

    End Sub

    Public Shared ReadOnly Property Instance As ITunes
        Get
            Return ITunes.myInstance
        End Get
    End Property

    Public Function GenerateMeta(field As MetaField) As String Implements IMetaProvider.GenerateMeta
        On Error Resume Next

        With base.CurrentTrack

            Select Case field

                Case MetaField.Album
                    Return .Album

                Case MetaField.AlbumArtist
                    Return .Artist

                Case MetaField.AlbumLength
                    Return .TrackCount.ToString()

                Case MetaField.AlbumPosition
                    Return .TrackNumber.ToString()

                Case MetaField.Artist
                    Return .Artist

                Case MetaField.Bitrate
                    Return .BitRate.ToString()

                Case MetaField.Channels

                Case MetaField.FIleExtension
                    Return Path.GetExtension(base.CurrentStreamURL)

                Case MetaField.FileName
                    Return Path.GetFileNameWithoutExtension(base.CurrentStreamURL)

                Case MetaField.FilePath
                    Return base.CurrentStreamURL

                Case MetaField.FileSize
                    Return New FileInfo(base.CurrentStreamURL).Length.ToString()

                Case MetaField.Genre
                    Return .Genre

                Case MetaField.PlayCount
                    Return .PlayedCount.ToString()

                Case MetaField.PlaylistLength
                    Return .Playlist.Tracks.Count.ToString()

                Case MetaField.PlaylistPosition


                Case MetaField.Rating
                    Return .Rating

                Case MetaField.SampleRate
                    Return .SampleRate.ToString()

                Case MetaField.Title
                    Return .Name

                Case MetaField.TrackLength
                    Return .Duration.ToString()

                Case MetaField.TrackPosition
                    Return .Time

                Case MetaField.VideoHeight


                Case MetaField.VideoWidth


                Case MetaField.Year
                    Return .Year.ToString()

            End Select

        End With

        Return "N/A"
    End Function

    Public Function GetState() As ProviderState Implements IMetaProvider.GetState
        If base Is Nothing Then
            If Process.GetProcessesByName("iTunes").Length = 0 Then
                Return ProviderState.Unavailable
            End If

            Try
                base = New iTunesApp()
            Catch cx As COMException
                Return ProviderState.Unavailable
            End Try
        End If

        Try
            If base.CurrentTrack Is Nothing Then
                Return ProviderState.Stopped
            Else
                If Not String.IsNullOrEmpty(base.CurrentStreamURL) Then
                    Return ProviderState.PlayingStream
                Else
                    Return ProviderState.PlayingAudio
                End If
            End If
        Catch icx As InvalidCastException
            base = Nothing
            Return ProviderState.Unavailable
        Catch cx As COMException
            'This happens when the user opens a dialog in iTunes.
            'Just report the player as available, but not playing.
            If cx.ErrorCode = -2147417846 Then
                Return ProviderState.Stopped
            Else
                base = Nothing
                Return ProviderState.Unavailable
            End If
        End Try
    End Function

    Public Function SupportsMeta(field As MetaField) As Boolean Implements IMetaProvider.SupportsMeta
        Return True
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
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

End Class