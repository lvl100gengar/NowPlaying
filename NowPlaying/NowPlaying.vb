Public Class NowPlaying
    Implements IDisposable

    Private mediaInfo As IMetaProvider
    Private intervalInteger As Integer
    Private updateTimer As Timers.Timer
    Private currentMedia As String
    Private currentState As ProviderState
    Private disposedValue As Boolean

    'Thread safety attempt
    Private isFetchingState As Boolean
    Private newProvider As IMetaProvider

    Public Event StateChanged(sender As Object, e As ProviderStateEventArgs)

    Public ReadOnly Property Provider As IMetaProvider
        Get
            Return Me.mediaInfo
        End Get
    End Property

    Public Property UpdateInterval As Integer
        Get
            Return Me.intervalInteger
        End Get
        Set(ByVal value As Integer)
            Me.intervalInteger = value
            Me.updateTimer.Interval = value * 1000
        End Set
    End Property

    Public Sub ChangePlayer(provider As IMetaProvider)
        If Not Me.mediaInfo Is provider Then
            'Prevent updates temporarily.
            updateTimer.Stop()

            'Queue the new player interface for use.
            Me.newProvider = provider

            If provider Is Nothing Then
                'Raise the StateChanged event with a Disabled provider state and prevent the timer from starting.
                RaiseEvent StateChanged(Me, New ProviderStateEventArgs(ProviderState.Disabled, False))
                Exit Sub
            End If

            'Re-enable updates.
            updateTimer.Start()
        End If
    End Sub

    Public Sub New()
        Me.New(3)
    End Sub

    Public Sub New(ByVal updateInterval As Integer)
        Me.currentMedia = String.Empty
        Me.updateTimer = New Timers.Timer()
        Me.UpdateInterval = updateInterval

        AddHandler updateTimer.Elapsed, AddressOf OnUpdateTimerElapsed
    End Sub

    Private Sub OnUpdateTimerElapsed(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs)
        'Make sure another thread is not currently getting the providers state.
        If isFetchingState Then
            Exit Sub
        End If

        isFetchingState = True

        If Not newProvider Is Nothing Then
            mediaInfo = newProvider
            newProvider = Nothing
            currentState = ProviderState.Unavailable

            'Raise the StateChanged event with an Initializing provider state.
            RaiseEvent StateChanged(Me, New ProviderStateEventArgs(ProviderState.Searching, False))
        End If

        'Store the previous state in local variables to compare against.
        Dim previousState As ProviderState = Me.currentState
        Dim previousMedia As String = Me.currentMedia

        'Update the global variables to reflect the current state.
        Me.currentState = Me.mediaInfo.GetState()

        'See if the player has just been opened.
        If previousState = ProviderState.Unavailable AndAlso currentState <> ProviderState.Unavailable Then
            RaiseEvent StateChanged(Me, New ProviderStateEventArgs(ProviderState.Connecting, False))
        End If

        'See if the player has just been closed.
        'If previousState <> ProviderState.Unavailable AndAlso currentState = ProviderState.Unavailable Then
        '    RaiseEvent StateChanged(Me, New StateEventArgs(ProviderState.Unavailable, False))
        'End If

        'Don't try fetching meta if the state is unavailable.
        If currentState = ProviderState.Unavailable OrElse currentState = ProviderState.Stopped Then
            If Not currentState.Equals(previousState) Then
                RaiseEvent StateChanged(Me, New ProviderStateEventArgs(currentState, False))
            End If

            isFetchingState = False
            Exit Sub
        End If

        Me.currentMedia = Me.BuildCurrentMediaID()

        If Not currentMedia.Equals(previousMedia) OrElse Not currentState.Equals(previousState) Then
            RaiseEvent StateChanged(Me, New ProviderStateEventArgs(currentState, Not currentMedia.Equals(previousMedia)))
        End If

        isFetchingState = False
    End Sub

    Private Function BuildCurrentMediaID() As String
        If Me.Provider Is Nothing Then
            Return String.Empty
        Else
            Return Me.Provider.GenerateMeta(MetaField.Album) + Me.Provider.GenerateMeta(MetaField.Artist) + Me.Provider.GenerateMeta(MetaField.Title) + Me.Provider.GenerateMeta(MetaField.RadioTitle)
        End If
    End Function

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                Me.updateTimer.Dispose()
                Me.Provider.Dispose()
            End If
        End If

        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Sub ForceUpdate()
        Dim currentState As ProviderState = Me.mediaInfo.GetState()

        If currentState <> ProviderState.Stopped OrElse currentState <> ProviderState.Unavailable Then
            RaiseEvent StateChanged(Me, New ProviderStateEventArgs(currentState, True))
        End If
    End Sub

End Class

Public Class ProviderStateEventArgs

    Private providerState As ProviderState
    Private newMedia As Boolean

    Public ReadOnly Property State As ProviderState
        Get
            Return Me.providerState
        End Get
    End Property

    Public ReadOnly Property IsNewMedia As Boolean
        Get
            Return Me.newMedia
        End Get
    End Property

    Public Sub New(state As ProviderState, isNewMedia As Boolean)
        Me.providerState = state
        Me.newMedia = isNewMedia
    End Sub

End Class