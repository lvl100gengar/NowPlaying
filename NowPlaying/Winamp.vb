Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports Interop
Imports Interop.Helpers


Public Class Winamp
    Implements IMetaProvider

    Private windowHandle As IntPtr
    Private hProcess As IntPtr
    Private processID As Integer
    Private disposedValue As Boolean
    Private Shared myInstance As New Winamp()

    Public Shared ReadOnly Property Instance As Winamp
        Get
            Return Winamp.myInstance
        End Get
    End Property

    Private ReadOnly Property CurrentFile() As String
        Get
            Return ReadRemoteString(SendUser(SendUser(0, WinampMessage.GetListPosition).ToInt32(), WinampMessage.GetFilePath), 260)
        End Get
    End Property

    Public ReadOnly Property PlayerTitle As String
        Get
            Return User32.GetWindowText(Me.windowHandle)
        End Get
    End Property

    Private Sub New()

    End Sub

    Private Function Bind() As Boolean
        'Find a window that is using the Winamp v1.x class.
        Me.windowHandle = User32.FindWindow("Winamp v1.x", Nothing)

        'Make sure FindWindow gave us a valid handle.
        If Me.windowHandle.Equals(IntPtr.Zero) Then
            Return False
        End If

        User32.GetWindowThreadProcessId(Me.windowHandle, Me.processID)

        If Me.processID = 0 Then
            Return False
        End If

        'Open the process that created the window.
        Me.hProcess = Kernel32.OpenProcess(2035711, False, processID)

        'Make sure OpenProcess gave us a valid handle.
        If Me.hProcess.Equals(IntPtr.Zero) Then
            Return False
        End If

        Return True
    End Function

    Private Function GetMetaData(ByVal metaName As String) As String
        If Me.hProcess.Equals(IntPtr.Zero) Then
            Return "NotBound"
        End If

        Dim baseAddress As IntPtr = Kernel32.VirtualAllocEx(Me.hProcess, Nothing, 4096UI, Constants.AllocationType.MEM_COMMIT, Constants.MemoryProtection.PAGE_READWRITE)

        If baseAddress.Equals(IntPtr.Zero) Then
            Return "VirtAllocErr"
        End If

        Dim info As New WinampMetadata

        info.filename = IntPtr.Add(baseAddress, 1024)
        info.metadata = IntPtr.Add(baseAddress, 2048)
        info.ret = IntPtr.Add(baseAddress, 3072)
        info.retlen = 1024

        Dim fileNameSize As Integer = Me.CurrentFile.Length + 1
        Dim fileNameAddress As IntPtr = Marshal.StringToHGlobalAnsi(Me.CurrentFile)
        Dim metaNameSize As Integer = metaName.Length + 1
        Dim metaNameAddress As IntPtr = Marshal.StringToHGlobalAnsi(metaName)
        Dim structureSize As Integer = Marshal.SizeOf(info)
        Dim structureAddress As IntPtr = Marshal.AllocHGlobal(structureSize)

        Marshal.StructureToPtr(info, structureAddress, False)

        Kernel32.WriteProcessMemory(Me.hProcess, baseAddress, structureAddress, structureSize, Nothing)
        Kernel32.WriteProcessMemory(Me.hProcess, info.filename, fileNameAddress, fileNameSize, Nothing)
        Kernel32.WriteProcessMemory(Me.hProcess, info.metadata, metaNameAddress, metaNameSize, Nothing)

        Dim tempResult As String = String.Empty

        If Me.SendUser(baseAddress.ToInt64(), WinampMessage.GetExtendedInfo).Equals(New IntPtr(1)) Then
            Dim returnBuffer As IntPtr = Marshal.AllocHGlobal(1024)
            Dim returnSize As UInteger

            If Kernel32.ReadProcessMemory(Me.hProcess, info.ret, returnBuffer, 1024, returnSize) Then
                tempResult = Marshal.PtrToStringAnsi(returnBuffer, returnSize)
                tempResult = tempResult.TrimEnd(Convert.ToChar(0))
            End If
        End If

        Kernel32.VirtualFreeEx(Me.hProcess, baseAddress, UIntPtr.Zero, Constants.FreeType.MEM_RELEASE)
        Marshal.FreeHGlobal(fileNameAddress)
        Marshal.FreeHGlobal(metaNameAddress)
        Marshal.FreeHGlobal(structureAddress)

        Return tempResult
    End Function

    Private Function ReadRemoteString(ByVal remoteBuf As IntPtr, ByVal maxLength As Integer) As String
        If remoteBuf.Equals(IntPtr.Zero) OrElse Me.hProcess.Equals(IntPtr.Zero) Then
            Return String.Empty
        End If

        'allocate a local buffer which will hold the returned string
        Dim returnVal As IntPtr = Marshal.AllocHGlobal(maxLength)

        'read the remote memory to local buffer 
        Kernel32.ReadProcessMemory(Me.hProcess, remoteBuf, returnVal, maxLength, Nothing)

        Dim value As String = Marshal.PtrToStringAnsi(returnVal)

        Marshal.FreeHGlobal(returnVal)

        Return value
    End Function

    Private Function SendUser(ByVal param As Int64, ByVal message As WinampMessage) As IntPtr
        Return User32.SendMessage(Me.windowHandle, 1024, New IntPtr(param), New IntPtr(message))
    End Function

    Public Function GenerateMeta(field As MetaField) As String Implements IMetaProvider.GenerateMeta
        On Error Resume Next

        Select Case field

            Case MetaField.Album
                Return GetMetaData("album")

            Case MetaField.AlbumArtist
                Return GetMetaData("albumartist")

            Case MetaField.AlbumLength
                Return GetMetaData("totaltracks")

            Case MetaField.AlbumPosition
                Return GetMetaData("tracknumber")

            Case MetaField.Artist
                Return GetMetaData("artist")

            Case MetaField.Bitrate
                Dim temp As Long = SendUser(WinampInfoField.Bitrate, WinampMessage.GetInfo).ToInt64() * 1000L 'kbps

                If temp = 0L Then
                    Return CLng(GetMetaData("bitrate")) * 1000L 'kbps
                Else
                    Return temp
                End If

            Case MetaField.Channels
                Return SendUser(WinampInfoField.Channels, WinampMessage.GetInfo).ToString()

            Case MetaField.FIleExtension
                Return Path.GetExtension(Me.CurrentFile)

            Case MetaField.FileName
                Return Path.GetFileNameWithoutExtension(Me.CurrentFile)

            Case MetaField.FilePath
                Return Me.CurrentFile

            Case MetaField.FileSize
                Return New FileInfo(Me.CurrentFile).Length.ToString()

            Case MetaField.Genre
                Return GetMetaData("genre")

            Case MetaField.PlayCount
                Return GetMetaData("playcount")

            Case MetaField.PlaylistLength
                Return SendUser(0, WinampMessage.GetListLength).ToString()

            Case MetaField.PlaylistPosition
                Return (SendUser(0, WinampMessage.GetListPosition).ToInt32() + 1).ToString()

            Case MetaField.RadioTitle
                Return ReadRemoteString(SendUser(SendUser(0, WinampMessage.GetListPosition).ToInt32(), WinampMessage.GetFileTitle), 260)

            Case MetaField.Rating
                Return SendUser(0, WinampMessage.GetRating).ToString()

            Case MetaField.SampleRate
                Return SendUser(WinampInfoField.SampleRateHertz, WinampMessage.GetInfo).ToString()

            Case MetaField.Title
                Return GetMetaData("title")

            Case MetaField.TrackLength
                Return SendUser(WinampTimeMode.Length, WinampMessage.GetOutputTime).ToString()

            Case MetaField.TrackPosition
                Return (SendUser(WinampTimeMode.Position, WinampMessage.GetOutputTime).ToInt64() \ 1000L).ToString()

            Case MetaField.VideoHeight


            Case MetaField.VideoWidth


            Case MetaField.Year
                Return GetMetaData("year")

        End Select

        Return "N/A"
    End Function

    Public Function GetState() As ProviderState Implements IMetaProvider.GetState
        If Not User32.GetClassName(Me.windowHandle).Equals("Winamp v1.x") Then
            If Me.disposedValue OrElse Not Me.Bind() Then
                Return ProviderState.Unavailable
            End If
        End If

        Select Case Me.SendUser(0L, WinampMessage.IsPlaying).ToString()
            Case "1" 'Playing
                Return GetMediaState(True)

            Case "3" 'Paused
                Return GetMediaState(False)
        End Select

        Return ProviderState.Stopped
    End Function

    Private Function GetMediaState(isPlaying As Boolean) As ProviderState
        If String.IsNullOrEmpty(Me.CurrentFile) Then
            Return ProviderState.Stopped
        Else
            If SendUser(WinampTimeMode.Length, WinampMessage.GetOutputTime).ToInt32() = -1 Then
                Return If(isPlaying, ProviderState.PlayingStream, ProviderState.PausedStream)
            Else
                If ReadRemoteString(SendUser(0, WinampMessage.IsVideo), 10).Equals("1") Then
                    Return If(isPlaying, ProviderState.PlayingVideo, ProviderState.PausedVideo)
                Else
                    Return If(isPlaying, ProviderState.PlayingAudio, ProviderState.PausedAudio)
                End If
            End If
        End If
    End Function

    Public Function SupportsMeta(field As MetaField) As Boolean Implements IMetaProvider.SupportsMeta
        Return True
    End Function

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            Kernel32.CloseHandle(hProcess)
        End If

        Me.disposedValue = True
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

End Class

<StructLayout(LayoutKind.Sequential)> _
Public Structure WinampMetadata

    Dim filename As IntPtr
    Dim metadata As IntPtr
    Dim ret As IntPtr
    Dim retlen As Int32

End Structure

Public Enum WinampMessage As Integer

    GetVersion = 0
    IsPlaying = 104         '0=Stopped 1=Playing 3=Paused [Returns]
    GetOutputTime = 105     '0=Position (ms) 1=Length (sec) [modes]
    GetListLength = 124
    GetListPosition = 125
    GetInfo = 126           'Info about the current media.
    GetFilePath = 211
    GetFileTitle = 212
    RefreshTitleCache = 247 'Reloads all playlist titles.
    GetBasicInfo = 291      'File Title and Length.
    GetExtendedInfo = 290   'File metadata.
    GetExtendedInfoHookable = 296   'File metadata.
    IsVideo = 501
    GetOutputPlugin = 625   'Name of the output plugin.
    GetRating = 640         'Rating 0-5

End Enum

Public Enum WinampTimeMode As Integer

    Position = 0 'Milliseconds
    Length = 1   'Seconds

End Enum

Public Enum WinampInfoField As Integer

    Samplerate = 0      'i.e. 44100
    Bitrate = 1         'i.e. 128
    Channels = 2        'i.e. 2
    VideoSize = 3 'LOWORD=w HIWORD=h
    VideoDescription = 4
    SampleRateHertz = 5

End Enum