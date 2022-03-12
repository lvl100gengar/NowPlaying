Imports System

Namespace Helpers

    Public Class ObjectHandle
        Implements IDisposable

        Dim theHandle As IntPtr
        Private disposedValue As Boolean

        Public ReadOnly Property Handle As IntPtr
            Get
                Return Me.theHandle
            End Get
        End Property

        Public ReadOnly Property IsValid As Boolean
            Get
                Return Not Me.Handle.Equals(IntPtr.Zero)
            End Get
        End Property

        Public Sub New(handle As IntPtr)
            Me.theHandle = handle
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                Kernel32.CloseHandle(Me.Handle)
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

End Namespace