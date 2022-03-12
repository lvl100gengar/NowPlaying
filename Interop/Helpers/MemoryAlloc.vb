Imports System
Imports System.Runtime.InteropServices

Namespace Helpers

    Public Class MemoryAlloc
        Implements IDisposable

        Private processHandle As IntPtr
        Private memoryAddress As IntPtr
        Private tempLocalHandle As IntPtr
        Private memorySize As UInteger
        Private disposedValue As Boolean

        Public ReadOnly Property Process As IntPtr
            Get
                Return Me.processHandle
            End Get
        End Property

        Public ReadOnly Property Address As IntPtr
            Get
                Return Me.memoryAddress
            End Get
        End Property

        Public ReadOnly Property Size As UInteger
            Get
                Return Me.memorySize
            End Get
        End Property

        Public Sub New(processID As IntPtr, address As IntPtr, size As Integer)
            Me.new(processID, address, size, Constants.AllocationType.MEM_COMMIT, Constants.MemoryProtection.PAGE_READWRITE)
        End Sub

        Public Sub New(processID As IntPtr, address As IntPtr, size As Integer, type As Constants.AllocationType, protection As Constants.MemoryProtection)
            Me.processHandle = Kernel32.OpenProcess(2035711, False, processID)
            Me.memoryAddress = Kernel32.VirtualAllocEx(Me.processHandle, address, size, type, protection)
            Me.memorySize = size
        End Sub

        Public Function Read(destinationAddress As IntPtr, destinationLength As Integer, ByRef bytesRead As Integer) As Boolean
            Return Kernel32.WriteProcessMemory(Me.Process, Me.Address, destinationAddress, destinationLength, bytesRead)
        End Function

        Public Function Write(sourceAddress As IntPtr, sourceLength As Integer, ByRef bytesWritten As Integer) As Boolean
            Return Kernel32.WriteProcessMemory(Me.Process, Me.Address, sourceAddress, sourceLength, bytesWritten)
        End Function

        Public Function Write(source As String, asUnicode As Boolean, ByRef bytesWritten As Integer) As Boolean
            If Not asUnicode Then
                Me.tempLocalHandle = Marshal.StringToHGlobalUni(source)
            Else
                Me.tempLocalHandle = Marshal.StringToHGlobalAnsi(source)
            End If

            Return Me.Write(Me.tempLocalHandle, source.Length, bytesWritten)
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                Kernel32.VirtualFreeEx(Me.Process, Me.Address, UIntPtr.Zero, Constants.FreeType.MEM_RELEASE)
                Kernel32.CloseHandle(Me.Process)
                Marshal.FreeHGlobal(Me.tempLocalHandle)
            End If

            Me.disposedValue = True
        End Sub

        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(False)
            MyBase.Finalize()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

    End Class

End Namespace