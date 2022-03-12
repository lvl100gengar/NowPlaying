Imports System
Imports System.Runtime.InteropServices

''' <summary>
''' Common functions found in the kernel32 library.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Kernel32

    ''' <summary>
    ''' Closes an open object handle.
    ''' </summary>
    ''' <param name="hObject">A valid handle to an open object.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error information, call GetLastError. If the application is running under a debugger, the function will throw an exception if it receives either a handle value that is not valid or a pseudo-handle value. This can happen if you close a handle twice, or if you call CloseHandle on a handle returned by the FindFirstFile function instead of calling the FindClose function.</returns>
    ''' <remarks></remarks>
    <DllImport("kernel32", CallingConvention:=CallingConvention.Winapi, ExactSpelling:=True, SetLastError:=True)> _
    Public Shared Function CloseHandle(<InAttribute()> hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function

    ''' <summary>
    ''' Opens an existing local process object.
    ''' </summary>
    ''' <param name="dwDesiredAccess">The access to the process object. This access right is checked against the security descriptor for the process.</param>
    ''' <param name="bInheritHandle">If this value is TRUE, processes created by this process will inherit the handle. Otherwise, the processes do not inherit this handle.</param>
    ''' <param name="dwProcessId">The identifier of the local process to be opened.</param>
    ''' <returns>If the function succeeds, the return value is an open handle to the specified process. If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
    ''' <remarks>To open a handle to another local process and obtain full access rights, you must enable the SeDebugPrivilege privilege. For more information, see Changing Privileges in a Token. The handle returned by the OpenProcess function can be used in any function that requires a handle to a process, such as the wait functions, provided the appropriate access rights were requested. When you are finished with the handle, be sure to close it using the CloseHandle function.</remarks>
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function OpenProcess(dwDesiredAccess As Constants.ProcessAcessRights, bInheritHandle As Boolean, dwProcessId As Integer) As IntPtr

    End Function

    ''' <summary>
    ''' Reads data from an area of memory in a specified process. The entire area to be read must be accessible or the operation fails.
    ''' </summary>
    ''' <param name="hProcess">A handle to the process with memory that is being read. The handle must have PROCESS_VM_READ access to the process.</param>
    ''' <param name="lpBaseAddress">A pointer to the base address in the specified process from which to read. Before any data transfer occurs, the system verifies that all data in the base address and memory of the specified size is accessible for read access, and if it is not accessible the function fails.</param>
    ''' <param name="lpBuffer">A pointer to a buffer that receives the contents from the address space of the specified process.</param>
    ''' <param name="nSize">The number of bytes to be read from the specified process.</param>
    ''' <param name="lpNumberOfBytesRead">A pointer to a variable that receives the number of bytes transferred into the specified buffer. If lpNumberOfBytesRead is NULL, the parameter is ignored.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is 0 (zero). To get extended error information, call GetLastError. The function fails if the requested read operation crosses into an area of the process that is inaccessible.</returns>
    ''' <remarks>ReadProcessMemory copies the data in the specified address range from the address space of the specified process into the specified buffer of the current process. Any process that has a handle with PROCESS_VM_READ access can call the function.</remarks>
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function ReadProcessMemory(hProcess As IntPtr, lpBaseAddress As IntPtr, <Out()> lpBuffer As IntPtr, nSize As UInt32, <Out()> ByRef lpNumberOfBytesRead As UInt32) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function

    ''' <summary>
    ''' Reserves or commits a region of memory within the virtual address space of a specified process. The function initializes the memory it allocates to zero, unless MEM_RESET is used. To specify the NUMA node for the physical memory, see VirtualAllocExNuma.
    ''' </summary>
    ''' <param name="hProcess">The handle to a process. The function allocates memory within the virtual address space of this process. The handle must have the PROCESS_VM_OPERATION access right.</param>
    ''' <param name="lpAddress">The pointer that specifies a desired starting address for the region of pages that you want to allocate. If you are reserving memory, the function rounds this address down to the nearest multiple of the allocation granularity. If you are committing memory that is already reserved, the function rounds this address down to the nearest page boundary. To determine the size of a page and the allocation granularity on the host computer, use the GetSystemInfo function. If lpAddress is NULL, the function determines where to allocate the region.</param>
    ''' <param name="dwSize">The size of the region of memory to allocate, in bytes. If lpAddress is NULL, the function rounds dwSize up to the next page boundary. If lpAddress is not NULL, the function allocates all pages that contain one or more bytes in the range from lpAddress to lpAddress+dwSize. This means, for example, that a 2-byte range that straddles a page boundary causes the function to allocate both pages.</param>
    ''' <param name="flAllocationType">The type of memory allocation.</param>
    ''' <param name="flProtect">The memory protection for the region of pages to be allocated. If the pages are being committed, you can specify any one of the memory protection constants.</param>
    ''' <returns>If the function succeeds, the return value is the base address of the allocated region of pages. If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
    ''' <remarks></remarks>
    <DllImport("kernel32.dll", SetLastError:=True, ExactSpelling:=True)> _
    Public Shared Function VirtualAllocEx(hProcess As IntPtr, lpAddress As IntPtr, <MarshalAs(UnmanagedType.U4)> dwSize As UInt32, <MarshalAs(UnmanagedType.U4)> flAllocationType As Constants.AllocationType, <MarshalAs(UnmanagedType.U4)> flProtect As Constants.MemoryProtection) As IntPtr

    End Function

    ''' <summary>
    ''' Releases, decommits, or releases and decommits a region of memory within the virtual address space of a specified process.
    ''' </summary>
    ''' <param name="hProcess">A handle to a process. The function frees memory within the virtual address space of the process. The handle must have the PROCESS_VM_OPERATION access right.</param>
    ''' <param name="lpAddress">A pointer to the starting address of the region of memory to be freed. If the dwFreeType parameter is MEM_RELEASE, lpAddress must be the base address returned by the VirtualAllocEx function when the region is reserved.</param>
    ''' <param name="dwSize">The size of the region of memory to free, in bytes. If the dwFreeType parameter is MEM_RELEASE, dwSize must be 0 (zero). The function frees the entire region that is reserved in the initial allocation call to VirtualAllocEx. If dwFreeType is MEM_DECOMMIT, the function decommits all memory pages that contain one or more bytes in the range from the lpAddress parameter to (lpAddress+dwSize). This means, for example, that a 2-byte region of memory that straddles a page boundary causes both pages to be decommitted. If lpAddress is the base address returned by VirtualAllocEx and dwSize is 0 (zero), the function decommits the entire region that is allocated by VirtualAllocEx. After that, the entire region is in the reserved state.</param>
    ''' <param name="dwFreeType">The type of free operation.</param>
    ''' <returns>If the function succeeds, the return value is a nonzero value. If the function fails, the return value is 0 (zero). To get extended error information, call GetLastError.</returns>
    ''' <remarks>Each page of memory in a process virtual address space has a Page State. The VirtualFreeEx function can decommit a range of pages that are in different states, some committed and some uncommitted. This means that you can decommit a range of pages without first determining the current commitment state of each page. Decommitting a page releases its physical storage, either in memory or in the paging file on disk. If a page is decommitted but not released, its state changes to reserved. Subsequently, you can call VirtualAllocEx to commit it, or VirtualFreeEx to release it. Attempting to read from or write to a reserved page results in an access violation exception. The VirtualFreeEx function can release a range of pages that are in different states, some reserved and some committed. This means that you can release a range of pages without first determining the current commitment state of each page. The entire range of pages originally reserved by VirtualAllocEx must be released at the same time. If a page is released, its state changes to free, and it is available for subsequent allocation operations. After memory is released or decommitted, you can never refer to the memory again. Any information that may have been in that memory is gone forever. Attempts to read from or write to a free page results in an access violation exception. If you need to keep information, do not decommit or free memory that contains the information. The VirtualFreeEx function can be used on an AWE region of memory and it invalidates any physical page mappings in the region when freeing the address space. However, the physical pages are not deleted, and the application can use them. The application must explicitly call FreeUserPhysicalPages to free the physical pages. When the process is terminated, all resources are automatically cleaned up.</remarks>
    <DllImport("kernel32.dll", SetLastError:=True, ExactSpelling:=True)> _
    Public Shared Function VirtualFreeEx(hProcess As IntPtr, lpAddress As IntPtr, dwSize As UIntPtr, dwFreeType As Constants.FreeType) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function

    ''' <summary>
    ''' Writes data to an area of memory in a specified process. The entire area to be written to must be accessible or the operation fails.
    ''' </summary>
    ''' <param name="hProcess">A handle to the process memory to be modified. The handle must have PROCESS_VM_WRITE and PROCESS_VM_OPERATION access to the process.</param>
    ''' <param name="lpBaseAddress">A pointer to the base address in the specified process to which data is written. Before data transfer occurs, the system verifies that all data in the base address and memory of the specified size is accessible for write access, and if it is not accessible, the function fails.</param>
    ''' <param name="lpBuffer">A pointer to the buffer that contains data to be written in the address space of the specified process.</param>
    ''' <param name="nSize">The number of bytes to be written to the specified process.</param>
    ''' <param name="lpNumberOfBytesWritten">A pointer to a variable that receives the number of bytes transferred into the specified process. This parameter is optional. If lpNumberOfBytesWritten is NULL, the parameter is ignored.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is 0 (zero). To get extended error information, call GetLastError. The function fails if the requested write operation crosses into an area of the process that is inaccessible.</returns>
    ''' <remarks>WriteProcessMemory copies the data from the specified buffer in the current process to the address range of the specified process. Any process that has a handle with PROCESS_VM_WRITE and PROCESS_VM_OPERATION access to the process to be written to can call the function. Typically but not always, the process with address space that is being written to is being debugged.</remarks>
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function WriteProcessMemory(hProcess As IntPtr, lpBaseAddress As IntPtr, lpBuffer As IntPtr, nSize As Int32, <Out()> ByRef lpNumberOfBytesWritten As UInt32) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function



End Class