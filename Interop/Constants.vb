Imports System

''' <summary>
''' Contains constants for use with many native functions.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Constants

    ''' <summary>
    ''' The type of memory allocation.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum AllocationType As Integer

        ''' <summary>
        ''' Allocates physical storage in memory or in the paging file on disk for the specified reserved memory pages. The function initializes the memory to zero. To reserve and commit pages in one step, call VirtualAllocEx with MEM_COMMIT | MEM_RESERVE. The function fails if you attempt to commit a page that has not been reserved. The resulting error code is ERROR_INVALID_ADDRESS. An attempt to commit a page that is already committed does not cause the function to fail. This means that you can commit pages without first determining the current commitment state of each page.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_COMMIT = 4096

        ''' <summary>
        ''' Reserves a range of the process's virtual address space without allocating any actual physical storage in memory or in the paging file on disk. You commit reserved pages by calling VirtualAllocEx again with MEM_COMMIT. To reserve and commit pages in one step, call VirtualAllocEx with MEM_COMMIT | MEM_RESERVE. Other memory allocation functions, such as malloc and LocalAlloc, cannot use reserved memory until it has been released.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_RESERVE = 8192

        ''' <summary>
        ''' Indicates that data in the memory range specified by lpAddress and dwSize is no longer of interest. The pages should not be read from or written to the paging file. However, the memory block will be used again later, so it should not be decommitted. This value cannot be used with any other value. Using this value does not guarantee that the range operated on with MEM_RESET will contain zeros. If you want the range to contain zeros, decommit the memory and then recommit it. When you use MEM_RESET, the VirtualAllocEx function ignores the value of fProtect. However, you must still set fProtect to a valid protection value, such as PAGE_NOACCESS. VirtualAllocEx returns an error if you use MEM_RESET and the range of memory is mapped to a file. A shared view is only acceptable if it is mapped to a paging file.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_RESET = 524288

        ''' <summary>
        ''' Allocates memory using large page support. The size and alignment must be a multiple of the large-page minimum. To obtain this value, use the GetLargePageMinimum function.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_LARGE_PAGES = 536870912

        ''' <summary>
        ''' Reserves an address range that can be used to map Address Windowing Extensions (AWE) pages. This value must be used with MEM_RESERVE and no other values.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_PHYSICAL = 4194304

        ''' <summary>
        ''' Allocates memory at the highest possible address. This can be slower than regular allocations, especially when there are many allocations.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_TOP_DOWN = 1048576

    End Enum

    ''' <summary>
    ''' Flags for the EnumDisplayDevice function.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DisplayDevicesFlags

        ''' <summary>
        ''' Retrieve all available attributes for the device.
        ''' </summary>
        ''' <remarks></remarks>
        EDD_GET_ALL_ATTRIBUTES = 0

        ''' <summary>
        '''  Retrieve the device interface name for GUID_DEVINTERFACE_MONITOR.
        ''' </summary>
        ''' <remarks></remarks>
        EDD_GET_DEVICE_INTERFACE_NAME = 1

    End Enum

    ''' <summary>
    ''' Device state flags.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags()> _
    Public Enum DisplayDeviceStateFlags As Integer

        ''' <summary>
        ''' DISPLAY_DEVICE_ACTIVE specifies whether a monitor is presented as being "on" by the respective GDI view. Windows Vista: EnumDisplayDevices will only enumerate monitors that can be presented as being "on."
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_ACTIVE = 1

        DISPLAY_DEVICE_MULTI_DRIVER = 2

        ''' <summary>
        ''' The primary desktop is on the device. For a system with a single display card, this is always set. For a system with multiple display cards, only one device can have this set.
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_PRIMARY_DEVICE = 4

        ''' <summary>
        ''' Represents a pseudo device used to mirror application drawing for remoting or other purposes. An invisible pseudo monitor is associated with this device. For example, NetMeeting uses it. Note that GetSystemMetrics (SM_MONITORS) only accounts for visible display monitors.
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_MIRRORING_DRIVER = 8

        ''' <summary>
        ''' The device is VGA compatible.
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_VGA_COMPATIBLE = 22

        ''' <summary>
        ''' The device is removable; it cannot be the primary display.
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_REMOVABLE = 32

        ''' <summary>
        ''' The device has more display modes than its output devices support.
        ''' </summary>
        ''' <remarks></remarks>
        DISPLAY_DEVICE_MODESPRUNED = 134217728

        DISPLAY_DEVICE_REMOTE = 61708864

        DISPLAY_DEVICE_DISCONNECT = 33554432

    End Enum

    ''' <summary>
    ''' The type of free operation.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FreeType As Integer

        ''' <summary>
        ''' Decommits the specified region of committed pages. After the operation, the pages are in the reserved state. The function does not fail if you attempt to decommit an uncommitted page. This means that you can decommit a range of pages without first determining their current commitment state. Do not use this value with MEM_RELEASE.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_DECOMMIT = 16384

        ''' <summary>
        ''' Releases the specified region of pages. After the operation, the pages are in the free state. If you specify this value, dwSize must be 0 (zero), and lpAddress must point to the base address returned by the VirtualAllocEx function when the region is reserved. The function fails if either of these conditions is not met. If any pages in the region are committed currently, the function first decommits, and then releases them. The function does not fail if you attempt to release pages that are in different states, some reserved and some committed. This means that you can release a range of pages without first determining the current commitment state. Do not use this value with MEM_DECOMMIT.
        ''' </summary>
        ''' <remarks></remarks>
        MEM_RELEASE = 32768

    End Enum

    ''' <summary>
    ''' The following are the memory-protection options; you must specify one of the following values when allocating or protecting a page in memory. Protection attributes cannot be assigned to a portion of a page; they can only be assigned to a whole page.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum MemoryProtection As Integer

        ''' <summary>
        ''' Enables execute access to the committed region of pages. An attempt to read from or write to the committed region results in an access violation. This flag is not supported by the CreateFileMapping function.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_EXECUTE = 16

        ''' <summary>
        ''' Enables execute or read-only access to the committed region of pages. An attempt to write to the committed region results in an access violation.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_EXECUTE_READ = 32

        ''' <summary>
        ''' Enables execute, read-only, or read/write access to the committed region of pages.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_EXECUTE_READWRITE = 64

        ''' <summary>
        ''' Enables execute, read-only, or copy-on-write access to a mapped view of a file mapping object. An attempt to write to a committed copy-on-write page results in a private copy of the page being made for the process. The private page is marked as PAGE_EXECUTE_READWRITE, and the change is written to the new page. This flag is not supported by the VirtualAlloc or VirtualAllocEx functions.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_EXECUTE_WRITECOPY = 128

        ''' <summary>
        ''' Disables all access to the committed region of pages. An attempt to read from, write to, or execute the committed region results in an access violation. This flag is not supported by the CreateFileMapping function.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_NOACCESS = 1

        ''' <summary>
        ''' Enables read-only access to the committed region of pages. An attempt to write to the committed region results in an access violation. If Data Execution Prevention is enabled, an attempt to execute code in the committed region results in an access violation.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_READONLY = 2

        ''' <summary>
        ''' Enables read-only or read/write access to the committed region of pages. If Data Execution Prevention is enabled, attempting to execute code in the committed region results in an access violation.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_READWRITE = 4

        ''' <summary>
        ''' Enables read-only or copy-on-write access to a mapped view of a file mapping object. An attempt to write to a committed copy-on-write page results in a private copy of the page being made for the process. The private page is marked as PAGE_READWRITE, and the change is written to the new page. If Data Execution Prevention is enabled, attempting to execute code in the committed region results in an access violation. This flag is not supported by the VirtualAlloc or VirtualAllocEx functions.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_WRITECOPY = 8

        ''' <summary>
        ''' Pages in the region become guard pages. Any attempt to access a guard page causes the system to raise a STATUS_GUARD_PAGE_VIOLATION exception and turn off the guard page status. Guard pages thus act as a one-time access alarm. For more information, see Creating Guard Pages. When an access attempt leads the system to turn off guard page status, the underlying page protection takes over. If a guard page exception occurs during a system service, the service typically returns a failure status indicator. This value cannot be used with PAGE_NOACCESS. This flag is not supported by the CreateFileMapping function.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_GUARD = 256

        ''' <summary>
        ''' Sets all pages to be non-cachable. Applications should not use this attribute except when explicitly required for a device. Using the interlocked functions with memory that is mapped with SEC_NOCACHE can result in an EXCEPTION_ILLEGAL_INSTRUCTION exception. The PAGE_NOCACHE flag cannot be used with the PAGE_GUARD, PAGE_NOACCESS, or PAGE_WRITECOMBINE flags. The PAGE_NOCACHE flag can be used only when allocating private memory with the VirtualAlloc, VirtualAllocEx, or VirtualAllocExNuma functions. To enable non-cached memory access for shared memory, specify the SEC_NOCACHE flag when calling the CreateFileMapping function.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_NOCACHE = 512

        ''' <summary>
        ''' Sets all pages to be write-combined. Applications should not use this attribute except when explicitly required for a device. Using the interlocked functions with memory that is mapped as write-combined can result in an EXCEPTION_ILLEGAL_INSTRUCTION exception. The PAGE_WRITECOMBINE flag cannot be specified with the PAGE_NOACCESS, PAGE_GUARD, and PAGE_NOCACHE flags. The PAGE_WRITECOMBINE flag can be used only when allocating private memory with the VirtualAlloc, VirtualAllocEx, or VirtualAllocExNuma functions. To enable write-combined memory access for shared memory, specify the SEC_WRITECOMBINE flag when calling the CreateFileMapping function.
        ''' </summary>
        ''' <remarks></remarks>
        PAGE_WRITECOMBINE = 1024

    End Enum

    ''' <summary>
    '''  The standard access rights used by all objects.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ProcessAcessRights As Integer

        ''' <summary>
        ''' Required to delete the object.
        ''' </summary>
        ''' <remarks></remarks>
        DELETE = 65536

        ''' <summary>
        ''' Required to read information in the security descriptor for the object, not including the information in the SACL. To read or write the SACL, you must request the ACCESS_SYSTEM_SECURITY access right.
        ''' </summary>
        ''' <remarks></remarks>
        READ_CONTROL = 131072

        ''' <summary>
        ''' The right to use the object for synchronization. This enables a thread to wait until the object is in the signaled state.
        ''' </summary>
        ''' <remarks></remarks>
        SYNCHRONIZE = 1048576

        ''' <summary>
        ''' Required to modify the DACL in the security descriptor for the object.
        ''' </summary>
        ''' <remarks></remarks>
        WRITE_DAC = 262144

        ''' <summary>
        ''' Required to change the owner in the security descriptor for the object.
        ''' </summary>
        ''' <remarks></remarks>
        WRITE_OWNER = 524288

        'PROCESS_ALL_ACCESS

        ''' <summary>
        ''' Required to create a process.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_CREATE_PROCESS = 128

        ''' <summary>
        ''' Required to create a thread.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_CREATE_THREAD = 2

        ''' <summary>
        ''' Required to duplicate a handle using DuplicateHandle.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_DUP_HANDLE = 64

        ''' <summary>
        ''' Required to retrieve certain information about a process, such as its token, exit code, and priority class (see OpenProcessToken, GetExitCodeProcess, GetPriorityClass, and IsProcessInJob).
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_QUERY_INFORMATION = 1024

        ''' <summary>
        ''' Required to retrieve certain information about a process (see QueryFullProcessImageName). A handle that has the PROCESS_QUERY_INFORMATION access right is automatically granted PROCESS_QUERY_LIMITED_INFORMATION.
        ''' </summary>
        ''' <remarks>Windows Server 2003 and Windows XP/2000:  This access right is not supported.</remarks>
        PROCESS_QUERY_LIMITED_INFORMATION = 4096

        ''' <summary>
        ''' Required to set certain information about a process, such as its priority class (see SetPriorityClass).
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_SET_INFORMATION = 512

        ''' <summary>
        ''' Required to set memory limits using SetProcessWorkingSetSize.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_SET_QUOTA = 256

        ''' <summary>
        ''' Required to suspend or resume a process.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_SUSPEND_RESUME = 2048

        ''' <summary>
        ''' Required to terminate a process using TerminateProcess.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_TERMINATE = 1

        ''' <summary>
        ''' Required to perform an operation on the address space of a process (see VirtualProtectEx and WriteProcessMemory).
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_VM_OPERATION = 8

        ''' <summary>
        ''' Required to read memory in a process using ReadProcessMemory.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_VM_READ = 16

        ''' <summary>
        ''' Required to write to memory in a process using WriteProcessMemory.
        ''' </summary>
        ''' <remarks></remarks>
        PROCESS_VM_WRITE = 32

    End Enum

    ''' <summary>
    ''' Specifies whether certain members of the DeviceMode structure have been initialized. If a member is initialized, its corresponding bit is set, otherwise the bit is clear. A driver supports only those DeviceMode members that are appropriate for the printer or display technology.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags()> _
    Public Enum DeviceModeFields As Integer

        DM_ORIENTATION = 1
        DM_PAPERSIZE = 2
        DM_PAPERLENGTH = 4
        DM_PAPERWIDTH = 8
        DM_SCALE = 16
        DM_POSITION = 32
        DM_NUP = 64
        DM_DISPLAYORIENTATION = 128
        DM_COPIES = 256
        DM_DEFAULTSOURCE = 512
        DM_PRINTQUALITY = 1024
        DM_COLOR = 2048
        DM_DUPLEX = 4096
        DM_YRESOLUTION = 8192
        DM_TTOPTION = 16384
        DM_COLLATE = 32768
        DM_FORMNAME = 65536
        DM_LOGPIXELS = 131072
        DM_BITSPERPEL = 262144
        DM_PELSWIDTH = 524288
        DM_PELSHEIGHT = 1048576
        DM_DISPLAYFLAGS = 2097152
        DM_DISPLAYFREQUENCY = 4194304
        DM_ICMMETHOD = 8388608
        DM_ICMINTENT = 16777216
        DM_MEDIATYPE = 33554432
        DM_DITHERTYPE = 67108864
        DM_PANNINGWIDTH = 134217728
        DM_PANNINGHEIGHT = 268435456
        DM_DISPLAYFIXEDOUTPUT = 536870912

    End Enum

    ''' <summary>
    ''' For printer devices only, the orientation of the paper.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DeviceModeOrientation As Short

        ''' <summary>
        ''' Portrait orientation.
        ''' </summary>
        ''' <remarks></remarks>
        DMORIENT_PORTRAIT = 1

        ''' <summary>
        ''' Landscape orientation.
        ''' </summary>
        ''' <remarks></remarks>
        DMORIENT_LANDSCAPE = 2

    End Enum

    ''' <summary>
    ''' Possible graphics mode values used by EnumDeviceSettings.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum GraphicsMode As Integer

        ''' <summary>
        ''' Retrieve the current settings for the display device.
        ''' </summary>
        ''' <remarks></remarks>
        ENUM_CURRENT_SETTINGS = -1

        ''' <summary>
        ''' Retrieve the settings for the display device that are currently stored in the registry.
        ''' </summary>
        ''' <remarks></remarks>
        ENUM_REGISTRY_SETTINGS = -2

    End Enum

    Public Shared Function IsFlagSet(ByVal flagCollection As Short, ByVal flag As Short) As Boolean
        Return (flagCollection And flag) = flag
    End Function

    Public Shared Function IsFlagSet(ByVal flagCollection As Integer, ByVal flag As Integer) As Boolean
        Return (flagCollection And flag) = flag
    End Function

    Public Shared Function IsFlagSet(ByVal flagCollection As Long, ByVal flag As Long) As Boolean
        Return (flagCollection And flag) = flag
    End Function

    Public Shared Function IsFlagSet(ByVal flagCollection As UShort, ByVal flag As UShort) As Boolean
        Return (flagCollection And flag) = flag
    End Function

    Public Shared Function IsFlagSet(ByVal flagCollection As UInteger, ByVal flag As UInteger) As Boolean
        Return (flagCollection And flag) = flag
    End Function

    Public Shared Function IsFlagSet(ByVal flagCollection As ULong, ByVal flag As ULong) As Boolean
        Return (flagCollection And flag) = flag
    End Function

End Class