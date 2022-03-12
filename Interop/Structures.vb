Imports System.Runtime.InteropServices

''' <summary>
''' Contains structures used by many native functions,
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Structures

    ''' <summary>
    ''' The DISPLAY_DEVICE structure receives information about the display device specified by the iDevNum parameter of the EnumDisplayDevices function.
    ''' </summary>
    ''' <remarks>The four string members are set based on the parameters passed to EnumDisplayDevices. If the lpDevice param is NULL, then DISPLAY_DEVICE is filled in with information about the display adapter(s). If it is a valid device name, then it is filled in with information about the monitor(s) for that device.</remarks>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
    Public Structure DisplayDevice

        ''' <summary>
        ''' Size, in bytes, of the DISPLAY_DEVICE structure. This must be initialized prior to calling EnumDisplayDevices.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.U4)> _
        Public cb As Integer

        ''' <summary>
        ''' An array of characters identifying the device name. This is either the adapter device or the monitor device.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public DeviceName As String

        ''' <summary>
        ''' An array of characters containing the device context string. This is either a description of the display adapter or of the display monitor.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public DeviceString As String

        ''' <summary>
        ''' Device state flags
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.U4)> _
        Public StateFlags As Constants.DisplayDeviceStateFlags

        ''' <summary>
        ''' Not used.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public DeviceID As String

        ''' <summary>
        ''' Not used.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public DeviceKey As String

    End Structure

    ''' <summary>
    ''' Contains information about the initialization and environment of a printer or a display device.
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure DeviceMode

        ''' <summary>
        ''' A zero-terminated character array that specifies the "friendly" name of the printer or display; for example, "PCL/HP LaserJet" in the case of PCL/HP LaserJet. This string is unique among device drivers. Note that this name may be truncated to fit in the dmDeviceName array.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public dmDeviceName As String

        ''' <summary>
        ''' The version number of the initialization data specification on which the structure is based. To ensure the correct version is used for any operating system, use DM_SPECVERSION.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmSpecVersion As Short

        ''' <summary>
        ''' The driver version number assigned by the driver developer.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDriverVersion As Short

        ''' <summary>
        ''' Specifies the size, in bytes, of the DeviceMode structure, not including any private driver-specific data that might follow the structure's public members. Set this member to sizeof (DEVMODE) to indicate the version of the DEVMODE structure being used.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmSize As Short

        ''' <summary>
        ''' Contains the number of bytes of private driver-data that follow this structure. If a device driver does not use device-specific information, set this member to zero.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDriverExtra As Short

        ''' <summary>
        ''' Specifies whether certain members of the DEVMODE structure have been initialized. If a member is initialized, its corresponding bit is set, otherwise the bit is clear. A driver supports only those DEVMODE members that are appropriate for the printer or display technology.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmFields As Constants.DeviceModeFields

        ''' <summary>
        ''' For printer devices only, selects the orientation of the paper. This member can be either DMORIENT_PORTRAIT (1) or DMORIENT_LANDSCAPE (2).
        ''' </summary>
        ''' <remarks></remarks>
        Public dmOrientation As Constants.DeviceModeOrientation

        ''' <summary>
        ''' For printer devices only, selects the size of the paper to print on. This member can be set to zero if the length and width of the paper are both set by the dmPaperLength and dmPaperWidth members.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPaperSize As Short

        ''' <summary>
        ''' For printer devices only, overrides the length of the paper specified by the dmPaperSize member, either for custom paper sizes or for devices such as dot-matrix printers that can print on a page of arbitrary length. These values, along with all other values in this structure that specify a physical length, are in tenths of a millimeter.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPaperLength As Short

        ''' <summary>
        ''' For printer devices only, overrides the width of the paper specified by the dmPaperSize member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPaperWidth As Short

        ''' <summary>
        ''' Specifies the factor by which the printed output is to be scaled. The apparent page size is scaled from the physical page size by a factor of dmScale /100. For example, a letter-sized page with a dmScale value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmScale As Short

        ''' <summary>
        ''' Selects the number of copies printed if the device supports multiple-page copies.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmCopies As Short

        ''' <summary>
        ''' Specifies the paper source. To retrieve a list of the available paper sources for a printer, use the DeviceCapabilities function with the DC_BINS flag.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDefaultSource As Short

        ''' <summary>
        ''' Specifies the printer resolution.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPrintQuality As Short

        ''' <summary>
        ''' Switches between color and monochrome on color printers.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmColor As Short

        ''' <summary>
        ''' Selects duplex or double-sided printing for printers capable of duplex printing.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDuplex As Short

        ''' <summary>
        ''' Specifies the y-resolution, in dots per inch, of the printer. If the printer initializes this member, the dmPrintQuality member specifies the x-resolution, in dots per inch, of the printer.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmYResolution As Short

        ''' <summary>
        ''' Specifies how TrueType fonts should be printed.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmTTOption As Short

        ''' <summary>
        ''' Specifies whether collation should be used when printing multiple copies. (This member is ignored unless the printer driver indicates support for collation by setting the dmFields member to DM_COLLATE.)
        ''' </summary>
        ''' <remarks></remarks>
        Public dmCollate As Short

        ''' <summary>
        ''' A zero-terminated character array that specifies the name of the form to use; for example, "Letter" or "Legal". A complete set of names can be retrieved by using the EnumForms function.
        ''' </summary>
        ''' <remarks></remarks>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public dmFormName As String

        ''' <summary>
        ''' The number of pixels per logical inch. Printer drivers do not use this member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmLogPixels As Short

        ''' <summary>
        ''' Specifies the color resolution, in bits per pixel, of the display device (for example: 4 bits for 16 colors, 8 bits for 256 colors, or 16 bits for 65,536 colors). Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmBitsPerPel As Integer ' Declared wrong in the full framework

        ''' <summary>
        ''' Specifies the width, in pixels, of the visible device surface. Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPelsWidth As Integer

        ''' <summary>
        ''' Specifies the height, in pixels, of the visible device surface. Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPelsHeight As Integer

        ''' <summary>
        ''' Specifies the device's display mode.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDisplayFlags As Integer

        ''' <summary>
        ''' Specifies the frequency, in hertz (cycles per second), of the display device in a particular mode. This value is also known as the display device's vertical refresh rate. Display drivers use this member. It is used, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member. When you call the EnumDisplaySettings function, the dmDisplayFrequency member may return with the value 0 or 1. These values represent the display hardware's default refresh rate. This default rate is typically set by switches on a display card or computer motherboard, or by a configuration program that does not use display functions such as ChangeDisplaySettings.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDisplayFrequency As Integer

        ''' <summary>
        ''' Specifies how ICM is handled.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmICMMethod As Integer

        ''' <summary>
        ''' Specifies which color matching method, or intent, should be used by default. This member is primarily for non-ICM applications. ICM applications can establish intents by using the ICM functions.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmICMIntent As Integer

        ''' <summary>
        ''' Specifies the type of media being printed on.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmMediaType As Integer

        ''' <summary>
        ''' Specifies how dithering is to be done.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmDitherType As Integer

        ''' <summary>
        ''' Not used; must be zero.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmReserved1 As Integer

        ''' <summary>
        ''' Not used; must be zero.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmReserved2 As Integer

        ''' <summary>
        ''' This member must be zero.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPanningWidth As Integer

        ''' <summary>
        ''' This member must be zero.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPanningHeight As Integer

        ''' <summary>
        ''' For display devices only, indicates the positional X coordinate of the display device in reference to the desktop area. The primary display device is always located at 0.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPositionX As Integer ' Using a PointL Struct does not work

        ''' <summary>
        ''' For display devices only, indicates the positional Y coordinate of the display device in reference to the desktop area. The primary display device is always located at 0.
        ''' </summary>
        ''' <remarks></remarks>
        Public dmPositionY As Integer

    End Structure

End Class