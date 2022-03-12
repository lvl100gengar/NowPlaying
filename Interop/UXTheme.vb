Imports System
Imports System.Runtime.InteropServices

''' <summary>
''' Contains common functions from the UXTheme library.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class UXTheme

    ''' <summary>
    ''' Causes a window to use a different set of visual style information than its class normally uses.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose visual style information is to be changed.</param>
    ''' <param name="pszSubAppName">Pointer to a string that contains the application name to use in place of the calling application's name. If this parameter is NULL, the calling application's name is used.</param>
    ''' <param name="pszSubIdList">Pointer to a string that contains a semicolon-separated list of CLSID names to use in place of the actual list passed by the window's class. If this parameter is NULL, the ID list from the calling class is used.</param>
    ''' <returns>If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
    ''' <remarks>The theme manager retains the pszSubAppName and the pszSubIdList associations through the lifetime of the window, even if visual styles subsequently change. The window is sent a WM_THEMECHANGED message at the end of a SetWindowTheme call, so that the new visual style can be found and applied.</remarks>
    <DllImport("uxtheme.dll", CharSet:=CharSet.Unicode)> _
    Public Shared Function SetWindowTheme(hWnd As IntPtr, pszSubAppName As String, pszSubIdList As String) As Integer

    End Function

End Class