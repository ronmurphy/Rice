Imports System.Runtime.InteropServices

Friend Class TaskbarControl

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(hWnd As IntPtr,
                                             hWndInsertAfter As IntPtr,
                                             x As Integer,
                                             y As Integer,
                                             cx As Integer,
                                             cy As Integer,
                                             uFlags As UInteger) As Boolean
    End Function

    <DllImport("shell32.dll", SetLastError:=True)>
    Private Shared Function SHAppBarMessage(dwMessage As UInteger, ByRef pData As APPBARDATA) As IntPtr
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure APPBARDATA
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uCallbackMessage As UInteger
        Public uEdge As UInteger
        Public rc As RECT
        Public lParam As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Private Const ABM_GETSTATE As UInteger = &H4
    Private Const ABM_SETSTATE As UInteger = &HA

    Private Const ABS_AUTOHIDE As Integer = &H1
    Private Const ABS_ALWAYSONTOP As Integer = &H2

    Private Const SWP_HIDEWINDOW As UInteger = &H80
    Private Const SWP_SHOWWINDOW As UInteger = &H40
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_FRAMECHANGED As UInteger = &H20

    ' Remember original state so we can restore it
    Private Shared _originalState As Integer = -1

    Public Shared Sub HideTaskbar()
        Dim hWnd = FindWindow("Shell_TrayWnd", Nothing)
        If hWnd = IntPtr.Zero Then Return

        ' Save the original taskbar state (auto-hide vs always-on-top)
        Dim abd As New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
        abd.hWnd = hWnd

        If _originalState = -1 Then
            _originalState = CInt(SHAppBarMessage(ABM_GETSTATE, abd))
        End If

        ' Set taskbar to auto-hide mode — this releases its work area reservation
        abd.lParam = ABS_AUTOHIDE
        SHAppBarMessage(ABM_SETSTATE, abd)

        ' Then hide the window so it doesn't pop up on mouse hover at screen edge
        SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                         SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_HIDEWINDOW)
    End Sub

    Public Shared Sub ShowTaskbar()
        Dim hWnd = FindWindow("Shell_TrayWnd", Nothing)
        If hWnd = IntPtr.Zero Then Return

        ' Show the window first
        SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                         SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW)

        ' Restore original taskbar state
        If _originalState >= 0 Then
            Dim abd As New APPBARDATA()
            abd.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
            abd.hWnd = hWnd
            abd.lParam = _originalState
            SHAppBarMessage(ABM_SETSTATE, abd)
        End If
    End Sub

End Class
