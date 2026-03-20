Imports System.Runtime.InteropServices

Module AppBarHelper

    <StructLayout(LayoutKind.Sequential)>
    Public Structure APPBARDATA
        Public cbSize As UInteger
        Public hWnd As IntPtr
        Public uCallbackMessage As UInteger
        Public uEdge As UInteger
        Public rc As RECT
        Public lParam As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <DllImport("shell32.dll", SetLastError:=True)>
    Private Function SHAppBarMessage(dwMessage As UInteger, ByRef pData As APPBARDATA) As UInteger
    End Function

    Private Const ABM_REMOVE As UInteger = &H1
    Private Const ABM_GETSTATE As UInteger = &H4
    Private Const ABM_GETTASKBARPOS As UInteger = &H5

    Public Sub RemoveTaskbarReservation()
        Dim abd As New APPBARDATA()
        abd.cbSize = CUInt(Marshal.SizeOf(abd))

        ' Get taskbar handle
        abd.hWnd = FindWindow("Shell_TrayWnd", Nothing)

        If abd.hWnd <> IntPtr.Zero Then
            SHAppBarMessage(ABM_REMOVE, abd)
        End If
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

End Module
