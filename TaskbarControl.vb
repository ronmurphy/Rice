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

    Private Const SWP_HIDEWINDOW As UInteger = &H80
    Private Const SWP_SHOWWINDOW As UInteger = &H40
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_FRAMECHANGED As UInteger = &H20

    Public Shared Sub HideTaskbar()
        Dim hWnd = FindWindow("Shell_TrayWnd", Nothing)
        If hWnd <> IntPtr.Zero Then
            SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                             SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_HIDEWINDOW)
        End If
    End Sub

    Public Shared Sub ShowTaskbar()
        Dim hWnd = FindWindow("Shell_TrayWnd", Nothing)
        If hWnd <> IntPtr.Zero Then
            SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                             SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW)
        End If
    End Sub

End Class


