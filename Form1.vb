Imports System.Runtime.InteropServices

Public Class Form1
    Inherits System.Windows.Forms.Form

    ' AppBar interop
    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure APPBARDATA
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uCallbackMessage As Integer
        Public uEdge As Integer
        Public rc As RECT
        Public lParam As Integer
    End Structure

    <DllImport("shell32.dll", SetLastError:=True)>
    Private Shared Function SHAppBarMessage(ByVal dwMessage As UInteger, ByRef pData As APPBARDATA) As IntPtr
    End Function

    Private Const ABM_NEW As UInteger = 0
    Private Const ABM_REMOVE As UInteger = 1
    Private Const ABM_QUERYPOS As UInteger = 2
    Private Const ABM_SETPOS As UInteger = 3

    Private Const ABE_LEFT As Integer = 0
    Private Const ABE_TOP As Integer = 1
    Private Const ABE_RIGHT As Integer = 2
    Private Const ABE_BOTTOM As Integer = 3

    Private Const WM_USER As Integer = &H400
    Private ReadOnly CallbackMessage As Integer = WM_USER + 1

    Private appbarRegistered As Boolean = False

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Ensure window is borderless and topmost; designer already sets FormBorderStyle.None
        Me.TopMost = True

        RegisterAppBar()
        SetAppBarPosition()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        UnregisterAppBar()
    End Sub

    Private Sub RegisterAppBar()
        If appbarRegistered Then Return

        Dim abd As APPBARDATA = New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
        abd.hWnd = Me.Handle
        abd.uCallbackMessage = CallbackMessage

        SHAppBarMessage(ABM_NEW, abd)
        appbarRegistered = True
    End Sub

    Private Sub UnregisterAppBar()
        If Not appbarRegistered Then Return

        Dim abd As APPBARDATA = New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
        abd.hWnd = Me.Handle

        SHAppBarMessage(ABM_REMOVE, abd)
        appbarRegistered = False
    End Sub

    Private Sub SetAppBarPosition()
        If Not appbarRegistered Then Return

        Dim abd As APPBARDATA = New APPBARDATA()
        abd.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
        abd.hWnd = Me.Handle
        abd.uEdge = ABE_TOP

        ' Request full width at the top with the current height
        Dim screenBounds = Screen.PrimaryScreen.Bounds
        abd.rc.left = screenBounds.Left
        abd.rc.right = screenBounds.Right
        abd.rc.top = screenBounds.Top
        abd.rc.bottom = abd.rc.top + Me.Height

        ' Query the system for an approved position
        SHAppBarMessage(ABM_QUERYPOS, abd)

        ' System may adjust abd.rc; now set it
        SHAppBarMessage(ABM_SETPOS, abd)

        ' Apply new bounds to the form
        Me.Bounds = New Rectangle(abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = CallbackMessage Then
            ' AppBar state changed (for example: resolution or taskbar moved)
            SetAppBarPosition()
        End If

        MyBase.WndProc(m)
    End Sub

End Class
