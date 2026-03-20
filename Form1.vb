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
    Private currentScreenDeviceName As String = String.Empty
    ' WinEvent hook for window create/destroy/show/hide
    Private Delegate Sub WinEventDelegate(hWinEventHook As IntPtr, eventType As UInteger, hWnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
    Private winEventHook1 As IntPtr = IntPtr.Zero
    Private winEventHook2 As IntPtr = IntPtr.Zero
    Private winEventCallback As WinEventDelegate
    Private currentRunning As New Dictionary(Of Integer, Button)()
    Private notifyIcon As System.Windows.Forms.NotifyIcon
    Private trayMenu As System.Windows.Forms.ContextMenuStrip
    Private trayToolTip As New ToolTip()

    Private Const EVENT_OBJECT_CREATE As UInteger = &H8000
    Private Const EVENT_OBJECT_DESTROY As UInteger = &H8001
    Private Const EVENT_OBJECT_SHOW As UInteger = &H8002
    Private Const EVENT_OBJECT_HIDE As UInteger = &H8003
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = &H3
    Private Const WINEVENT_OUTOFCONTEXT As UInteger = 0

    Private TaskbarControl As New TaskbarControl()

    <DllImport("user32.dll")>
    Private Shared Function SetWinEventHook(eventMin As UInteger, eventMax As UInteger, hmodWinEventProc As IntPtr, callback As WinEventDelegate, idProcess As UInteger, idThread As UInteger, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnhookWinEvent(hWinEventHook As IntPtr) As Boolean
    End Function



    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Ensure window is borderless and topmost; designer already sets FormBorderStyle.None
        Me.TopMost = True

        RegisterAppBar()
        SetAppBarPosition()
        AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
        ' Install WinEvent hooks for window create/destroy and foreground changes
        winEventCallback = AddressOf WinEventProc
        winEventHook1 = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_DESTROY, IntPtr.Zero, winEventCallback, 0, 0, WINEVENT_OUTOFCONTEXT)
        winEventHook2 = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_HIDE, IntPtr.Zero, winEventCallback, 0, 0, WINEVENT_OUTOFCONTEXT)

    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        ' Start/placeholder button
        Dim btnStart As New Button()
        btnStart.Text = "Start"
        btnStart.AutoSize = True
        btnStart.Margin = New Padding(4, 4, 4, 4)
        btnStart.FlatStyle = FlatStyle.System
        AddHandler btnStart.Click, AddressOf ShowStartForm
        flowLeft.Controls.Add(btnStart)

        ' Load pinned apps (from Taskbar pinned shortcuts)
        LoadPinnedShortcuts()

        ' Populate real system tray
        PopulateSystemTray()

        ' Make click area respond
        AddHandler panelClickArea.Click, Sub(s, ev) MessageBox.Show("Centered click area")

        ' Setup real system tray NotifyIcon
        Try
            trayMenu = New ContextMenuStrip()
            trayMenu.Items.Add("Open", Nothing, Sub(sa, ea) Me.Invoke(Sub() Me.Show()))
            trayMenu.Items.Add("Exit", Nothing, Sub(sa, ea) Me.Invoke(Sub() Me.Close()))

            notifyIcon = New NotifyIcon()
            notifyIcon.Icon = SystemIcons.Application
            notifyIcon.ContextMenuStrip = trayMenu
            notifyIcon.Text = "Rice"
            notifyIcon.Visible = True
            AddHandler notifyIcon.DoubleClick, AddressOf NotifyIcon_DoubleClick
        Catch
        End Try

        ' Start clock
        AddHandler timerClock.Tick, AddressOf TimerClock_Tick
        timerClock.Start()

        ' Ensure clock label is dock-filled inside centerContainer; we'll center centerContainer in SetAppBarPosition
        Try
            lblClock.Dock = DockStyle.Fill
            lblClock.AutoSize = False
            lblClock.TextAlign = ContentAlignment.MiddleCenter
            lblClock.Font = New Font("Segoe UI", 12.0F, FontStyle.Regular, GraphicsUnit.Point)
            lblClock.Visible = True
            lblClock.BringToFront()
        Catch
        End Try

        ' Color layout regions when debugging to help visualize columns
        Try
            If System.Diagnostics.Debugger.IsAttached Then
                flowLeft.BackColor = Color.LightGreen
                centerPanel.BackColor = Color.LightBlue
                panelRight.BackColor = Color.LightCoral
                lblClock.ForeColor = Color.White
            Else
                flowLeft.BackColor = Color.Transparent
                centerPanel.BackColor = Color.Transparent
                panelRight.BackColor = Color.Transparent
                lblClock.ForeColor = Color.Black
            End If
        Catch
        End Try

        ' Running apps refresh
        ' We'll use WinEvent hooks instead of polling for running apps

        ' initial population of running apps
        RefreshRunningApps()
    End Sub

    Private Sub ShowStartForm(sender As Object, e As EventArgs)
        Try
            Dim f As New frmStart()
            ' Show as modal so it behaves like a start menu
            f.StartPosition = FormStartPosition.CenterScreen
            f.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Failed to open Start: " & ex.Message)
        End Try
    End Sub

    Private Sub NotifyIcon_DoubleClick(sender As Object, e As EventArgs)
        If Me.IsHandleCreated Then
            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 Me.WindowState = FormWindowState.Normal
                                                 Me.Show()
                                                 Me.BringToFront()
                                             End Sub))
        End If
    End Sub

    Private Sub TimerClock_Tick(sender As Object, e As EventArgs)
        lblClock.Text = DateTime.Now.ToString("h:mm:ss tt")
    End Sub

    Private Sub PopulateSystemTray()
        If panelTray Is Nothing Then Return
        Try
            panelTray.Controls.Clear()

            Dim p = SystemInformation.PowerStatus
            Dim items = {
                If(Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable(), "Net", "No Net"),
                If(p.PowerLineStatus = PowerLineStatus.Online, "AC", CInt(p.BatteryLifePercent * 100).ToString() & "%")
            }
            For Each t In items
                Dim lbl = New Label()
                lbl.AutoSize = True
                lbl.TextAlign = ContentAlignment.MiddleCenter
                lbl.Margin = New Padding(4, 4, 4, 4)
                lbl.Font = New Font("Segoe UI", 9.0F)
                lbl.Text = t
                trayToolTip.SetToolTip(lbl, t)
                panelTray.Controls.Add(lbl)
            Next
        Catch
        End Try
    End Sub

    Private Sub LoadPinnedShortcuts()
        ' Only use Taskbar pinned shortcuts (per-user and common)
        Dim added = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim taskband = IO.Path.Combine(appData, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
        Dim commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
        Dim commonTaskband = IO.Path.Combine(commonAppData, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")

        Try
            Dim shortcuts As New List(Of String)
            If IO.Directory.Exists(taskband) Then shortcuts.AddRange(IO.Directory.GetFiles(taskband, "*.lnk"))
            If IO.Directory.Exists(commonTaskband) Then shortcuts.AddRange(IO.Directory.GetFiles(commonTaskband, "*.lnk"))

            ' Keep a reasonable maximum to avoid flooding the bar
            Dim maxPinned = 40
            For Each f In shortcuts.Distinct()
                Try
                    Dim target = ResolveShortcutTarget(f)
                    If String.IsNullOrEmpty(target) OrElse Not IO.File.Exists(target) Then Continue For
                    ' skip shortcuts that point to this application
                    Try
                        Dim mepath = Process.GetCurrentProcess().MainModule.FileName
                        If String.Equals(IO.Path.GetFullPath(target), IO.Path.GetFullPath(mepath), StringComparison.OrdinalIgnoreCase) Then Continue For
                    Catch
                    End Try
                    ' filter out installers, uninstallers, msi packages, or non-exe targets
                    Dim ext = IO.Path.GetExtension(target).ToLowerInvariant()
                    If ext <> ".exe" Then Continue For
                    Dim fname = IO.Path.GetFileName(target).ToLowerInvariant()
                    If fname.Contains("uninstall") OrElse fname.Contains("setup") OrElse fname.Contains("installer") Then Continue For

                    ' keep a reasonable maximum to avoid flooding the bar
                    If added.Count >= maxPinned Then Exit For
                    Dim icon = GetIconFromShortcut(f)
                    If icon IsNot Nothing Then
                        Dim btn = New Button()
                        btn.Width = 36
                        btn.Height = 36
                        btn.BackgroundImage = icon.ToBitmap()
                        btn.BackgroundImageLayout = ImageLayout.Zoom
                        btn.Tag = f
                        btn.Margin = New Padding(2)
                        AddHandler btn.Click, AddressOf PinnedShortcut_Click
                        flowLeft.Controls.Add(btn)
                        added.Add(f)
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    Private Sub PinnedShortcut_Click(sender As Object, e As EventArgs)
        Dim btn = DirectCast(sender, Button)
        Dim path = TryCast(btn.Tag, String)
        If Not String.IsNullOrEmpty(path) Then
            Try
                Process.Start(New ProcessStartInfo(path) With {.UseShellExecute = True})
            Catch ex As Exception
                MessageBox.Show("Failed to launch: " & ex.Message)
            End Try
        End If
    End Sub

    Private Function GetIconFromShortcut(shortcutPath As String) As Icon
        Try
            Dim target = ResolveShortcutTarget(shortcutPath)
            If Not String.IsNullOrEmpty(target) AndAlso IO.File.Exists(target) Then
                Return Icon.ExtractAssociatedIcon(target)
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Function ResolveShortcutTarget(shortcutPath As String) As String
        Try
            Dim shell = CreateObject("WScript.Shell")
            Dim lnk = shell.CreateShortcut(shortcutPath)
            Dim target = lnk.TargetPath
            Return If(target, String.Empty)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Sub RefreshRunningApps()
        ' incremental update: compare current running processes with currentRunning dictionary
        Try
            Dim procs = Process.GetProcesses().Where(Function(p) p.MainWindowHandle <> IntPtr.Zero).ToList()
            Dim currentPids = New HashSet(Of Integer)(procs.Select(Function(p) p.Id))

            ' remove entries that no longer exist
            For Each pid In currentRunning.Keys.ToList()
                If Not currentPids.Contains(pid) Then
                    RemoveRunningApp(pid)
                End If
            Next

            ' add new processes
            Dim myPid = Process.GetCurrentProcess().Id
            For Each p In procs
                If p.Id = myPid Then Continue For
                If Not currentRunning.ContainsKey(p.Id) Then
                    AddRunningApp(p)
                End If
            Next
        Catch
        End Try
    End Sub

    Private Sub AddRunningApp(p As Process)
        Try
            Dim icon As Icon = Nothing
            Dim path As String = Nothing
            Try
                path = p.MainModule.FileName
            Catch
            End Try
            If Not String.IsNullOrEmpty(path) AndAlso IO.File.Exists(path) Then
                icon = Icon.ExtractAssociatedIcon(path)
            End If

            Dim btn = New Button()
            btn.Width = 36
            btn.Height = 36
            If icon IsNot Nothing Then
                btn.BackgroundImage = icon.ToBitmap()
                btn.BackgroundImageLayout = ImageLayout.Zoom
            Else
                btn.Text = p.ProcessName.Substring(0, Math.Min(3, p.ProcessName.Length))
                btn.Font = New Font("Segoe UI", 6)
            End If
            btn.Tag = "running:" & p.Id
            btn.Margin = New Padding(2)
            AddHandler btn.Click, Sub(s, ev)
                                      Try
                                          NativeMethods.SetForegroundWindow(p.MainWindowHandle)
                                      Catch
                                      End Try
                                  End Sub

            currentRunning(p.Id) = btn
            ' add after pinned items: find index after pinned + start button
            flowLeft.Controls.Add(btn)
        Catch
        End Try
    End Sub

    Private Sub RemoveRunningApp(pid As Integer)
        Try
            If currentRunning.ContainsKey(pid) Then
                Dim btn = currentRunning(pid)
                If flowLeft.Controls.Contains(btn) Then
                    flowLeft.Controls.Remove(btn)
                End If
                btn.Dispose()
                currentRunning.Remove(pid)
            End If
        Catch
        End Try
    End Sub

    Private Sub WinEventProc(hWinEventHook As IntPtr, eventType As UInteger, hWnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
        ' When windows are created/destroyed/shown/hidden, refresh running apps incrementally
        If Me.IsHandleCreated Then
            Me.BeginInvoke(New MethodInvoker(AddressOf RefreshRunningApps))
        End If
    End Sub

    Private Class NativeMethods
        <DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function
    End Class

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
        If winEventHook1 <> IntPtr.Zero Then UnhookWinEvent(winEventHook1)
        If winEventHook2 <> IntPtr.Zero Then UnhookWinEvent(winEventHook2)
        If notifyIcon IsNot Nothing Then
            notifyIcon.Visible = False
            notifyIcon.Dispose()
            notifyIcon = Nothing
        End If
        If trayMenu IsNot Nothing Then
            trayMenu.Dispose()
            trayMenu = Nothing
        End If
        trayToolTip.Dispose()

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

        ' Choose the screen that currently contains the form (or fallback to cursor)
        Dim scr As Screen = Nothing
        Try
            scr = Screen.FromRectangle(Me.Bounds)
        Catch
            scr = Screen.FromPoint(Cursor.Position)
        End Try

        If scr Is Nothing Then scr = Screen.PrimaryScreen

        abd.rc.left = scr.Bounds.Left
        abd.rc.right = scr.Bounds.Right
        abd.rc.top = scr.Bounds.Top
        abd.rc.bottom = abd.rc.top + Me.Height

        ' Query the system for an approved position on that monitor
        SHAppBarMessage(ABM_QUERYPOS, abd)
        SHAppBarMessage(ABM_SETPOS, abd)

        ' Apply new bounds to the form
        Me.Bounds = New Rectangle(abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top)

        ' Set left/center/right columns to each be one third of the screen width
        Try
            Dim totalW = Me.ClientSize.Width
            Dim colW = Math.Max(100, totalW \ 3)
            flowLeft.Width = colW
            panelRight.Width = colW
            ' centerPanel is Dock=Fill so it will size automatically between left and right

            ' move overflow pinned items into center overflow panel if left is too full
            ' simple approach: move last controls to overflow until left fits or overflow reaches centerContainer width
            Try
                Dim leftWidthAcc = 0
                For i = 0 To flowLeft.Controls.Count - 1
                    Dim c = flowLeft.Controls(i)
                    leftWidthAcc += c.Width + c.Margin.Horizontal
                Next
                While leftWidthAcc > flowLeft.Width AndAlso flowLeft.Controls.Count > 0
                    Dim c = flowLeft.Controls(flowLeft.Controls.Count - 1)
                    flowLeft.Controls.RemoveAt(flowLeft.Controls.Count - 1)
                    overflowPanelInCenter.Controls.Add(c)
                    leftWidthAcc = leftWidthAcc - (c.Width + c.Margin.Horizontal)
                    If overflowPanelInCenter.PreferredSize.Width > centerContainer.Width Then Exit While
                End While
            Catch
            End Try
        Catch
        End Try

        ' remember which screen we're on
        currentScreenDeviceName = scr.DeviceName

        ' Position centerContainer so the clock is at the true horizontal center of the active monitor
        Try
            ' scr.Bounds.Width is the monitor width; flowLeft.Width is how far centerPanel starts from the left edge
            Dim screenCenterInPanel = (scr.Bounds.Width \ 2) - flowLeft.Width
            Dim cw = centerContainer.Width
            Dim ch = centerContainer.Height
            Dim cx = screenCenterInPanel - (cw \ 2)
            cx = Math.Max(0, Math.Min(cx, centerPanel.ClientSize.Width - cw))
            centerContainer.Location = New Point(cx, Math.Max(0, (Me.ClientSize.Height - ch) \ 2))
        Catch
        End Try
    End Sub

    Private Sub OnDisplaySettingsChanged(sender As Object, e As EventArgs)
        ' Recalculate position when display configuration changes
        If Me.InvokeRequired Then
            Me.BeginInvoke(New MethodInvoker(AddressOf SetAppBarPosition))
        Else
            SetAppBarPosition()
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = CallbackMessage Then
            ' AppBar state changed (for example: resolution or taskbar moved)
            SetAppBarPosition()
        End If

        MyBase.WndProc(m)
    End Sub

    Protected Overrides Sub OnMove(e As EventArgs)
        MyBase.OnMove(e)
        ' If the form moved to another monitor, reposition the AppBar
        Dim scr = Screen.FromRectangle(Me.Bounds)
        If scr.DeviceName <> currentScreenDeviceName Then
            SetAppBarPosition()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Set native taskbar to auto-hide (releases its work area) then hide it
        TaskbarControl.HideTaskbar()
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        TaskbarControl.ShowTaskbar()
    End Sub
End Class
