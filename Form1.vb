Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Windows.Automation

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region "DWM Mica / Acrylic"
    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmExtendFrameIntoClientArea(hwnd As IntPtr, ByRef margins As DWM_MARGINS) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DWM_MARGINS
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
    Private Const DWMWA_SYSTEMBACKDROP_TYPE As Integer = 38
    Private Const DWMWA_MICA_EFFECT As Integer = 1029
    Private Const DWMSBT_MAINWINDOW As Integer = 2      ' Mica
    Private Const DWMSBT_TRANSIENTWINDOW As Integer = 3  ' Acrylic
    Private Const DWMSBT_TABBEDWINDOW As Integer = 4     ' Mica Alt

    Private Const WS_CAPTION_STYLE As Integer = &HC00000
    Private Const WM_NCCALCSIZE As Integer = &H83

    Private taskbarMicaApplied As Boolean = False
    Private isDarkMode As Boolean = True

    ''' <summary>
    ''' Detects Windows light/dark mode from registry.
    ''' HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize
    ''' AppsUseLightTheme: 0 = dark, 1 = light
    ''' </summary>
    Private Shared Function DetectDarkMode() As Boolean
        Try
            Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim val = key.GetValue("AppsUseLightTheme")
                    If val IsNot Nothing Then Return CInt(val) = 0
                End If
            End Using
        Catch
        End Try
        Return True ' default to dark
    End Function

    Private Sub ApplyThemeColors()
        isDarkMode = DetectDarkMode()

        If taskbarMicaApplied Then
            ' Near-black/white to let backdrop show — Form doesn't support true transparent
            Me.BackColor = If(isDarkMode, Color.FromArgb(255, 1, 1, 1), Color.FromArgb(255, 254, 254, 254))
        Else
            Me.BackColor = If(isDarkMode, Color.FromArgb(255, 32, 32, 32), Color.FromArgb(255, 243, 243, 243))
        End If

        Dim textColor = If(isDarkMode, Color.FromArgb(240, 240, 240), Color.FromArgb(30, 30, 30))
        Dim hoverBg = If(isDarkMode, Color.FromArgb(40, 255, 255, 255), Color.FromArgb(40, 0, 0, 0))
        Dim hoverDown = If(isDarkMode, Color.FromArgb(60, 255, 255, 255), Color.FromArgb(60, 0, 0, 0))

        ' Update clock
        Try
            lblClock.ForeColor = textColor
        Catch
        End Try

        ' Update all flat-styled buttons in flowLeft
        Try
            For Each ctrl As Control In flowLeft.Controls
                If TypeOf ctrl Is Button Then
                    Dim btn = DirectCast(ctrl, Button)
                    btn.ForeColor = textColor
                    btn.FlatAppearance.MouseOverBackColor = hoverBg
                    btn.FlatAppearance.MouseDownBackColor = hoverDown
                End If
            Next
        Catch
        End Try

        ' Update tray PictureBoxes
        Try
            For Each ctrl As Control In panelTray.Controls
                If TypeOf ctrl Is PictureBox Then
                    ' hover handlers already set — just update any text labels if needed
                End If
            Next
        Catch
        End Try
    End Sub

    Private Sub ApplyTaskbarBackdrop()
        Try
            isDarkMode = DetectDarkMode()

            ' Tell DWM whether to use dark or light frame
            Dim darkMode As Integer = If(isDarkMode, 1, 0)
            DwmSetWindowAttribute(Me.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, darkMode, 4)

            ' Extend frame into entire client area
            Dim m As New DWM_MARGINS() With {.Left = -1, .Right = -1, .Top = -1, .Bottom = -1}
            DwmExtendFrameIntoClientArea(Me.Handle, m)

            ' Use Acrylic for the taskbar — it's a thin strip so translucent blur looks great
            Dim backdropType As Integer = DWMSBT_TRANSIENTWINDOW
            Dim hr = DwmSetWindowAttribute(Me.Handle, DWMWA_SYSTEMBACKDROP_TYPE, backdropType, 4)

            If hr <> 0 Then
                ' Fallback to Mica
                backdropType = DWMSBT_MAINWINDOW
                hr = DwmSetWindowAttribute(Me.Handle, DWMWA_SYSTEMBACKDROP_TYPE, backdropType, 4)
            End If

            If hr <> 0 Then
                ' Fallback to older Mica attribute
                Dim micaOn As Integer = 1
                DwmSetWindowAttribute(Me.Handle, DWMWA_MICA_EFFECT, micaOn, 4)
            End If

            taskbarMicaApplied = True
        Catch
            taskbarMicaApplied = False
        End Try

        ' Apply theme-aware colors
        ApplyThemeColors()
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp = MyBase.CreateParams
            ' DWM needs WS_CAPTION for backdrop effects on borderless windows
            cp.Style = cp.Style Or WS_CAPTION_STYLE
            cp.ExStyle = cp.ExStyle Or &H2000000 ' WS_EX_COMPOSITED — reduce flicker
            Return cp
        End Get
    End Property
#End Region

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
    Private lastActivatedPid As Integer = -1
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

    ' System tray embedding P/Invoke
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowEx(hwndParent As IntPtr, hwndChildAfter As IntPtr, lpszClass As String, lpszWindow As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetParent(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function MoveWindow(hWnd As IntPtr, X As Integer, Y As Integer, nWidth As Integer, nHeight As Integer, bRepaint As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    Private Const SW_SHOW As Integer = 5
    Private Const SW_HIDE As Integer = 0

    ' System tray state
    Private trayIconButtons As New List(Of PictureBox)()
    Private trayRefreshTimer As Timer



    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Ensure window is borderless and topmost; designer already sets FormBorderStyle.None
        Me.TopMost = True

        ' Apply Acrylic/Mica backdrop to the taskbar
        ApplyTaskbarBackdrop()

        RegisterAppBar()
        SetAppBarPosition()
        AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
        ' Install WinEvent hooks for window create/destroy and foreground changes
        winEventCallback = AddressOf WinEventProc
        winEventHook1 = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_DESTROY, IntPtr.Zero, winEventCallback, 0, 0, WINEVENT_OUTOFCONTEXT)
        winEventHook2 = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_HIDE, IntPtr.Zero, winEventCallback, 0, 0, WINEVENT_OUTOFCONTEXT)

        ''brad code, trying to make frmStart faster to open by preloading it in the background
        ' calling .update() as how for some reason .load() does not seem to exist any more?
        frmStart.Update()

    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        ' IMPORTANT: Enumerate tray icons BEFORE MyBase.OnLoad raises Load event,
        ' which triggers Form1_Load → HideTaskbar(). UI Automation can't find
        ' buttons in a hidden Shell_TrayWnd.
        PopulateSystemTray()

        MyBase.OnLoad(e)

        ' Start button — use a Windows logo-style grid icon
        Dim btnStart As New Button()
        btnStart.Size = New Size(36, 36)
        btnStart.Margin = New Padding(2)
        ApplyFlatStyle(btnStart, "Start")
        ' Draw a 4-square Windows logo
        Dim startBmp As New Bitmap(24, 24)
        Using g = Graphics.FromImage(startBmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            Dim clr = Color.FromArgb(96, 165, 250)
            Using br As New SolidBrush(clr)
                g.FillRectangle(br, 3, 3, 8, 8)    ' top-left
                g.FillRectangle(br, 13, 3, 8, 8)   ' top-right
                g.FillRectangle(br, 3, 13, 8, 8)   ' bottom-left
                g.FillRectangle(br, 13, 13, 8, 8)  ' bottom-right
            End Using
        End Using
        btnStart.BackgroundImage = startBmp
        btnStart.BackgroundImageLayout = ImageLayout.Center
        AddHandler btnStart.Click, AddressOf ShowStartForm
        flowLeft.Controls.Add(btnStart)

        ' Load pinned apps (from Taskbar pinned shortcuts)
        LoadPinnedShortcuts()

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

        ' Make panels transparent so backdrop shows through
        Try
            flowLeft.BackColor = Color.Transparent
            centerPanel.BackColor = Color.Transparent
            panelRight.BackColor = Color.Transparent
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
            EmbedSystemTray()
        Catch
            ' Fallback: show basic status labels if embedding fails
            Try
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
        End Try
    End Sub

    Private Sub EmbedSystemTray()
        ' Build our own system tray by enumerating real tray icons + system status
        BuildCustomTray()
    End Sub

    Private Sub RestoreSystemTray()
        ' Nothing to restore — we built our own tray, didn't steal Windows'
        If trayRefreshTimer IsNot Nothing Then
            trayRefreshTimer.Stop()
            trayRefreshTimer.Dispose()
            trayRefreshTimer = Nothing
        End If
    End Sub

#Region "Custom System Tray"
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetClassName(hWnd As IntPtr, lpClassName As System.Text.StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function OpenProcess(dwDesiredAccess As UInteger, bInheritHandle As Boolean, dwProcessId As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(hObject As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function VirtualAllocEx(hProcess As IntPtr, lpAddress As IntPtr, dwSize As UInteger, flAllocationType As UInteger, flProtect As UInteger) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function VirtualFreeEx(hProcess As IntPtr, lpAddress As IntPtr, dwSize As UInteger, dwFreeType As UInteger) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function ReadProcessMemory(hProcess As IntPtr, lpBaseAddress As IntPtr, lpBuffer As IntPtr, nSize As Integer, ByRef lpNumberOfBytesRead As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessageTimeout(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr,
                                                fuFlags As UInteger, uTimeout As UInteger, ByRef lpdwResult As IntPtr) As IntPtr
    End Function

    Private Const TB_BUTTONCOUNT As UInteger = &H418
    Private Const TB_GETBUTTON As UInteger = &H417
    Private Const TB_GETBUTTONINFO As UInteger = &H441

    Private Const PROCESS_VM_OPERATION As UInteger = &H8
    Private Const PROCESS_VM_READ As UInteger = &H10
    Private Const PROCESS_VM_WRITE As UInteger = &H20
    Private Const MEM_COMMIT As UInteger = &H1000
    Private Const MEM_RELEASE As UInteger = &H8000
    Private Const PAGE_READWRITE As UInteger = &H4

    ' TBBUTTON structure size for 64-bit
    Private Const TBBUTTON_SIZE_64 As Integer = 32

    <StructLayout(LayoutKind.Sequential)>
    Private Structure TBBUTTON64
        Public iBitmap As Integer
        Public idCommand As Integer
        Public fsState As Byte
        Public fsStyle As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)>
        Public bReserved() As Byte
        Public dwData As Long   ' 64-bit pointer
        Public iString As Long  ' 64-bit pointer
    End Structure

    Private Structure TrayIconInfo
        Public Icon As Bitmap
        Public Tooltip As String
        Public ProcessName As String
        Public Hwnd As IntPtr
        Public CallbackMsg As UInteger
        Public Id As UInteger
    End Structure

    Private Sub BuildCustomTray()
        panelTray.Controls.Clear()
        trayIconButtons.Clear()

        ' 1. Add system status icons (these always work)
        AddSystemStatusIcons()

        ' 2. Try to enumerate real tray notification icons from the toolbar
        Try
            Dim trayIcons = EnumerateTrayIcons()
            For Each iconInfo In trayIcons
                If iconInfo.Icon IsNot Nothing Then
                    AddTrayIconButton(iconInfo)
                End If
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Tray icon enumeration failed: " & ex.Message)
        End Try

        ' 3. Set up periodic refresh (every 5 seconds)
        If trayRefreshTimer Is Nothing Then
            trayRefreshTimer = New Timer() With {.Interval = 5000}
            AddHandler trayRefreshTimer.Tick, Sub(s, ev)
                                                   Try
                                                       RefreshSystemStatus()
                                                   Catch
                                                   End Try
                                               End Sub
            trayRefreshTimer.Start()
        End If
    End Sub

    Private Sub AddSystemStatusIcons()
        ' Network status
        Dim netConnected = Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()
        Dim netIcon = CreateStatusIcon(If(netConnected, "🌐", "❌"), If(netConnected, Color.FromArgb(96, 165, 250), Color.Gray))
        Dim netPb = AddIconToTray(netIcon, If(netConnected, "Network: Connected", "Network: Disconnected"))
        netPb.Name = "sysNet"
        AddHandler netPb.Click, Sub(s, ev)
                                     Try : Process.Start(New ProcessStartInfo("ms-settings:network") With {.UseShellExecute = True}) : Catch : End Try
                                 End Sub

        ' Volume — click opens volume mixer
        Dim volIcon = CreateStatusIcon("🔊", Color.FromArgb(200, 200, 200))
        Dim volPb = AddIconToTray(volIcon, "Volume")
        volPb.Name = "sysVol"
        AddHandler volPb.Click, Sub(s, ev)
                                     Try : Process.Start(New ProcessStartInfo("sndvol.exe") With {.UseShellExecute = True}) : Catch : End Try
                                 End Sub

        ' Battery (only show if system has one)
        Dim ps = SystemInformation.PowerStatus
        If ps.BatteryChargeStatus <> BatteryChargeStatus.NoSystemBattery Then
            Dim pct = CInt(ps.BatteryLifePercent * 100)
            Dim charging = ps.PowerLineStatus = PowerLineStatus.Online
            Dim batSymbol = If(charging, "🔌", If(pct > 50, "🔋", If(pct > 20, "🪫", "⚠")))
            Dim batColor = If(pct > 50, Color.FromArgb(74, 222, 128), If(pct > 20, Color.FromArgb(250, 204, 21), Color.FromArgb(248, 113, 113)))
            Dim batIcon = CreateStatusIcon(batSymbol, batColor)
            Dim batPb = AddIconToTray(batIcon, $"Battery: {pct}%" & If(charging, " (Charging)", ""))
            batPb.Name = "sysBat"
            AddHandler batPb.Click, Sub(s, ev)
                                         Try : Process.Start(New ProcessStartInfo("ms-settings:batterysaver") With {.UseShellExecute = True}) : Catch : End Try
                                     End Sub
        End If
    End Sub

    Private Function CreateStatusIcon(symbol As String, clr As Color) As Bitmap
        Dim bmp As New Bitmap(24, 24)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
            Using fnt As New Font("Segoe UI Emoji", 14, FontStyle.Regular)
                Using br As New SolidBrush(clr)
                    Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(symbol, fnt, br, New RectangleF(0, 0, 24, 24), sf)
                End Using
            End Using
        End Using
        Return bmp
    End Function

    Private Function AddIconToTray(icon As Bitmap, tooltip As String) As PictureBox
        Dim pb As New PictureBox() With {
            .Size = New Size(32, 36),
            .SizeMode = PictureBoxSizeMode.CenterImage,
            .Image = icon,
            .Margin = New Padding(2),
            .Cursor = Cursors.Hand,
            .BackColor = Color.Transparent
        }
        trayToolTip.SetToolTip(pb, tooltip)
        AddHandler pb.MouseEnter, Sub(s, ev) pb.BackColor = Color.FromArgb(40, 255, 255, 255)
        AddHandler pb.MouseLeave, Sub(s, ev) pb.BackColor = Color.Transparent
        panelTray.Controls.Add(pb)
        trayIconButtons.Add(pb)
        Return pb
    End Function

    Private Sub AddTrayIconButton(info As TrayIconInfo)
        Dim pb = AddIconToTray(info.Icon, If(String.IsNullOrEmpty(info.Tooltip), info.ProcessName, info.Tooltip))
        ' Forward click to the original tray icon's window
        If info.Hwnd <> IntPtr.Zero AndAlso info.CallbackMsg > 0 Then
            AddHandler pb.Click, Sub(s, ev)
                                      Try
                                          ' Send WM_LBUTTONUP to the icon's owner window
                                          Dim WM_LBUTTONUP As UInteger = &H202
                                          SendMessage(info.Hwnd, info.CallbackMsg, New IntPtr(info.Id), New IntPtr(WM_LBUTTONUP))
                                      Catch
                                      End Try
                                  End Sub
        End If
    End Sub

    ' --- ITrayNotify COM interfaces (undocumented Shell internal API) ---
    ' CLSID for the Shell TrayNotify object
    Private Shared ReadOnly CLSID_TrayNotify As New Guid("25DEAD04-1EAC-4911-9E3A-AD0A4AB560FD")

    ' ITrayNotify (Windows 8+) — Shell's internal interface for managing notification icons
    <ComImport(), Guid("FB852B2C-6BAD-4605-9551-F15F87830935"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface ITrayNotify
        Sub RegisterCallback(<MarshalAs(UnmanagedType.Interface)> callback As INotificationCB, <Out> ByRef token As UInteger)
        Sub UnregisterCallback(token As UInteger)
        Sub SetPreference(pNotifyItem As IntPtr)
        Sub EnableAutoTray(<MarshalAs(UnmanagedType.Bool)> enabled As Boolean)
    End Interface

    ' INotificationCB — callback interface that receives tray icon notifications
    <ComImport(), Guid("D782CCBA-AFB0-43F1-94DB-FDA3779EACCB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface INotificationCB
        Sub Notify(dwEvent As UInteger, pNotifyItem As IntPtr)
    End Interface

    ' Our implementation of INotificationCB that collects all tray icons
    <ComVisible(True), ClassInterface(ClassInterfaceType.None)>
    Private Class TrayIconCollector
        Implements INotificationCB

        Public ReadOnly Icons As New List(Of TrayIconInfo)()

        Public Sub Notify(dwEvent As UInteger, pNotifyItem As IntPtr) Implements INotificationCB.Notify
            If pNotifyItem = IntPtr.Zero Then Return
            Try
                ' NOTIFYITEM layout on 64-bit Windows 8+:
                '   Offset  0: PWSTR pszExeName  (8 bytes - pointer)
                '   Offset  8: PWSTR pszTip      (8 bytes - pointer)
                '   Offset 16: HICON hIcon       (8 bytes)
                '   Offset 24: HWND  hWnd        (8 bytes)
                '   Offset 32: DWORD dwPreference(4 bytes) — 0=show notifications only, 1=hide, 2=show
                '   Offset 36: UINT  uID         (4 bytes)
                '   Offset 40: GUID  guidItem    (16 bytes)
                Dim pExeName = Marshal.ReadIntPtr(pNotifyItem, 0)
                Dim pTip = Marshal.ReadIntPtr(pNotifyItem, 8)
                Dim hIcon = Marshal.ReadIntPtr(pNotifyItem, 16)
                Dim hWnd = Marshal.ReadIntPtr(pNotifyItem, 24)
                Dim preference = Marshal.ReadInt32(pNotifyItem, 32)
                Dim uID = CUInt(Marshal.ReadInt32(pNotifyItem, 36))

                Dim exeName As String = Nothing
                If pExeName <> IntPtr.Zero Then exeName = Marshal.PtrToStringUni(pExeName)

                Dim tip As String = Nothing
                If pTip <> IntPtr.Zero Then tip = Marshal.PtrToStringUni(pTip)

                Dim info As New TrayIconInfo()
                info.Hwnd = hWnd
                info.Id = uID
                info.CallbackMsg = 0
                info.Tooltip = If(Not String.IsNullOrEmpty(tip), tip, "")
                info.ProcessName = If(Not String.IsNullOrEmpty(exeName), IO.Path.GetFileName(exeName), "")

                ' Extract the icon bitmap from the HICON handle
                If hIcon <> IntPtr.Zero Then
                    Try
                        Using ico = Icon.FromHandle(hIcon)
                            info.Icon = New Bitmap(ico.ToBitmap(), 24, 24)
                        End Using
                    Catch
                    End Try
                End If

                ' If no icon from HICON, try extracting from the exe path
                If info.Icon Is Nothing AndAlso Not String.IsNullOrEmpty(exeName) AndAlso IO.File.Exists(exeName) Then
                    Try
                        Dim ico = Icon.ExtractAssociatedIcon(exeName)
                        If ico IsNot Nothing Then
                            info.Icon = New Bitmap(ico.ToBitmap(), 24, 24)
                            ico.Dispose()
                        End If
                    Catch
                    End Try
                End If

                ' Last resort: text icon from first letter
                If info.Icon Is Nothing Then
                    Dim label = If(Not String.IsNullOrEmpty(info.ProcessName), info.ProcessName,
                                If(Not String.IsNullOrEmpty(info.Tooltip), info.Tooltip, "?"))
                    Dim bmp As New Bitmap(24, 24)
                    Using g = Graphics.FromImage(bmp)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit
                        Using fnt As New Font("Segoe UI", 10, FontStyle.Bold)
                            Using br As New SolidBrush(Color.FromArgb(200, 200, 200))
                                Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                                g.DrawString(label.Substring(0, 1).ToUpper(), fnt, br, New RectangleF(0, 0, 24, 24), sf)
                            End Using
                        End Using
                    End Using
                    info.Icon = bmp
                End If

                Icons.Add(info)

                System.Diagnostics.Debug.WriteLine($"  TrayIcon: exe={exeName}, tip={tip}, pref={preference}, hIcon={hIcon}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"  TrayIcon parse error: {ex.Message}")
            End Try
        End Sub
    End Class

    Private Function EnumerateTrayIcons() As List(Of TrayIconInfo)
        Dim result As New List(Of TrayIconInfo)()

        Try
            ' Create the Shell TrayNotify COM object
            Dim trayNotifyType = Type.GetTypeFromCLSID(CLSID_TrayNotify)
            If trayNotifyType Is Nothing Then
                System.Diagnostics.Debug.WriteLine("ITrayNotify: CLSID not found")
                Return result
            End If

            Dim trayNotifyObj = Activator.CreateInstance(trayNotifyType)
            Dim trayNotify = DirectCast(trayNotifyObj, ITrayNotify)

            ' Create our callback collector
            Dim collector As New TrayIconCollector()

            ' Register the callback — Shell will call Notify() synchronously
            ' for every existing notification icon
            Dim token As UInteger = 0
            trayNotify.RegisterCallback(collector, token)

            ' Immediately unregister — we only needed the initial enumeration
            Try
                trayNotify.UnregisterCallback(token)
            Catch
            End Try

            ' Release COM object
            Try
                Marshal.ReleaseComObject(trayNotify)
            Catch
            End Try

            result = collector.Icons
            System.Diagnostics.Debug.WriteLine($"ITrayNotify enumerated {result.Count} tray icon(s)")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"ITrayNotify failed: {ex.Message}")
        End Try

        Return result
    End Function

    Private Function CreateTextIcon(text As String) As Bitmap
        Dim bmp As New Bitmap(24, 24)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
            Using fnt As New Font("Segoe UI", 10, FontStyle.Bold)
                Using br As New SolidBrush(Color.FromArgb(200, 200, 200))
                    Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(text, fnt, br, New RectangleF(0, 0, 24, 24), sf)
                End Using
            End Using
        End Using
        Return bmp
    End Function

    Private Sub RefreshSystemStatus()
        If panelTray Is Nothing OrElse Not panelTray.IsHandleCreated Then Return

        ' Update network icon
        Try
            Dim netPb = panelTray.Controls.Find("sysNet", False).FirstOrDefault()
            If netPb IsNot Nothing Then
                Dim connected = Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()
                Dim oldImg = netPb.BackgroundImage
                DirectCast(netPb, PictureBox).Image = CreateStatusIcon(
                    If(connected, "🌐", "❌"),
                    If(connected, Color.FromArgb(96, 165, 250), Color.Gray))
                trayToolTip.SetToolTip(netPb, If(connected, "Network: Connected", "Network: Disconnected"))
            End If
        Catch
        End Try

        ' Update battery icon
        Try
            Dim batPb = panelTray.Controls.Find("sysBat", False).FirstOrDefault()
            If batPb IsNot Nothing Then
                Dim ps = SystemInformation.PowerStatus
                Dim pct = CInt(ps.BatteryLifePercent * 100)
                Dim charging = ps.PowerLineStatus = PowerLineStatus.Online
                Dim batSymbol = If(charging, "🔌", If(pct > 50, "🔋", If(pct > 20, "🪫", "⚠")))
                Dim batColor = If(pct > 50, Color.FromArgb(74, 222, 128), If(pct > 20, Color.FromArgb(250, 204, 21), Color.FromArgb(248, 113, 113)))
                DirectCast(batPb, PictureBox).Image = CreateStatusIcon(batSymbol, batColor)
                trayToolTip.SetToolTip(batPb, $"Battery: {pct}%" & If(charging, " (Charging)", ""))
            End If
        Catch
        End Try
    End Sub
#End Region

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
                        btn.Size = New Size(36, 36)
                        btn.BackgroundImage = New Bitmap(icon.ToBitmap(), 24, 24)
                        btn.BackgroundImageLayout = ImageLayout.Center
                        btn.Tag = f
                        btn.Margin = New Padding(2)
                        Dim appName = IO.Path.GetFileNameWithoutExtension(f)
                        ApplyFlatStyle(btn, appName)
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

    ''' <summary>
    ''' Applies the flat tray-icon style to any button: no border, transparent background,
    ''' subtle hover highlight, hand cursor.
    ''' </summary>
    Private Sub ApplyFlatStyle(btn As Button, Optional tooltipText As String = Nothing)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.BorderColor = Color.FromArgb(0, 0, 0, 0)
        If isDarkMode Then
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 255, 255, 255)
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(60, 255, 255, 255)
            btn.ForeColor = Color.FromArgb(240, 240, 240)
        Else
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 0, 0, 0)
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(60, 0, 0, 0)
            btn.ForeColor = Color.FromArgb(30, 30, 30)
        End If
        btn.BackColor = Color.Transparent
        btn.Cursor = Cursors.Hand
        If Not String.IsNullOrEmpty(tooltipText) Then
            trayToolTip.SetToolTip(btn, tooltipText)
        End If
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
            btn.Size = New Size(36, 36)
            If icon IsNot Nothing Then
                btn.BackgroundImage = New Bitmap(icon.ToBitmap(), 24, 24)
                btn.BackgroundImageLayout = ImageLayout.Center
            Else
                btn.Text = p.ProcessName.Substring(0, Math.Min(3, p.ProcessName.Length))
                btn.Font = New Font("Segoe UI", 7, FontStyle.Regular)
            End If
            Dim pid = p.Id
            btn.Tag = "running:" & pid
            btn.Margin = New Padding(2)
            ApplyFlatStyle(btn, p.ProcessName)
            AddHandler btn.Click, Sub(s, ev)
                                      Try
                                          Dim proc = Process.GetProcessById(pid)
                                          Dim hWnd = proc.MainWindowHandle
                                          If hWnd = IntPtr.Zero Then Return

                                          If NativeMethods.IsIconic(hWnd) Then
                                              ' Minimized → restore
                                              NativeMethods.ShowWindow(hWnd, 9) ' SW_RESTORE
                                              NativeMethods.SetForegroundWindow(hWnd)
                                              lastActivatedPid = pid
                                          ElseIf lastActivatedPid = pid Then
                                              ' Same app clicked again → minimize (toggle)
                                              NativeMethods.ShowWindow(hWnd, 6) ' SW_MINIMIZE
                                              lastActivatedPid = -1
                                          Else
                                              ' Different app → bring to front
                                              NativeMethods.SetForegroundWindow(hWnd)
                                              lastActivatedPid = pid
                                          End If
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

        <DllImport("user32.dll")>
        Public Shared Function IsIconic(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function
    End Class

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' Restore embedded system tray back to Windows before closing
        RestoreSystemTray()

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
        ' Hide the title bar that WS_CAPTION adds — DWM needs the caption for backdrop
        If m.Msg = WM_NCCALCSIZE AndAlso m.WParam <> IntPtr.Zero Then
            m.Result = IntPtr.Zero
            Return
        End If

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
