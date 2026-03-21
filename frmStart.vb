Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Drawing.Text

' search for TrayNotifyWnd to find the system tray? 
Public Class frmStart
    Inherits Form

#Region "P/Invoke - Mica / Acrylic / DWM"
    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmExtendFrameIntoClientArea(hwnd As IntPtr, ByRef margins As MARGINS) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure MARGINS
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
    Private Const DWMWA_SYSTEMBACKDROP_TYPE As Integer = 38
    Private Const DWMWA_MICA_EFFECT As Integer = 1029

    ' Backdrop types for DWMWA_SYSTEMBACKDROP_TYPE (Win11 22H2+)
    Private Const DWMSBT_MAINWINDOW As Integer = 2   ' Mica
    Private Const DWMSBT_TABBEDWINDOW As Integer = 4 ' Mica Alt
    Private Const DWMSBT_TRANSIENTWINDOW As Integer = 3 ' Acrylic

    ' Window style constants for Mica on borderless
    Private Const WS_CAPTION As Integer = &HC00000
    Private Const WS_THICKFRAME As Integer = &H40000
    Private Const WM_NCCALCSIZE As Integer = &H83
    Private Const WM_NCHITTEST As Integer = &H84
    Private Const GWL_STYLE As Integer = -16
    Private Const SWP_FRAMECHANGED As UInteger = &H20
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4

    <DllImport("user32.dll")>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    Private micaApplied As Boolean = False
    Private micaEnabled As Boolean = True ' False  ' Off by default — user can toggle
#End Region

#Region "P/Invoke - Shell icon extraction"
    <StructLayout(LayoutKind.Sequential)>
    Private Structure SHSIZE
        Public cx As Integer
        Public cy As Integer
        Public Sub New(w As Integer, h As Integer)
            cx = w
            cy = h
        End Sub
    End Structure

    <ComImport(), Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IShellItemImageFactory
        <PreserveSig>
        Function GetImage(ByVal size As SHSIZE, ByVal flags As UInteger, ByRef phbm As IntPtr) As Integer
    End Interface

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function SHCreateItemFromParsingName(pszPath As String, pbc As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer
    End Function

    <DllImport("gdi32.dll")>
    Private Shared Function DeleteObject(hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("gdi32.dll", EntryPoint:="GetObjectW")>
    Private Shared Function GetBitmapObject(hObject As IntPtr, nCount As Integer, ByRef lpObject As BITMAPINFO_NATIVE) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure BITMAPINFO_NATIVE
        Public bmType As Integer
        Public bmWidth As Integer
        Public bmHeight As Integer
        Public bmWidthBytes As Integer
        Public bmPlanes As Short
        Public bmBitsPixel As Short
        Public bmBits As IntPtr
    End Structure

    <Flags>
    Public Enum SIIGBF As Integer
        RESIZETOFIT = &H0
        BIGGERSIZEOK = &H1
        MEMORYONLY = &H2
        ICONONLY = &H4
        THUMBNAILONLY = &H8
        INCACHEONLY = &H10
        SCALEUP = &H100 ' Essential for Jumbo quality
    End Enum




#End Region

#Region "Data model"
    Private Class AppEntry
        Public Property Name As String
        Public Property Key As String
        Public Property ShortcutPath As String
        Public Property Group As String
        Public Property Category As String = ""
    End Class

    Private allEntries As New List(Of AppEntry)()
    Private imgDict As New Dictionary(Of String, Bitmap)(StringComparer.OrdinalIgnoreCase)
    Private Const ICON_SIZE As Integer = 32
    Private Const LARGE_ICON_SIZE As Integer = 40
#End Region

#Region "UI Controls"
    Private txtSearch As TextBox
    Private pnlPinned As Panel
    Private pnlCategories As Panel
    Private pnlAllApps As Panel
    Private pnlUserBar As Panel
    Private scrollAllApps As VScrollBar
    Private allAppsScrollOffset As Integer = 0
    Private hoveredCard As Rectangle = Rectangle.Empty
    Private hoveredCardEntry As AppEntry = Nothing
    Private pinnedHoveredIdx As Integer = -1
    Private catHoveredIdx As Integer = -1
    Private allAppsHoveredIdx As Integer = -1
    Private showLetterGrid As Boolean = False
    Private letterGridHoveredIdx As Integer = -1
    Private expandedCategories As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private hoveredCatHeader As String = Nothing
    Private catScrollOffsets As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private Const CAT_MAX_EXPANDED_ROWS As Integer = 3  ' max visible rows when expanded before scrolling

    ' Drag and drop state
    Private dragStartPoint As Point = Point.Empty
    Private isDragging As Boolean = False
    Private dragDropTargetCat As String = Nothing  ' highlighted category during drag-over
    Private dragDropNewTarget As Boolean = False    ' hovering over "new category" drop zone
    Private dragSourceCat As String = Nothing       ' category the dragged item came from
    Private catDragStartPoint As Point = Point.Empty

    ' Layout constants
    Private Const FORM_WIDTH As Integer = 860
    Private Const FORM_HEIGHT As Integer = 640
    Private Const CARD_W As Integer = 88
    Private Const CARD_H As Integer = 88
    Private Const CARD_GAP As Integer = 8
    Private Const SECTION_PAD As Integer = 20
    Private Const ALLAPPS_WIDTH As Integer = 240
    Private Const ALLAPPS_ROW_H As Integer = 34
    Private Const ALLAPPS_LETTER_H As Integer = 28
    Private Const CAT_CARD_W As Integer = 72
    Private Const CAT_CARD_H As Integer = 72
    ' Computed available width for left content area (pinned + categories)
    ' = FormWidth - AllAppsWidth - divider - leftPadding - rightPadding - scrollbar margin
    Private Const LEFT_CONTENT_W As Integer = FORM_WIDTH - ALLAPPS_WIDTH - 1 - SECTION_PAD - 10 - 20

    ' Colors
    Private ReadOnly clrBackground As Color = Color.FromArgb(220, 32, 32, 32)
    'Private ReadOnly clrBackground As Color = Color.FromArgb(220, 30, 80, 180) ' for debugging

    Private ReadOnly clrSurface As Color = Color.FromArgb(180, 48, 48, 48)
    Private ReadOnly clrSurfaceHover As Color = Color.FromArgb(200, 70, 70, 70)
    Private ReadOnly clrAccent As Color = Color.FromArgb(255, 96, 165, 250)
    Private ReadOnly clrTextPrimary As Color = Color.FromArgb(255, 240, 240, 240)
    Private ReadOnly clrTextSecondary As Color = Color.FromArgb(180, 200, 200, 200)
    Private ReadOnly clrDivider As Color = Color.FromArgb(60, 255, 255, 255)
    Private ReadOnly clrSearchBg As Color = Color.FromArgb(160, 60, 60, 60)
    Private ReadOnly clrCardBg As Color = Color.FromArgb(120, 55, 55, 55)
    Private ReadOnly clrCardHover As Color = Color.FromArgb(160, 80, 80, 80)
    Private ReadOnly clrLetterBg As Color = Color.FromArgb(80, 96, 165, 250)
    Private ReadOnly clrGroupBorder As Color = Color.FromArgb(50, 255, 255, 255)
    Private ReadOnly clrGroupBg As Color = Color.FromArgb(40, 255, 255, 255)

    ' Pinned entries (max ~12)
    Private pinnedEntries As New List(Of AppEntry)()
    ' Category buckets (custom categories first, then auto-categories)
    Private categories As New List(Of (Name As String, Entries As List(Of AppEntry)))()
    ' Custom category names — used to identify which categories are user-created
    Private customCategoryNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly clrCustomCatBorder As Color = Color.FromArgb(80, 96, 165, 250)
#End Region

    Public Sub New()
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(FORM_WIDTH, FORM_HEIGHT)
        Me.BackColor = Color.FromArgb(255, 30, 30, 30)
        Me.ShowInTaskbar = False
        Me.KeyPreview = True

        AddHandler Me.KeyDown, AddressOf Form_KeyDown
        AddHandler Me.Deactivate, AddressOf Form_Deactivate

        BuildUI()
        LoadStartMenuApps()
        BuildPinned()
        LoadSavedPins()
        BuildCategories()
    End Sub

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        If micaEnabled Then ApplyMicaBackdrop()
    End Sub

    Private Sub ApplyMicaBackdrop()
        Try
            ' Enable dark mode for window chrome
            Dim darkMode As Integer = 1
            DwmSetWindowAttribute(Me.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, darkMode, 4)

            ' Extend frame into entire client area — must be done BEFORE setting backdrop
            Dim m As New MARGINS() With {.Left = -1, .Right = -1, .Top = -1, .Bottom = -1}
            DwmExtendFrameIntoClientArea(Me.Handle, m)

            ' Try Mica first (Win11 22H2+)
            Dim backdropType As Integer = DWMSBT_MAINWINDOW
            Dim hr = DwmSetWindowAttribute(Me.Handle, DWMWA_SYSTEMBACKDROP_TYPE, backdropType, 4)

            If hr <> 0 Then
                ' Fallback: try the older Mica attribute (Win11 21H2)
                Dim micaOn As Integer = 1
                hr = DwmSetWindowAttribute(Me.Handle, DWMWA_MICA_EFFECT, micaOn, 4)
            End If

            ' If either succeeded, mark as applied so we don't paint over it
            micaApplied = True

            ' Force a
            Me.Invalidate(True)
        Catch
            ' Mica not available — fall back to solid dark
            micaApplied = False
        End Try
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp = MyBase.CreateParams
            ' Only add WS_CAPTION when Mica is active — DWM needs it for backdrop.
            ' WM_NCCALCSIZE hides the title bar visually.
            If micaEnabled Then
                cp.Style = cp.Style Or WS_CAPTION
            End If
            ' Reduce flicker
            cp.ExStyle = cp.ExStyle Or &H2000000 ' WS_EX_COMPOSITED
            Return cp
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_NCCALCSIZE AndAlso m.WParam <> IntPtr.Zero Then
            ' Always remove the non-client area (title bar) while keeping WS_CAPTION
            m.Result = IntPtr.Zero
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form_Deactivate(sender As Object, e As EventArgs)
        ' Close when user clicks away, like a real start menu
        Me.Close()
    End Sub

    Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Escape Then
            If showLetterGrid Then
                showLetterGrid = False
                letterGridHoveredIdx = -1
                pnlAllApps.Invalidate()
                e.Handled = True
                Return
            End If
            Me.Close()
        ElseIf Not txtSearch.Focused Then
            ' Redirect typing to search
            txtSearch.Focus()
        End If
    End Sub

#Region "UI Construction"
    Private Sub BuildUI()
        ' Top user bar with avatar, search, power
        pnlUserBar = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .BackColor = Color.Transparent
        }
        AddHandler pnlUserBar.Paint, AddressOf PaintUserBar
        AddHandler pnlUserBar.MouseClick, AddressOf UserBarClick

        ' Search box inside user bar — we overlay it
        txtSearch = New TextBox() With {
            .Font = New Font("Segoe UI", 11, FontStyle.Regular),
            .ForeColor = clrTextPrimary,
            .BackColor = Color.FromArgb(255, 50, 50, 50),
            .BorderStyle = BorderStyle.None,
            .Width = 320,
            .Height = 28
        }
        txtSearch.Location = New Point((FORM_WIDTH - txtSearch.Width) \ 2, 14)
        pnlUserBar.Controls.Add(txtSearch)
        AddHandler txtSearch.TextChanged, AddressOf OnSearchTextChanged

        ' Main content panel
        Dim pnlMain As New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .Padding = New Padding(0)
        }

        ' Add Fill panel FIRST, then Top panel — WinForms docks higher Z-order first,
        ' so pnlUserBar (added last) claims its 56px at top, then pnlMain fills the rest.
        Me.Controls.Add(pnlMain)
        Me.Controls.Add(pnlUserBar)

        ' All Apps panel on the right
        pnlAllApps = New Panel() With {
            .Dock = DockStyle.Right,
            .Width = ALLAPPS_WIDTH,
            .BackColor = Color.Transparent
        }
        AddHandler pnlAllApps.Paint, AddressOf PaintAllApps
        AddHandler pnlAllApps.MouseClick, AddressOf AllAppsClick
        AddHandler pnlAllApps.MouseDown, AddressOf AllAppsMouseDown
        AddHandler pnlAllApps.MouseMove, AddressOf AllAppsMouseMove
        AddHandler pnlAllApps.MouseLeave, AddressOf AllAppsMouseLeave
        AddHandler pnlAllApps.MouseWheel, AddressOf AllAppsWheel
        AddHandler pnlAllApps.MouseUp, AddressOf AllAppsRightClick
        pnlMain.Controls.Add(pnlAllApps)

        ' Divider line
        Dim divider As New Panel() With {
            .Dock = DockStyle.Right,
            .Width = 1,
            .BackColor = clrDivider
        }
        pnlMain.Controls.Add(divider)

        ' Left area: scrollable panel containing pinned + categories
        Dim pnlLeft As New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .AutoScroll = True,
            .Padding = New Padding(SECTION_PAD, 10, 10, 10)
        }
        pnlMain.Controls.Add(pnlLeft)

        ' Categories below pinned (add SECOND — Dock.Top stacks in reverse add order)
        pnlCategories = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 300,
            .BackColor = Color.Transparent,
            .AllowDrop = True
        }
        AddHandler pnlCategories.Paint, AddressOf PaintCategories
        AddHandler pnlCategories.MouseClick, AddressOf CategoriesClick
        AddHandler pnlCategories.MouseDown, AddressOf CategoriesMouseDown
        AddHandler pnlCategories.MouseMove, AddressOf CategoriesMouseMove
        AddHandler pnlCategories.MouseLeave, AddressOf CategoriesMouseLeave
        AddHandler pnlCategories.MouseUp, AddressOf CategoriesRightClick
        AddHandler pnlCategories.DragEnter, AddressOf CategoriesDragEnter
        AddHandler pnlCategories.DragOver, AddressOf CategoriesDragOver
        AddHandler pnlCategories.DragLeave, AddressOf CategoriesDragLeave
        AddHandler pnlCategories.DragDrop, AddressOf CategoriesDragDrop
        ' Scrolling within expanded categories is handled by click on the scrollbar area
        pnlLeft.Controls.Add(pnlCategories)

        ' Pinned section at top (add LAST so it docks above categories)
        pnlPinned = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 260,
            .BackColor = Color.Transparent
        }
        AddHandler pnlPinned.Paint, AddressOf PaintPinned
        AddHandler pnlPinned.MouseClick, AddressOf PinnedClick
        AddHandler pnlPinned.MouseMove, AddressOf PinnedMouseMove
        AddHandler pnlPinned.MouseLeave, AddressOf PinnedMouseLeave
        pnlLeft.Controls.Add(pnlPinned)
    End Sub
#End Region

#Region "Data Loading"
    Private Sub LoadStartMenuApps()
        Dim roots As New List(Of String)()
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))

        Dim files As New List(Of String)()
        For Each root In roots.Distinct(StringComparer.OrdinalIgnoreCase)
            Try
                If String.IsNullOrEmpty(root) Then Continue For
                If Directory.Exists(root) Then
                    Try
                        files.AddRange(Directory.GetFiles(root, "*.lnk", SearchOption.AllDirectories))
                    Catch
                    End Try
                End If
            Catch
            End Try
        Next

        Dim groupRoots = New String() {
            Environment.GetFolderPath(Environment.SpecialFolder.Programs),
            Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms),
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
            Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
        }.Where(Function(r) Not String.IsNullOrEmpty(r)).OrderByDescending(Function(r) r.Length).ToArray()

        Dim resolved As New Dictionary(Of String, AppEntry)(StringComparer.OrdinalIgnoreCase)
        For Each f In files
            Try
                Dim shell = CreateObject("WScript.Shell")
                Dim lnk = shell.CreateShortcut(f)
                Dim target = If(lnk.TargetPath, String.Empty)
                Dim args = If(lnk.Arguments, String.Empty)

                Dim display = Path.GetFileNameWithoutExtension(f)
                If String.IsNullOrEmpty(display) AndAlso Not String.IsNullOrEmpty(target) Then
                    display = Path.GetFileNameWithoutExtension(target)
                End If
                If String.IsNullOrEmpty(display) Then Continue For

                Dim key As String = String.Empty
                If Not String.IsNullOrEmpty(target) Then
                    Dim tname = Path.GetFileName(target).ToLowerInvariant()
                    If tname = "explorer.exe" AndAlso Not String.IsNullOrEmpty(args) AndAlso args.IndexOf("appsfolder", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        key = args.Trim()
                    Else
                        Try : key = Path.GetFullPath(target) : Catch : key = target : End Try
                    End If
                ElseIf Not String.IsNullOrEmpty(args) AndAlso args.IndexOf("appsfolder", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    key = args.Trim()
                End If
                If String.IsNullOrEmpty(key) Then key = f

                Dim groupName = "Programs"
                Try
                    Dim dir = Path.GetDirectoryName(f)
                    For Each r In groupRoots
                        If dir.StartsWith(r, StringComparison.OrdinalIgnoreCase) Then
                            Dim rel = dir.Substring(r.Length).TrimStart(Path.DirectorySeparatorChar)
                            If Not String.IsNullOrEmpty(rel) Then groupName = rel.Replace(Path.DirectorySeparatorChar, " "c)
                            Exit For
                        End If
                    Next
                Catch
                End Try

                If resolved.ContainsKey(key) Then
                    Dim existing = resolved(key)
                    Dim userStart = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
                    Dim existingIsUser = Not String.IsNullOrEmpty(existing.ShortcutPath) AndAlso existing.ShortcutPath.StartsWith(userStart, StringComparison.OrdinalIgnoreCase)
                    Dim newIsUser = f.StartsWith(userStart, StringComparison.OrdinalIgnoreCase)
                    If newIsUser AndAlso Not existingIsUser Then
                        resolved(key) = New AppEntry() With {.Name = display, .Key = key, .ShortcutPath = f, .Group = groupName}
                    End If
                Else
                    resolved(key) = New AppEntry() With {.Name = display, .Key = key, .ShortcutPath = f, .Group = groupName}
                End If
            Catch
            End Try
        Next

        ' Enumerate Store/UWP apps via shell:AppsFolder
        Try
            Dim shellApp = CreateObject("Shell.Application")
            Dim appsFolder = shellApp.NameSpace("shell:AppsFolder")
            If appsFolder IsNot Nothing Then
                For Each item In appsFolder.Items()
                    Try
                        Dim appName As String = item.Name
                        Dim appPath As String = item.Path
                        If String.IsNullOrEmpty(appName) OrElse String.IsNullOrEmpty(appPath) Then Continue For
                        If resolved.ContainsKey(appPath) Then Continue For
                        Dim launchPath = "shell:AppsFolder\" & appPath
                        resolved(appPath) = New AppEntry() With {
                            .Name = appName, .Key = appPath, .ShortcutPath = launchPath, .Group = "Apps"
                        }
                    Catch
                    End Try
                Next
            End If
        Catch
        End Try

        ' Deduplicate by display name, build final list
        allEntries.Clear()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each entry In resolved.Values.OrderBy(Function(x) x.Name)
            If seen.Contains(entry.Name) Then Continue For
            seen.Add(entry.Name)
            ' Auto-categorize
            entry.Category = CategorizeApp(entry)
            allEntries.Add(entry)
            ' Load icon
            LoadEntryIcon(entry)
        Next
    End Sub

    'Private Sub LoadEntryIcon(entry As AppEntry)
    '    If imgDict.ContainsKey(entry.Key) Then Return
    '    Try
    '        ' 1) Try exe path directly
    '        If File.Exists(entry.Key) Then
    '            Dim ico = Icon.ExtractAssociatedIcon(entry.Key)
    '            If ico IsNot Nothing Then
    '                imgDict(entry.Key) = ResizeIcon(ico.ToBitmap())
    '                Return
    '            End If
    '        End If
    '        ' 2) Try shortcut file
    '        If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso entry.ShortcutPath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) AndAlso File.Exists(entry.ShortcutPath) Then
    '            Dim ico = Icon.ExtractAssociatedIcon(entry.ShortcutPath)
    '            If ico IsNot Nothing Then
    '                imgDict(entry.Key) = ResizeIcon(ico.ToBitmap())
    '                Return
    '            End If
    '        End If
    '        ' 3) Shell image factory (UWP/Store)
    '        Dim parsingName As String = Nothing
    '        If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso entry.ShortcutPath.StartsWith("shell:AppsFolder", StringComparison.OrdinalIgnoreCase) Then
    '            parsingName = entry.ShortcutPath
    '        Else
    '            parsingName = "shell:AppsFolder\" & entry.Key
    '        End If
    '        Dim bmpShell = GetShellIconBitmap(parsingName, LARGE_ICON_SIZE)
    '        If bmpShell IsNot Nothing Then
    '            imgDict(entry.Key) = bmpShell
    '            Return
    '        End If
    '    Catch
    '    End Try
    '    ' Fallback placeholder
    '    imgDict(entry.Key) = CreatePlaceholder(entry.Name)
    'End
    '
    'Private Sub LoadEntryIcon(entry As AppEntry)
    '    If imgDict.ContainsKey(entry.Key) Then Return

    '    ' Unified path: Try Shell Image Factory first (works for EXE, LNK, and UWP)
    '    Dim parsingName As String = ""
    '    If entry.ShortcutPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) Then
    '        parsingName = entry.ShortcutPath
    '    ElseIf File.Exists(entry.ShortcutPath) Then
    '        parsingName = entry.ShortcutPath
    '    ElseIf File.Exists(entry.Key) Then
    '        parsingName = entry.Key
    '    End If

    '    If Not String.IsNullOrEmpty(parsingName) Then
    '        ' Request BIGGERSIZEOK and SCALEUP for the best quality
    '        Dim bmpShell = GetShellIconBitmap(parsingName, LARGE_ICON_SIZE, SIIGBF.BIGGERSIZEOK Or SIIGBF.SCALEUP)
    '        If bmpShell IsNot Nothing Then
    '            imgDict(entry.Key) = bmpShell
    '            Return
    '        End If
    '    End If

    '    ' Fallback to your custom placeholder
    '    imgDict(entry.Key) = CreatePlaceholder(entry.Name)
    'End Sub

    Private Sub LoadEntryIcon(entry As AppEntry)
        If imgDict.ContainsKey(entry.Key) Then Return

        Dim parsingName As String = ""

        ' 1) Prioritize existing shell: paths (already formatted)
        If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso entry.ShortcutPath.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) Then
            parsingName = entry.ShortcutPath

            ' 2) Check for physical .lnk shortcuts
        ElseIf Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso File.Exists(entry.ShortcutPath) Then
            parsingName = entry.ShortcutPath

            ' 3) Check for direct EXE paths
        ElseIf File.Exists(entry.Key) Then
            parsingName = entry.Key

            ' 4) Final fallback for UWP/Store apps that only provide an AUMID (App User Model ID)
            ' This handles cases where entry.Key is something like "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"
        Else
            If Not String.IsNullOrEmpty(entry.Key) Then
                parsingName = "shell:AppsFolder\" & entry.Key
            End If
        End If

        ' Now try to extract the high-quality icon using the factory
        If Not String.IsNullOrEmpty(parsingName) Then
            ' We use SCALEUP and BIGGERSIZEOK to ensure we get 256px or the closest match
            Dim bmpShell = GetShellIconBitmap(parsingName, LARGE_ICON_SIZE, SIIGBF.BIGGERSIZEOK Or SIIGBF.SCALEUP)

            If bmpShell IsNot Nothing Then
                imgDict(entry.Key) = bmpShell
                Return
            End If
        End If

        ' Fallback to your custom designer-style placeholder if the factory fails
        imgDict(entry.Key) = CreatePlaceholder(entry.Name)
    End Sub

    Private Function ResizeIcon(bmp As Bitmap) As Bitmap
        If bmp.Width = LARGE_ICON_SIZE AndAlso bmp.Height = LARGE_ICON_SIZE Then Return bmp
        Dim result = New Bitmap(LARGE_ICON_SIZE, LARGE_ICON_SIZE, PixelFormat.Format32bppPArgb)
        Using g = Graphics.FromImage(result)
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.SmoothingMode = SmoothingMode.HighQuality
            g.DrawImage(bmp, 0, 0, LARGE_ICON_SIZE, LARGE_ICON_SIZE)
        End Using
        Return result
    End Function

    Private Function CreatePlaceholder(name As String) As Bitmap
        Dim bmp = New Bitmap(LARGE_ICON_SIZE, LARGE_ICON_SIZE, PixelFormat.Format32bppPArgb)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.HighQuality
            Using br = New SolidBrush(clrAccent)
                g.FillEllipse(br, 2, 2, LARGE_ICON_SIZE - 4, LARGE_ICON_SIZE - 4)
            End Using
            If Not String.IsNullOrEmpty(name) Then
                Dim letter = name.Substring(0, 1).ToUpper()
                Using fnt = New Font("Segoe UI", 16, FontStyle.Bold)
                    Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                        g.DrawString(letter, fnt, Brushes.White, New RectangleF(0, 0, LARGE_ICON_SIZE, LARGE_ICON_SIZE), sf)
                    End Using
                End Using
            End If
        End Using
        Return bmp
    End Function

    Private Function HBitmapToBitmapAlpha(hbm As IntPtr) As Bitmap
        ' Image.FromHbitmap() strips alpha — we MUST read raw DIB pixels directly.
        Dim bm As New BITMAPINFO_NATIVE()
        Dim cb = GetBitmapObject(hbm, Marshal.SizeOf(GetType(BITMAPINFO_NATIVE)), bm)

        If cb = 0 OrElse bm.bmBits = IntPtr.Zero OrElse bm.bmBitsPixel <> 32 Then
            ' Not a 32bpp DIB section — fall back (will lose alpha, but at least won't crash)
            Return Bitmap.FromHbitmap(hbm)
        End If

        Dim w = bm.bmWidth
        Dim h = bm.bmHeight

        ' Create a managed bitmap with premultiplied alpha format
        Dim result As New Bitmap(w, h, PixelFormat.Format32bppPArgb)

        ' Lock the managed bitmap so we can copy raw pixels into it
        Dim lockRect As New Rectangle(0, 0, w, h)
        Dim bmpData = result.LockBits(lockRect, ImageLockMode.WriteOnly, PixelFormat.Format32bppPArgb)

        ' The DIB from Shell is bottom-up (first row in memory = bottom of image).
        ' GDI+ Bitmap is top-down (first row in memory = top of image).
        ' We must flip rows during copy.
        Dim srcStride = w * 4
        For row = 0 To h - 1
            Dim srcRow = bm.bmBits + (h - 1 - row) * srcStride   ' bottom-up source
            Dim dstRow = bmpData.Scan0 + row * bmpData.Stride     ' top-down destination
            CopyMemory(dstRow, srcRow, srcStride)
        Next

        result.UnlockBits(bmpData)
        Return result
    End Function

    ' Fast memory copy for pixel data
    <DllImport("kernel32.dll", EntryPoint:="RtlMoveMemory")>
    Private Shared Sub CopyMemory(dest As IntPtr, src As IntPtr, byteCount As Integer)
    End Sub


    Private Function GetShellIconBitmap(parsingName As String, size As Integer, flags As SIIGBF) As Bitmap
        Try
            Dim iid As New Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe") ' IShellItem
            Dim pUnk As IntPtr = IntPtr.Zero
            Dim hr = SHCreateItemFromParsingName(parsingName, IntPtr.Zero, iid, pUnk)
            If hr <> 0 Then Return Nothing

            Dim factory = TryCast(Marshal.GetObjectForIUnknown(pUnk), IShellItemImageFactory)
            Marshal.Release(pUnk)
            If factory Is Nothing Then Return Nothing

            Dim hbm As IntPtr = IntPtr.Zero
            ' Use the SIIGBF.ICONONLY flag if you don't want folder thumbnails
            Dim res = factory.GetImage(New SHSIZE(size, size), flags Or SIIGBF.ICONONLY, hbm)

            If res = 0 AndAlso hbm <> IntPtr.Zero Then
                Try
                    ' FIX: Instead of FromHbitmap, we use this to preserve Alpha
                    'Dim raw = Image.FromHbitmap(hbm)
                    'Dim result = New Bitmap(raw.Width, raw.Height, PixelFormat.Format32bppPArgb)

                    'Using g = Graphics.FromImage(result)
                    '    g.Clear(Color.Transparent)
                    '    g.InterpolationMode = InterpolationMode.HighQualityBicubic
                    '    g.DrawImage(raw, 0, 0, size, size)
                    'End Using

                    'raw.Dispose()
                    'Return result

                    Dim result = HBitmapToBitmapAlpha(hbm)
                    Return result
                Finally
                    DeleteObject(hbm)
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("Icon Error: " & ex.Message)
        End Try
        Return Nothing
    End Function



    Private Function CategorizeApp(entry As AppEntry) As String
        Dim n = entry.Name.ToLowerInvariant()
        Dim k = entry.Key.ToLowerInvariant()
        Dim g = If(entry.Group, "").ToLowerInvariant()

        ' Developer Tools
        If n.Contains("visual studio") OrElse n.Contains("code") OrElse n.Contains("powershell") OrElse
           n.Contains("command prompt") OrElse n.Contains("git ") OrElse n.Contains("docker") OrElse
           n.Contains("terminal") OrElse n.Contains("developer") OrElse n.Contains("debugg") OrElse
           n.Contains("godot") OrElse n.Contains("registry") OrElse g.Contains("developer") Then
            Return "Developer Tools"
        End If
        ' Creative Apps
        If n.Contains("photoshop") OrElse n.Contains("figma") OrElse n.Contains("lightroom") OrElse
           n.Contains("gimp") OrElse n.Contains("blender") OrElse n.Contains("inkscape") OrElse
           n.Contains("premiere") OrElse n.Contains("after effects") OrElse n.Contains("obs") OrElse
           n.Contains("openshot") OrElse n.Contains("paint") OrElse n.Contains("photo") OrElse
           n.Contains("video") OrElse n.Contains("movie") Then
            Return "Creative Apps"
        End If
        ' Quick Access
        If n.Contains("file explorer") OrElse n.Contains("calculator") OrElse n.Contains("notepad") OrElse
           n.Contains("settings") OrElse n.Contains("snipping") OrElse n.Contains("onedrive") OrElse
           n.Contains("control panel") OrElse n.Contains("task manager") Then
            Return "Quick Access"
        End If
        ' Browsers
        If n.Contains("edge") OrElse n.Contains("chrome") OrElse n.Contains("firefox") OrElse
           n.Contains("zen") OrElse n.Contains("brave") OrElse n.Contains("opera") Then
            Return "Browsers"
        End If
        ' Gaming
        If n.Contains("xbox") OrElse n.Contains("steam") OrElse n.Contains("epic") OrElse
           n.Contains("hytale") OrElse n.Contains("minecraft") OrElse n.Contains("game") Then
            Return "Gaming"
        End If

        Return ""
    End Function

    Private Sub BuildPinned()
        pinnedEntries.Clear()
        ' First try taskbar pinned shortcuts
        Try
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim taskband = Path.Combine(appData, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
            If Directory.Exists(taskband) Then
                Dim shortcuts = Directory.GetFiles(taskband, "*.lnk")
                For Each s In shortcuts.Take(8)
                    Try
                        Dim key = ResolveShortcutKey(s)
                        Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                        If entry IsNot Nothing AndAlso Not pinnedEntries.Contains(entry) Then
                            pinnedEntries.Add(entry)
                        End If
                    Catch
                    End Try
                Next
            End If
        Catch
        End Try

        ' Fill up to 12 with popular apps if needed
        If pinnedEntries.Count < 12 Then
            Dim popularNames = {"Microsoft Edge", "File Explorer", "Settings", "Visual Studio",
                                "Visual Studio Code", "Notepad", "Calculator", "Microsoft Store",
                                "Spotify", "Discord", "Steam", "Task Manager"}
            For Each pn In popularNames
                If pinnedEntries.Count >= 12 Then Exit For
                Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Name, pn, StringComparison.OrdinalIgnoreCase))
                If entry IsNot Nothing AndAlso Not pinnedEntries.Contains(entry) Then
                    pinnedEntries.Add(entry)
                End If
            Next
        End If
    End Sub

    Private Sub LoadSavedPins()
        ' If user has saved pinned apps, replace the defaults with their saved list
        Try
            Dim saved = My.Settings.PinnedAppKeys
            If Not String.IsNullOrEmpty(saved) Then
                Dim keys = saved.Split("|"c)
                Dim restored As New List(Of AppEntry)()
                For Each key In keys
                    If String.IsNullOrEmpty(key) Then Continue For
                    Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                    If entry IsNot Nothing Then restored.Add(entry)
                Next
                If restored.Count > 0 Then
                    pinnedEntries.Clear()
                    pinnedEntries.AddRange(restored)
                End If
            End If
        Catch
            ' First run or settings unavailable — keep defaults from BuildPinned
        End Try
    End Sub

    Private Sub SavePins()
        Try
            Dim keys = String.Join("|", pinnedEntries.Select(Function(e) e.Key))
            My.Settings.PinnedAppKeys = keys
            My.Settings.Save()
        Catch
        End Try
    End Sub

    Private Sub BuildCategories()
        categories.Clear()

        ' Load custom categories first (they appear at top)
        LoadCustomCategories()

        ' Then add auto-categories
        Dim catGroups = allEntries.Where(Function(a) Not String.IsNullOrEmpty(a.Category)).
            GroupBy(Function(a) a.Category).
            OrderBy(Function(g) g.Key).ToList()

        For Each cg In catGroups
            ' Don't duplicate if a custom category has the same name
            If Not customCategoryNames.Contains(cg.Key) Then
                categories.Add((cg.Key, cg.ToList()))
            End If
        Next
    End Sub

    Private Sub LoadCustomCategories()
        ' Format: "CatName:key1,key2,key3|CatName2:key4,key5"
        Try
            Dim saved = My.Settings.CustomCategories
            If String.IsNullOrEmpty(saved) Then Return
            For Each catBlock In saved.Split("|"c)
                If String.IsNullOrEmpty(catBlock) Then Continue For
                Dim colonIdx = catBlock.IndexOf(":"c)
                If colonIdx < 0 Then Continue For
                Dim catName = catBlock.Substring(0, colonIdx)
                Dim keysStr = catBlock.Substring(colonIdx + 1)
                customCategoryNames.Add(catName)

                Dim entries As New List(Of AppEntry)()
                If Not String.IsNullOrEmpty(keysStr) Then
                    For Each key In keysStr.Split(","c)
                        If String.IsNullOrEmpty(key) Then Continue For
                        Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                        If entry IsNot Nothing Then entries.Add(entry)
                    Next
                End If
                categories.Add((catName, entries))
            Next
        Catch
        End Try
    End Sub

    Private Sub SaveCustomCategories()
        Try
            Dim parts As New List(Of String)()
            For Each cat In categories
                If Not customCategoryNames.Contains(cat.Name) Then Continue For
                Dim keys = String.Join(",", cat.Entries.Select(Function(e) e.Key))
                parts.Add(cat.Name & ":" & keys)
            Next
            My.Settings.CustomCategories = String.Join("|", parts)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    Private Sub AddAppToCustomCategory(catName As String, entry As AppEntry)
        For i = 0 To categories.Count - 1
            If String.Equals(categories(i).Name, catName, StringComparison.OrdinalIgnoreCase) Then
                If Not categories(i).Entries.Any(Function(e) e.Key = entry.Key) Then
                    categories(i).Entries.Add(entry)
                    SaveCustomCategories()
                    pnlCategories.Invalidate()
                End If
                Return
            End If
        Next
    End Sub

    Private Sub CreateCustomCategory(name As String)
        If String.IsNullOrEmpty(name) Then Return
        If customCategoryNames.Contains(name) Then Return
        customCategoryNames.Add(name)
        ' Insert at position 0 so custom categories stay at top
        categories.Insert(0, (name, New List(Of AppEntry)()))
        SaveCustomCategories()
        pnlCategories.Invalidate()
    End Sub

    Private Sub RemoveCustomCategory(catName As String)
        If Not customCategoryNames.Contains(catName) Then Return
        customCategoryNames.Remove(catName)
        categories.RemoveAll(Function(c) String.Equals(c.Name, catName, StringComparison.OrdinalIgnoreCase))
        SaveCustomCategories()
        pnlCategories.Invalidate()
    End Sub

    Private Sub RemoveAppFromCustomCategory(catName As String, entryKey As String)
        For i = 0 To categories.Count - 1
            If String.Equals(categories(i).Name, catName, StringComparison.OrdinalIgnoreCase) Then
                categories(i).Entries.RemoveAll(Function(e) e.Key = entryKey)
                SaveCustomCategories()
                pnlCategories.Invalidate()
                Return
            End If
        Next
    End Sub

    Private Function ResolveShortcutKey(shortcutPath As String) As String
        Try
            Dim shell = CreateObject("WScript.Shell")
            Dim lnk = shell.CreateShortcut(shortcutPath)
            Dim target = If(lnk.TargetPath, String.Empty)
            Dim args = If(lnk.Arguments, String.Empty)
            If Not String.IsNullOrEmpty(target) Then
                Dim tname = Path.GetFileName(target).ToLowerInvariant()
                If tname = "explorer.exe" AndAlso Not String.IsNullOrEmpty(args) AndAlso args.IndexOf("appsfolder", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return args.Trim()
                Else
                    Return Path.GetFullPath(target)
                End If
            ElseIf Not String.IsNullOrEmpty(args) AndAlso args.IndexOf("appsfolder", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return args.Trim()
            End If
        Catch
        End Try
        Return shortcutPath
    End Function
#End Region

#Region "Painting"
    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        If micaApplied Then
            ' Paint semi-transparent so Mica shows through but text has a surface for ClearType
            Using br = New SolidBrush(Color.FromArgb(180, 28, 28, 28))
                e.Graphics.FillRectangle(br, Me.ClientRectangle)
            End Using
            Return
        End If
        ' Mica not available — paint solid dark background
        Using br = New SolidBrush(Color.FromArgb(255, 30, 30, 30))
            e.Graphics.FillRectangle(br, Me.ClientRectangle)
        End Using
    End Sub

    Private Sub PaintUserBar(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.HighQuality
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

        ' User name on left
        Dim userName = Environment.UserName
        Using fnt = New Font("Segoe UI", 12, FontStyle.Regular)
            Using br = New SolidBrush(clrTextPrimary)
                ' Circle avatar
                Dim avatarRect = New Rectangle(SECTION_PAD, 12, 32, 32)
                Using avBr = New SolidBrush(clrAccent)
                    g.FillEllipse(avBr, avatarRect)
                End Using
                Dim initial = userName.Substring(0, 1).ToUpper()
                Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(initial, fnt, Brushes.White, New RectangleF(avatarRect.X, avatarRect.Y, avatarRect.Width, avatarRect.Height), sf)
                End Using
                g.DrawString(userName, fnt, br, SECTION_PAD + 40, 18)
            End Using
        End Using

        ' Mica toggle switch
        Dim toggleX = pnlUserBar.Width - 160
        Dim toggleY = 20
        Dim trackW = 40
        Dim trackH = 20
        Dim knobSize = 16
        Dim knobPad = 2

        ' Track (pill shape)
        Dim trackRect = New Rectangle(toggleX, toggleY, trackW, trackH)
        Dim trackColor = If(micaApplied, clrAccent, Color.FromArgb(120, 80, 80, 80))
        Using trackPath = New GraphicsPath()
            Dim r = trackH \ 2
            trackPath.AddArc(trackRect.X, trackRect.Y, trackH, trackH, 90, 180)
            trackPath.AddArc(trackRect.Right - trackH, trackRect.Y, trackH, trackH, 270, 180)
            trackPath.CloseFigure()
            Using br = New SolidBrush(trackColor)
                g.FillPath(br, trackPath)
            End Using
        End Using

        ' Knob (circle)
        Dim knobX = If(micaApplied, toggleX + trackW - knobSize - knobPad, toggleX + knobPad)
        Dim knobY = toggleY + knobPad
        Using br = New SolidBrush(Color.White)
            g.FillEllipse(br, knobX, knobY, knobSize, knobSize)
        End Using

        ' Label
        Dim lblText = If(micaApplied, "Mica", "Mica")
        Using fnt = New Font("Segoe UI", 7.5F, FontStyle.Regular)
            Using br = New SolidBrush(clrTextSecondary)
                g.DrawString(lblText, fnt, br, toggleX + trackW + 4, toggleY + 2)
            End Using
        End Using

        ' Power button on far right
        Using fnt = New Font("Segoe UI Symbol", 14, FontStyle.Regular)
            Using br = New SolidBrush(clrTextSecondary)
                g.DrawString(ChrW(&H23FB), fnt, br, pnlUserBar.Width - 50, 14) ' power symbol ⏻
            End Using
        End Using
    End Sub

    Private Sub PaintPinned(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.HighQuality
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

        ' Section header
        Dim headerY = 4
        Using fnt = New Font("Segoe UI", 13, FontStyle.Bold)
            Using br = New SolidBrush(clrTextPrimary)
                g.DrawString("Pinned", fnt, br, 4, headerY)
            End Using
        End Using

        ' Render pinned cards in a grid
        Dim startY = headerY + 32
        Dim x = 4
        Dim y = startY
        Dim query = txtSearch.Text.Trim()
        Dim isSearching = Not String.IsNullOrEmpty(query)
        Dim entriesToShow = If(isSearching,
            allEntries.Where(Function(a) a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).Take(12).ToList(),
            pinnedEntries)

        If isSearching Then
            Using fnt = New Font("Segoe UI", 13, FontStyle.Bold)
                Using br = New SolidBrush(clrTextPrimary)
                    g.DrawString("Search Results", fnt, br, 4, headerY)
                End Using
            End Using
        End If

        For i = 0 To entriesToShow.Count - 1
            Dim entry = entriesToShow(i)
            Dim cardRect = New Rectangle(x, y, CARD_W, CARD_H)
            Dim isHovered = (pinnedHoveredIdx = i)

            ' Card background with rounded corners
            DrawRoundedCard(g, cardRect, If(isHovered, clrCardHover, clrCardBg), 8)

            ' Icon
            Dim icon As Bitmap = Nothing
            If imgDict.ContainsKey(entry.Key) Then icon = imgDict(entry.Key)
            If icon IsNot Nothing Then
                Dim iconX = cardRect.X + (CARD_W - ICON_SIZE) \ 2
                Dim iconY = cardRect.Y + 12
                g.DrawImage(icon, iconX, iconY, ICON_SIZE, ICON_SIZE)
            End If

            ' Name
            Using fnt = New Font("Segoe UI", 8, FontStyle.Regular)
                Using br = New SolidBrush(clrTextPrimary)
                    Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Near, .Trimming = StringTrimming.EllipsisCharacter}
                        Dim nameRect = New RectangleF(cardRect.X + 2, cardRect.Y + ICON_SIZE + 16, CARD_W - 4, CARD_H - ICON_SIZE - 18)
                        g.DrawString(entry.Name, fnt, br, nameRect, sf)
                    End Using
                End Using
            End Using

            x += CARD_W + CARD_GAP
            If x + CARD_W > LEFT_CONTENT_W Then
                x = 4
                y += CARD_H + CARD_GAP
            End If
        Next

        ' Auto-size pinned panel height based on content
        Dim neededH = y + CARD_H + CARD_GAP + 10
        If neededH <> pnlPinned.Height AndAlso neededH > 80 Then
            pnlPinned.Height = neededH
        End If
    End Sub

    Private Sub PaintCategories(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.HighQuality
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

        ' Section header
        Using fnt = New Font("Segoe UI", 13, FontStyle.Bold)
            Using br = New SolidBrush(clrTextPrimary)
                g.DrawString("Categories", fnt, br, 4, 4)
            End Using
        End Using

        ' "New category" drop target (visible during drag)
        If isDragging OrElse dragDropNewTarget Then
            Dim newZoneRect = New Rectangle(LEFT_CONTENT_W - 80, 2, 76, 26)
            Dim zoneColor = If(dragDropNewTarget, clrAccent, Color.FromArgb(80, 96, 165, 250))
            DrawRoundedCard(g, newZoneRect, zoneColor, 6)
            Using fnt = New Font("Segoe UI", 9, FontStyle.Bold)
                Using br = New SolidBrush(If(dragDropNewTarget, Color.White, clrTextSecondary))
                    Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                        g.DrawString("+ New", fnt, br, New RectangleF(newZoneRect.X, newZoneRect.Y, newZoneRect.Width, newZoneRect.Height), sf)
                    End Using
                End Using
            End Using
        End If

        If categories.Count = 0 Then Return

        Dim panelW = LEFT_CONTENT_W
        Dim cols = 3
        Dim cardSpacing = 6
        Dim groupPad = 10  ' padding inside the group border
        Dim headerH = 28   ' height for category header inside group
        Dim collapsedBlockW = groupPad * 2 + cols * (CAT_CARD_W + cardSpacing) - cardSpacing
        Dim curX = 4
        Dim curY = 34
        Dim catIdx = 0
        Dim maxBottom = curY
        Dim rowMaxBottom = curY  ' tallest group in current row

        For Each cat In categories
            Dim hasChevron = (cat.Entries.Count > 3)
            Dim isExpanded = hasChevron AndAlso expandedCategories.Contains(cat.Name)

            ' Determine which entries are visible
            Dim entriesToShow As List(Of AppEntry)
            If Not hasChevron Then
                ' 3 or fewer: always show all, no expand/collapse
                entriesToShow = cat.Entries
            ElseIf isExpanded Then
                ' Expanded: show all but cap visible rows for scrolling
                Dim scrollOff = 0
                If catScrollOffsets.ContainsKey(cat.Name) Then scrollOff = catScrollOffsets(cat.Name)
                Dim maxVisible = CAT_MAX_EXPANDED_ROWS * cols
                entriesToShow = cat.Entries.Skip(scrollOff * cols).Take(maxVisible).ToList()
            Else
                ' Collapsed: show first row only
                entriesToShow = cat.Entries.Take(cols).ToList()
            End If

            Dim entryRows = CInt(Math.Ceiling(entriesToShow.Count / CDbl(cols)))
            Dim scrollBarH = 0
            If isExpanded AndAlso cat.Entries.Count > CAT_MAX_EXPANDED_ROWS * cols Then
                scrollBarH = 14  ' space for scroll indicator
            End If
            Dim groupH = groupPad + headerH + entryRows * (CAT_CARD_H + cardSpacing) - cardSpacing + groupPad + scrollBarH
            Dim groupW = collapsedBlockW

            ' Wrap to next row if this group won't fit
            If curX > 4 AndAlso curX + groupW > panelW Then
                curX = 4
                curY = rowMaxBottom + 10
                rowMaxBottom = curY
            End If

            ' Draw group border container (accent border for custom categories, glow during drag target)
            Dim isCustom = customCategoryNames.Contains(cat.Name)
            Dim isDragTarget = (dragDropTargetCat IsNot Nothing AndAlso String.Equals(dragDropTargetCat, cat.Name, StringComparison.OrdinalIgnoreCase))
            Dim groupRect = New Rectangle(curX, curY, groupW, groupH)
            Dim borderClr As Color
            Dim bgClr As Color
            If isDragTarget Then
                borderClr = clrAccent
                bgClr = Color.FromArgb(60, 96, 165, 250)
            ElseIf isCustom Then
                borderClr = clrCustomCatBorder
                bgClr = clrGroupBg
            Else
                borderClr = clrGroupBorder
                bgClr = clrGroupBg
            End If
            DrawRoundedBorder(g, groupRect, bgClr, borderClr, 10)

            ' Category header
            Dim isHeaderHov = (hoveredCatHeader = cat.Name)
            Using fnt = New Font("Segoe UI", 8.5F, FontStyle.Bold)
                Dim headerColor = If(isHeaderHov AndAlso hasChevron, clrTextPrimary, clrTextSecondary)
                Using br = New SolidBrush(headerColor)
                    g.DrawString(cat.Name.ToUpper(), fnt, br, curX + groupPad, curY + groupPad)
                End Using
                ' Chevron indicator — only for categories with 4+ entries
                If hasChevron Then
                    Dim chevron = If(isExpanded, ChrW(&H25B2), ChrW(&H25BC)) ' up/down triangle
                    Using br = New SolidBrush(clrTextSecondary)
                        Using chevFnt = New Font("Segoe UI", 7)
                            Dim chevX = curX + groupW - groupPad - 14
                            g.DrawString(chevron, chevFnt, br, chevX, curY + groupPad + 2)
                        End Using
                    End Using
                End If
            End Using

            ' Draw cards inside the group
            Dim itemX = curX + groupPad
            Dim itemY = curY + groupPad + headerH
            Dim col = 0
            For Each entry In entriesToShow
                Dim cardRect = New Rectangle(itemX, itemY, CAT_CARD_W, CAT_CARD_H)
                Dim isHov = (catHoveredIdx = catIdx)

                DrawRoundedCard(g, cardRect, If(isHov, clrCardHover, clrCardBg), 6)

                Dim icon As Bitmap = Nothing
                If imgDict.ContainsKey(entry.Key) Then icon = imgDict(entry.Key)
                If icon IsNot Nothing Then
                    Dim ix = cardRect.X + (CAT_CARD_W - 24) \ 2
                    g.DrawImage(icon, ix, cardRect.Y + 8, 24, 24)
                End If

                Using fnt = New Font("Segoe UI", 7, FontStyle.Regular)
                    Using br = New SolidBrush(clrTextPrimary)
                        Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
                            Dim nr = New RectangleF(cardRect.X, cardRect.Y + 36, CAT_CARD_W, CAT_CARD_H - 38)
                            g.DrawString(entry.Name, fnt, br, nr, sf)
                        End Using
                    End Using
                End Using

                col += 1
                catIdx += 1
                If col >= cols Then
                    col = 0
                    itemX = curX + groupPad
                    itemY += CAT_CARD_H + cardSpacing
                Else
                    itemX += CAT_CARD_W + cardSpacing
                End If
            Next

            ' Draw scroll arrows for expanded categories with overflow
            If isExpanded AndAlso cat.Entries.Count > CAT_MAX_EXPANDED_ROWS * cols Then
                Dim scrollOff = 0
                If catScrollOffsets.ContainsKey(cat.Name) Then scrollOff = catScrollOffsets(cat.Name)
                Dim totalRows = CInt(Math.Ceiling(cat.Entries.Count / CDbl(cols)))
                Dim maxScroll = Math.Max(0, totalRows - CAT_MAX_EXPANDED_ROWS)
                Dim barY = curY + groupH - groupPad - 8

                ' Left arrow (scroll up) and right arrow (scroll down)
                Using arrowFnt = New Font("Segoe UI", 8, FontStyle.Bold)
                    ' Left arrow
                    Dim leftColor = If(scrollOff > 0, clrTextPrimary, Color.FromArgb(60, 255, 255, 255))
                    Using br = New SolidBrush(leftColor)
                        g.DrawString(ChrW(&H25C0), arrowFnt, br, curX + groupPad, barY) ' ◀
                    End Using

                    ' Page indicator: "2 / 5"
                    Dim pageText = $"{scrollOff + 1} / {maxScroll + 1}"
                    Using br = New SolidBrush(clrTextSecondary)
                        Using sf = New StringFormat() With {.Alignment = StringAlignment.Center}
                            g.DrawString(pageText, arrowFnt, br, New RectangleF(curX, barY, groupW, 14), sf)
                        End Using
                    End Using

                    ' Right arrow
                    Dim rightColor = If(scrollOff < maxScroll, clrTextPrimary, Color.FromArgb(60, 255, 255, 255))
                    Using br = New SolidBrush(rightColor)
                        Dim rightX = curX + groupW - groupPad - 12
                        g.DrawString(ChrW(&H25B6), arrowFnt, br, rightX, barY) ' ▶
                    End Using
                End Using
            End If

            ' Track the bottom of this group for panel sizing
            Dim thisBottom = curY + groupH
            If thisBottom > rowMaxBottom Then rowMaxBottom = thisBottom
            If thisBottom > maxBottom Then maxBottom = thisBottom

            curX += groupW + 10
        Next

        ' Auto-size the panel height
        Dim neededH = maxBottom + 10
        If pnlCategories.Height <> neededH Then
            pnlCategories.Height = neededH
        End If
    End Sub

    Private Sub PaintAllApps(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.HighQuality
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

        ' Header
        Using fnt = New Font("Segoe UI", 13, FontStyle.Bold)
            Using br = New SolidBrush(clrTextPrimary)
                g.DrawString("All Apps", fnt, br, 12, 4)
            End Using
        End Using

        ' Build sorted list grouped by first letter
        Dim query = txtSearch.Text.Trim()
        Dim filtered = allEntries.Where(Function(a) String.IsNullOrEmpty(query) OrElse a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(Function(a) a.Name).ToList()

        Dim y = 34 - allAppsScrollOffset
        Dim visibleIdx = 0
        Dim currentLetter As String = ""
        Dim clipRect = New Rectangle(0, 34, pnlAllApps.ClientSize.Width, pnlAllApps.ClientSize.Height - 34)
        g.SetClip(clipRect)

        For i = 0 To filtered.Count - 1
            Dim entry = filtered(i)
            Dim firstLetter = entry.Name.Substring(0, 1).ToUpper()
            If Not Char.IsLetter(firstLetter(0)) Then firstLetter = "#"

            ' Draw letter header
            If firstLetter <> currentLetter Then
                currentLetter = firstLetter
                If y + ALLAPPS_LETTER_H > 34 AndAlso y < pnlAllApps.ClientSize.Height Then
                    Using fnt = New Font("Segoe UI", 10, FontStyle.Bold)
                        Using br = New SolidBrush(clrAccent)
                            g.DrawString(currentLetter, fnt, br, 12, y + 2)
                        End Using
                    End Using
                End If
                y += ALLAPPS_LETTER_H
            End If

            ' Draw app row
            If y + ALLAPPS_ROW_H > 34 AndAlso y < pnlAllApps.ClientSize.Height Then
                Dim rowRect = New Rectangle(4, y, pnlAllApps.ClientSize.Width - 8, ALLAPPS_ROW_H)
                Dim isHov = (allAppsHoveredIdx = i)
                If isHov Then
                    DrawRoundedCard(g, rowRect, clrCardHover, 4)
                End If

                ' Icon
                Dim icon As Bitmap = Nothing
                If imgDict.ContainsKey(entry.Key) Then icon = imgDict(entry.Key)
                If icon IsNot Nothing Then
                    g.DrawImage(icon, 14, y + (ALLAPPS_ROW_H - 20) \ 2, 20, 20)
                End If

                ' Name
                Using fnt = New Font("Segoe UI", 9, FontStyle.Regular)
                    Using br = New SolidBrush(clrTextPrimary)
                        Using sf = New StringFormat() With {.LineAlignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter, .FormatFlags = StringFormatFlags.NoWrap}
                            g.DrawString(entry.Name, fnt, br, New RectangleF(40, y, pnlAllApps.ClientSize.Width - 50, ALLAPPS_ROW_H), sf)
                        End Using
                    End Using
                End Using
            End If

            y += ALLAPPS_ROW_H
        Next

        g.ResetClip()

        ' Draw letter grid overlay when active
        If showLetterGrid Then
            PaintLetterGrid(g)
        End If
    End Sub

    Private Sub PaintLetterGrid(g As Graphics)
        Dim panelW = pnlAllApps.ClientSize.Width
        Dim panelH = pnlAllApps.ClientSize.Height

        ' Dim background
        Using br = New SolidBrush(Color.FromArgb(220, 25, 25, 25))
            g.FillRectangle(br, 0, 0, panelW, panelH)
        End Using

        Dim allLetters = {"#", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                          "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        Dim available = GetAvailableLetters()

        Dim cols = 7
        Dim cellSize = 30
        Dim cellGap = 4
        Dim gridW = cols * (cellSize + cellGap) - cellGap
        Dim gridX = (panelW - gridW) \ 2
        Dim gridY = 40

        ' Title
        Using fnt = New Font("Segoe UI", 11, FontStyle.Bold)
            Using br = New SolidBrush(clrTextPrimary)
                Using sf = New StringFormat() With {.Alignment = StringAlignment.Center}
                    g.DrawString("Jump to", fnt, br, New RectangleF(0, 10, panelW, 24), sf)
                End Using
            End Using
        End Using

        Using fnt = New Font("Segoe UI", 11, FontStyle.Bold)
            Using sf = New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                For i = 0 To allLetters.Length - 1
                    Dim col = i Mod cols
                    Dim row = i \ cols
                    Dim cx = gridX + col * (cellSize + cellGap)
                    Dim cy = gridY + row * (cellSize + cellGap)
                    Dim cellRect = New Rectangle(cx, cy, cellSize, cellSize)

                    Dim hasEntries = available.Contains(allLetters(i))
                    Dim isHovered = (letterGridHoveredIdx = i)

                    ' Cell background
                    Dim cellColor As Color
                    If isHovered AndAlso hasEntries Then
                        cellColor = clrAccent
                    ElseIf hasEntries Then
                        cellColor = clrCardBg
                    Else
                        cellColor = Color.FromArgb(40, 40, 40, 40)
                    End If
                    DrawRoundedCard(g, cellRect, cellColor, 6)

                    ' Letter text
                    Dim textColor = If(hasEntries, clrTextPrimary, Color.FromArgb(60, 200, 200, 200))
                    If isHovered AndAlso hasEntries Then textColor = Color.White
                    Using br = New SolidBrush(textColor)
                        g.DrawString(allLetters(i), fnt, br, New RectangleF(cx, cy, cellSize, cellSize), sf)
                    End Using
                Next
            End Using
        End Using
    End Sub

    Private Sub DrawRoundedCard(g As Graphics, rect As Rectangle, fillColor As Color, radius As Integer)
        Using path = MakeRoundedPath(rect, radius)
            Using br = New SolidBrush(fillColor)
                g.FillPath(br, path)
            End Using
        End Using
    End Sub

    Private Function MakeRoundedPath(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path = New GraphicsPath()
        Dim d = radius * 2
        path.AddArc(rect.X, rect.Y, d, d, 180, 90)
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    Private Sub DrawRoundedBorder(g As Graphics, rect As Rectangle, fillColor As Color, borderColor As Color, radius As Integer)
        Using path = MakeRoundedPath(rect, radius)
            Using br = New SolidBrush(fillColor)
                g.FillPath(br, path)
            End Using
            Using pn = New Pen(borderColor, 1)
                g.DrawPath(pn, path)
            End Using
        End Using
    End Sub
#End Region

#Region "Interaction"
    Private Sub OnSearchTextChanged(sender As Object, e As EventArgs)
        pnlPinned.Invalidate()
        pnlAllApps.Invalidate()
    End Sub

    Private Sub LaunchEntry(entry As AppEntry)
        If entry Is Nothing OrElse String.IsNullOrEmpty(entry.ShortcutPath) Then Return
        Try
            Process.Start(New ProcessStartInfo(entry.ShortcutPath) With {.UseShellExecute = True})
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Failed to launch: " & ex.Message)
        End Try
    End Sub

    ' --- Pinned section mouse handling ---
    Private Function GetPinnedEntryAt(pt As Point) As (Index As Integer, Entry As AppEntry)
        Dim query = txtSearch.Text.Trim()
        Dim isSearching = Not String.IsNullOrEmpty(query)
        Dim entriesToShow = If(isSearching,
            allEntries.Where(Function(a) a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).Take(12).ToList(),
            DirectCast(pinnedEntries, IList(Of AppEntry)))

        Dim startY = 36
        Dim x = 4
        Dim y = startY
        For i = 0 To entriesToShow.Count - 1
            Dim cardRect = New Rectangle(x, y, CARD_W, CARD_H)
            If cardRect.Contains(pt) Then Return (i, entriesToShow(i))
            x += CARD_W + CARD_GAP
            If x + CARD_W > LEFT_CONTENT_W Then
                x = 4
                y += CARD_H + CARD_GAP
            End If
        Next
        Return (-1, Nothing)
    End Function

    Private Sub PinnedClick(sender As Object, e As MouseEventArgs)
        Dim hit = GetPinnedEntryAt(e.Location)
        If hit.Entry IsNot Nothing Then LaunchEntry(hit.Entry)
    End Sub

    Private Sub PinnedMouseMove(sender As Object, e As MouseEventArgs)
        Dim hit = GetPinnedEntryAt(e.Location)
        If hit.Index <> pinnedHoveredIdx Then
            pinnedHoveredIdx = hit.Index
            pnlPinned.Invalidate()
        End If
        pnlPinned.Cursor = If(hit.Index >= 0, Cursors.Hand, Cursors.Default)
    End Sub

    Private Sub PinnedMouseLeave(sender As Object, e As EventArgs)
        If pinnedHoveredIdx >= 0 Then
            pinnedHoveredIdx = -1
            pnlPinned.Invalidate()
        End If
    End Sub

    ' --- Categories mouse handling ---
    Private Const CAT_COLS As Integer = 3
    Private Const CAT_GROUP_PAD As Integer = 10
    Private Const CAT_HEADER_H As Integer = 28

    Private Function GetCatGroupLayout() As List(Of (Name As String, Rect As Rectangle, HeaderRect As Rectangle, Entries As List(Of AppEntry)))
        Dim result As New List(Of (Name As String, Rect As Rectangle, HeaderRect As Rectangle, Entries As List(Of AppEntry)))()
        Dim panelW = LEFT_CONTENT_W
        Dim cardSpacing = 6
        Dim groupW = CAT_GROUP_PAD * 2 + CAT_COLS * (CAT_CARD_W + cardSpacing) - cardSpacing
        Dim curX = 4
        Dim curY = 34
        Dim rowMaxBottom = curY

        For Each cat In categories
            Dim hasChevron = (cat.Entries.Count > 3)
            Dim isExpanded = hasChevron AndAlso expandedCategories.Contains(cat.Name)

            Dim entriesToShow As List(Of AppEntry)
            If Not hasChevron Then
                entriesToShow = cat.Entries
            ElseIf isExpanded Then
                Dim scrollOff = 0
                If catScrollOffsets.ContainsKey(cat.Name) Then scrollOff = catScrollOffsets(cat.Name)
                Dim maxVisible = CAT_MAX_EXPANDED_ROWS * CAT_COLS
                entriesToShow = cat.Entries.Skip(scrollOff * CAT_COLS).Take(maxVisible).ToList()
            Else
                entriesToShow = cat.Entries.Take(CAT_COLS).ToList()
            End If

            Dim entryRows = CInt(Math.Ceiling(entriesToShow.Count / CDbl(CAT_COLS)))
            Dim scrollBarH = 0
            If isExpanded AndAlso cat.Entries.Count > CAT_MAX_EXPANDED_ROWS * CAT_COLS Then
                scrollBarH = 14
            End If
            Dim groupH = CAT_GROUP_PAD + CAT_HEADER_H + entryRows * (CAT_CARD_H + cardSpacing) - cardSpacing + CAT_GROUP_PAD + scrollBarH

            If curX > 4 AndAlso curX + groupW > panelW Then
                curX = 4
                curY = rowMaxBottom + 10
                rowMaxBottom = curY
            End If

            Dim groupRect = New Rectangle(curX, curY, groupW, groupH)
            Dim headerRect = New Rectangle(curX, curY, groupW, CAT_GROUP_PAD + CAT_HEADER_H)
            result.Add((cat.Name, groupRect, headerRect, entriesToShow))

            If curY + groupH > rowMaxBottom Then rowMaxBottom = curY + groupH
            curX += groupW + 10
        Next
        Return result
    End Function

    Private Function GetCatEntryAt(pt As Point) As (Index As Integer, Entry As AppEntry, HeaderName As String)
        Dim layout = GetCatGroupLayout()
        Dim catIdx = 0
        Dim cardSpacing = 6

        For Each grp In layout
            ' Check if clicking on the header area
            If grp.HeaderRect.Contains(pt) Then
                Return (-1, Nothing, grp.Name)
            End If

            ' Check individual cards
            Dim itemX = grp.Rect.X + CAT_GROUP_PAD
            Dim itemY = grp.Rect.Y + CAT_GROUP_PAD + CAT_HEADER_H
            Dim col = 0
            For Each entry In grp.Entries
                Dim cardRect = New Rectangle(itemX, itemY, CAT_CARD_W, CAT_CARD_H)
                If cardRect.Contains(pt) Then Return (catIdx, entry, Nothing)
                col += 1
                catIdx += 1
                If col >= CAT_COLS Then
                    col = 0
                    itemX = grp.Rect.X + CAT_GROUP_PAD
                    itemY += CAT_CARD_H + cardSpacing
                Else
                    itemX += CAT_CARD_W + cardSpacing
                End If
            Next
        Next
        Return (-1, Nothing, Nothing)
    End Function

    Private Sub CategoriesClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left Then Return
        ' Check if clicking the scrollbar area of an expanded category
        Dim layout = GetCatGroupLayout()
        For Each grp In layout
            If Not grp.Rect.Contains(e.Location) Then Continue For
            Dim cat = categories.FirstOrDefault(Function(c) c.Name = grp.Name)
            If cat.Entries IsNot Nothing AndAlso cat.Entries.Count > CAT_MAX_EXPANDED_ROWS * CAT_COLS AndAlso expandedCategories.Contains(grp.Name) Then
                ' Check if click is in the scrollbar area (bottom 14px of group)
                Dim scrollBarY = grp.Rect.Bottom - CAT_GROUP_PAD - 4
                If e.Y >= scrollBarY AndAlso e.Y <= grp.Rect.Bottom Then
                    Dim scrollOff = 0
                    If catScrollOffsets.ContainsKey(grp.Name) Then scrollOff = catScrollOffsets(grp.Name)
                    Dim totalRows = CInt(Math.Ceiling(cat.Entries.Count / CDbl(CAT_COLS)))
                    Dim maxScroll = Math.Max(0, totalRows - CAT_MAX_EXPANDED_ROWS)

                    ' Click left half = scroll up, right half = scroll down
                    Dim midX = grp.Rect.X + grp.Rect.Width \ 2
                    If e.X < midX Then
                        scrollOff = Math.Max(0, scrollOff - 1)
                    Else
                        scrollOff = Math.Min(maxScroll, scrollOff + 1)
                    End If
                    catScrollOffsets(grp.Name) = scrollOff
                    pnlCategories.Invalidate()
                    Return
                End If
            End If
        Next

        Dim hit = GetCatEntryAt(e.Location)
        ' Clicking a header toggles expand/collapse — only for categories with 4+ entries
        If hit.HeaderName IsNot Nothing Then
            Dim cat = categories.FirstOrDefault(Function(c) c.Name = hit.HeaderName)
            If cat.Entries IsNot Nothing AndAlso cat.Entries.Count > 3 Then
                If expandedCategories.Contains(hit.HeaderName) Then
                    expandedCategories.Remove(hit.HeaderName)
                    catScrollOffsets.Remove(hit.HeaderName)
                Else
                    expandedCategories.Add(hit.HeaderName)
                End If
                pnlCategories.Invalidate()
            End If
            Return
        End If
        If hit.Entry IsNot Nothing Then LaunchEntry(hit.Entry)
    End Sub

    Private Sub CategoriesMouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            catDragStartPoint = e.Location
        End If
    End Sub

    Private Sub CategoriesMouseMove(sender As Object, e As MouseEventArgs)
        ' Start drag from a custom category card
        If e.Button = MouseButtons.Left AndAlso catDragStartPoint <> Point.Empty AndAlso Not isDragging Then
            Dim dx = Math.Abs(e.X - catDragStartPoint.X)
            Dim dy = Math.Abs(e.Y - catDragStartPoint.Y)
            If dx > SystemInformation.DragSize.Width OrElse dy > SystemInformation.DragSize.Height Then
                Dim hit = GetCatEntryAt(catDragStartPoint)
                If hit.Entry IsNot Nothing Then
                    ' Find which category this entry belongs to
                    Dim layout = GetCatGroupLayout()
                    Dim catIdx = 0
                    For Each grp In layout
                        For Each ent In grp.Entries
                            If catIdx = hit.Index AndAlso customCategoryNames.Contains(grp.Name) Then
                                isDragging = True
                                dragSourceCat = grp.Name
                                pnlCategories.Invalidate()
                                pnlCategories.DoDragDrop(hit.Entry, DragDropEffects.Move Or DragDropEffects.Copy)
                                isDragging = False
                                dragSourceCat = Nothing
                                catDragStartPoint = Point.Empty
                                pnlCategories.Invalidate()
                                Return
                            End If
                            catIdx += 1
                        Next
                    Next
                End If
                catDragStartPoint = Point.Empty
            End If
        End If

        Dim hitMove = GetCatEntryAt(e.Location)
        Dim needRepaint = False

        If hitMove.Index <> catHoveredIdx Then
            catHoveredIdx = hitMove.Index
            needRepaint = True
        End If

        Dim newHeaderHov = hitMove.HeaderName
        If newHeaderHov <> hoveredCatHeader Then
            hoveredCatHeader = newHeaderHov
            needRepaint = True
        End If

        If needRepaint Then pnlCategories.Invalidate()
        Dim showHand = hitMove.Index >= 0
        If hitMove.HeaderName IsNot Nothing Then
            Dim cat = categories.FirstOrDefault(Function(c) c.Name = hitMove.HeaderName)
            If cat.Entries IsNot Nothing AndAlso cat.Entries.Count > 3 Then showHand = True
        End If
        pnlCategories.Cursor = If(showHand, Cursors.Hand, Cursors.Default)
    End Sub

    Private Sub CategoriesMouseLeave(sender As Object, e As EventArgs)
        Dim needRepaint = (catHoveredIdx >= 0 OrElse hoveredCatHeader IsNot Nothing)
        catHoveredIdx = -1
        hoveredCatHeader = Nothing
        If needRepaint Then pnlCategories.Invalidate()
    End Sub

    Private Sub CategoriesRightClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Right Then Return
        Dim hit = GetCatEntryAt(e.Location)

        ' Right-click on a custom category header — show rename/delete
        If hit.HeaderName IsNot Nothing AndAlso customCategoryNames.Contains(hit.HeaderName) Then
            Dim catName = hit.HeaderName
            Dim cm = New ContextMenuStrip()
            cm.Items.Add("Rename category", Nothing, Sub(s, ev)
                                                         Dim newName = InputBox("New name:", "Rename Category", catName)
                                                         If Not String.IsNullOrEmpty(newName) AndAlso newName <> catName Then
                                                             RenameCustomCategory(catName, newName)
                                                         End If
                                                     End Sub)
            cm.Items.Add("Delete category", Nothing, Sub(s, ev)
                                                         If MessageBox.Show("Delete custom category """ & catName & """?",
                                                                            "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                                             RemoveCustomCategory(catName)
                                                         End If
                                                     End Sub)
            cm.Show(pnlCategories, e.Location)
            Return
        End If

        ' Right-click on an app card inside a custom category — show remove option
        If hit.Entry IsNot Nothing Then
            ' Find which category this entry belongs to by checking the layout
            Dim layout = GetCatGroupLayout()
            Dim catIdx = 0
            For Each grp In layout
                For Each ent In grp.Entries
                    If catIdx = hit.Index Then
                        If customCategoryNames.Contains(grp.Name) Then
                            Dim grpName = grp.Name
                            Dim entryKey = hit.Entry.Key
                            Dim cm = New ContextMenuStrip()
                            cm.Items.Add("Remove from " & grpName, Nothing, Sub(s, ev)
                                                                                RemoveAppFromCustomCategory(grpName, entryKey)
                                                                            End Sub)
                            cm.Items.Add("Open", Nothing, Sub(s, ev) LaunchEntry(hit.Entry))
                            cm.Show(pnlCategories, e.Location)
                            Return
                        End If
                    End If
                    catIdx += 1
                Next
            Next
        End If
    End Sub

    Private Sub RenameCustomCategory(oldName As String, newName As String)
        If Not customCategoryNames.Contains(oldName) Then Return
        If customCategoryNames.Contains(newName) Then Return
        customCategoryNames.Remove(oldName)
        customCategoryNames.Add(newName)
        For i = 0 To categories.Count - 1
            If String.Equals(categories(i).Name, oldName, StringComparison.OrdinalIgnoreCase) Then
                categories(i) = (newName, categories(i).Entries)
                Exit For
            End If
        Next
        ' Update expanded/scroll state
        If expandedCategories.Contains(oldName) Then
            expandedCategories.Remove(oldName)
            expandedCategories.Add(newName)
        End If
        If catScrollOffsets.ContainsKey(oldName) Then
            catScrollOffsets(newName) = catScrollOffsets(oldName)
            catScrollOffsets.Remove(oldName)
        End If
        SaveCustomCategories()
        pnlCategories.Invalidate()
    End Sub

    ' --- Category drag-and-drop handlers ---
    Private Function GetDropTargetCategory(pt As Point) As String
        ' Check if hovering over the "+" new category drop zone (beside "Categories" header)
        ' The "+" zone is a small area to the right of the "Categories" text
        Dim newZoneRect = New Rectangle(LEFT_CONTENT_W - 80, 0, 76, 30)
        If newZoneRect.Contains(pt) Then Return "##NEW##"

        ' Check existing custom category groups
        Dim layout = GetCatGroupLayout()
        For Each grp In layout
            If customCategoryNames.Contains(grp.Name) AndAlso grp.Rect.Contains(pt) Then
                Return grp.Name
            End If
        Next
        Return Nothing
    End Function

    Private Sub CategoriesDragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(GetType(AppEntry)) Then
            e.Effect = If(dragSourceCat IsNot Nothing, DragDropEffects.Move, DragDropEffects.Copy)
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub CategoriesDragOver(sender As Object, e As DragEventArgs)
        If Not e.Data.GetDataPresent(GetType(AppEntry)) Then
            e.Effect = DragDropEffects.None
            Return
        End If

        Dim pt = pnlCategories.PointToClient(New Point(e.X, e.Y))
        Dim target = GetDropTargetCategory(pt)

        Dim dropEffect = If(dragSourceCat IsNot Nothing, DragDropEffects.Move, DragDropEffects.Copy)
        Dim changed = False
        If target = "##NEW##" Then
            If Not dragDropNewTarget Then changed = True
            dragDropNewTarget = True
            dragDropTargetCat = Nothing
            e.Effect = dropEffect
        ElseIf target IsNot Nothing Then
            ' Don't allow dropping on the same category
            If dragSourceCat IsNot Nothing AndAlso String.Equals(target, dragSourceCat, StringComparison.OrdinalIgnoreCase) Then
                If dragDropTargetCat IsNot Nothing OrElse dragDropNewTarget Then changed = True
                dragDropTargetCat = Nothing
                dragDropNewTarget = False
                e.Effect = DragDropEffects.None
            Else
                If dragDropTargetCat <> target OrElse dragDropNewTarget Then changed = True
                dragDropTargetCat = target
                dragDropNewTarget = False
                e.Effect = dropEffect
            End If
        Else
            If dragDropTargetCat IsNot Nothing OrElse dragDropNewTarget Then changed = True
            dragDropTargetCat = Nothing
            dragDropNewTarget = False
            e.Effect = DragDropEffects.None
        End If

        If changed Then pnlCategories.Invalidate()
    End Sub

    Private Sub CategoriesDragLeave(sender As Object, e As EventArgs)
        If dragDropTargetCat IsNot Nothing OrElse dragDropNewTarget Then
            dragDropTargetCat = Nothing
            dragDropNewTarget = False
            pnlCategories.Invalidate()
        End If
    End Sub

    Private Sub CategoriesDragDrop(sender As Object, e As DragEventArgs)
        Dim sourceCat = dragSourceCat  ' capture before reset
        dragDropTargetCat = Nothing
        dragDropNewTarget = False

        If Not e.Data.GetDataPresent(GetType(AppEntry)) Then Return
        Dim entry = DirectCast(e.Data.GetData(GetType(AppEntry)), AppEntry)

        Dim pt = pnlCategories.PointToClient(New Point(e.X, e.Y))
        Dim target = GetDropTargetCategory(pt)

        ' Don't drop onto the same category it came from
        If target IsNot Nothing AndAlso target <> "##NEW##" AndAlso
           sourceCat IsNot Nothing AndAlso String.Equals(target, sourceCat, StringComparison.OrdinalIgnoreCase) Then
            pnlCategories.Invalidate()
            Return
        End If

        If target = "##NEW##" Then
            Dim name = InputBox("Category name:", "New Custom Category")
            If Not String.IsNullOrEmpty(name) Then
                CreateCustomCategory(name)
                AddAppToCustomCategory(name, entry)
                ' Remove from source if it was a move between custom categories
                If sourceCat IsNot Nothing Then
                    RemoveAppFromCustomCategory(sourceCat, entry.Key)
                End If
            End If
        ElseIf target IsNot Nothing Then
            AddAppToCustomCategory(target, entry)
            ' Remove from source if it was a move between custom categories
            If sourceCat IsNot Nothing Then
                RemoveAppFromCustomCategory(sourceCat, entry.Key)
            End If
        End If

        pnlCategories.Invalidate()
    End Sub

    ' --- All Apps mouse handling ---
    Private Function GetAllAppsEntryAt(pt As Point) As (Index As Integer, Entry As AppEntry, LetterHeader As String)
        If pt.Y < 34 Then Return (-1, Nothing, Nothing)
        Dim query = txtSearch.Text.Trim()
        Dim filtered = allEntries.Where(Function(a) String.IsNullOrEmpty(query) OrElse a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(Function(a) a.Name).ToList()

        Dim y = 34 - allAppsScrollOffset
        Dim currentLetter As String = ""
        For i = 0 To filtered.Count - 1
            Dim entry = filtered(i)
            Dim firstLetter = entry.Name.Substring(0, 1).ToUpper()
            If Not Char.IsLetter(firstLetter(0)) Then firstLetter = "#"
            If firstLetter <> currentLetter Then
                currentLetter = firstLetter
                ' Check if click is on the letter header row
                Dim letterRect = New Rectangle(0, y, pnlAllApps.ClientSize.Width, ALLAPPS_LETTER_H)
                If letterRect.Contains(pt) Then Return (-1, Nothing, currentLetter)
                y += ALLAPPS_LETTER_H
            End If
            Dim rowRect = New Rectangle(4, y, pnlAllApps.ClientSize.Width - 8, ALLAPPS_ROW_H)
            If rowRect.Contains(pt) Then Return (i, entry, Nothing)
            y += ALLAPPS_ROW_H
        Next
        Return (-1, Nothing, Nothing)
    End Function

    Private Function GetAvailableLetters() As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim query = txtSearch.Text.Trim()
        For Each entry In allEntries
            If Not String.IsNullOrEmpty(query) AndAlso entry.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
            Dim fl = entry.Name.Substring(0, 1).ToUpper()
            If Not Char.IsLetter(fl(0)) Then fl = "#"
            result.Add(fl)
        Next
        Return result
    End Function

    Private Function GetScrollOffsetForLetter(letter As String) As Integer
        Dim query = txtSearch.Text.Trim()
        Dim filtered = allEntries.Where(Function(a) String.IsNullOrEmpty(query) OrElse a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(Function(a) a.Name).ToList()
        Dim y = 0
        Dim currentLetter As String = ""
        For Each entry In filtered
            Dim fl = entry.Name.Substring(0, 1).ToUpper()
            If Not Char.IsLetter(fl(0)) Then fl = "#"
            If fl <> currentLetter Then
                currentLetter = fl
                If currentLetter = letter Then Return y
                y += ALLAPPS_LETTER_H
            End If
            y += ALLAPPS_ROW_H
        Next
        Return 0
    End Function

    Private Sub AllAppsClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left Then Return

        ' Handle letter grid overlay clicks
        If showLetterGrid Then
            Dim gridIdx = GetLetterGridIndexAt(e.Location)
            If gridIdx >= 0 Then
                Dim allLetters = {"#", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                                  "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
                Dim available = GetAvailableLetters()
                If gridIdx < allLetters.Length AndAlso available.Contains(allLetters(gridIdx)) Then
                    allAppsScrollOffset = GetScrollOffsetForLetter(allLetters(gridIdx))
                End If
            End If
            showLetterGrid = False
            letterGridHoveredIdx = -1
            pnlAllApps.Invalidate()
            Return
        End If

        ' Check if clicking a letter header — show the letter grid
        Dim hit = GetAllAppsEntryAt(e.Location)
        If hit.LetterHeader IsNot Nothing Then
            showLetterGrid = True
            pnlAllApps.Invalidate()
            Return
        End If

        If hit.Entry IsNot Nothing Then LaunchEntry(hit.Entry)
    End Sub

    Private Function GetLetterGridIndexAt(pt As Point) As Integer
        ' Grid layout: 7 columns, starts at y=40, each cell 30x30 with 4px gap
        Dim gridX = (pnlAllApps.ClientSize.Width - (7 * 34)) \ 2
        Dim gridY = 40
        Dim cellSize = 30
        Dim cellGap = 4

        Dim col = (pt.X - gridX) \ (cellSize + cellGap)
        Dim row = (pt.Y - gridY) \ (cellSize + cellGap)
        If col < 0 OrElse col >= 7 OrElse row < 0 Then Return -1

        ' Check if actually inside the cell (not in the gap)
        Dim cellLeft = gridX + col * (cellSize + cellGap)
        Dim cellTop = gridY + row * (cellSize + cellGap)
        If pt.X < cellLeft OrElse pt.X > cellLeft + cellSize Then Return -1
        If pt.Y < cellTop OrElse pt.Y > cellTop + cellSize Then Return -1

        Dim idx = row * 7 + col
        If idx > 27 Then Return -1 ' # + A-Z = 28 items
        Return idx
    End Function

    Private Sub AllAppsMouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            dragStartPoint = e.Location
        End If
    End Sub

    Private Sub AllAppsMouseMove(sender As Object, e As MouseEventArgs)
        If showLetterGrid Then
            Dim gridIdx = GetLetterGridIndexAt(e.Location)
            If gridIdx <> letterGridHoveredIdx Then
                letterGridHoveredIdx = gridIdx
                pnlAllApps.Invalidate()
            End If
            pnlAllApps.Cursor = If(gridIdx >= 0, Cursors.Hand, Cursors.Default)
            Return
        End If

        ' Start drag if mouse moved far enough with button held
        If e.Button = MouseButtons.Left AndAlso dragStartPoint <> Point.Empty AndAlso Not isDragging Then
            Dim dx = Math.Abs(e.X - dragStartPoint.X)
            Dim dy = Math.Abs(e.Y - dragStartPoint.Y)
            If dx > SystemInformation.DragSize.Width OrElse dy > SystemInformation.DragSize.Height Then
                Dim hit = GetAllAppsEntryAt(dragStartPoint)
                If hit.Entry IsNot Nothing Then
                    isDragging = True
                    pnlCategories.Invalidate()  ' Show the "+ New" drop zone
                    pnlAllApps.DoDragDrop(hit.Entry, DragDropEffects.Copy)
                    isDragging = False
                    dragStartPoint = Point.Empty
                    pnlCategories.Invalidate()  ' Hide it when done
                    Return
                End If
            End If
        End If

        Dim hitMove = GetAllAppsEntryAt(e.Location)
        If hitMove.Index <> allAppsHoveredIdx Then
            allAppsHoveredIdx = hitMove.Index
            pnlAllApps.Invalidate()
        End If
        Dim showHand = hitMove.Index >= 0 OrElse hitMove.LetterHeader IsNot Nothing
        pnlAllApps.Cursor = If(showHand, Cursors.Hand, Cursors.Default)
    End Sub

    Private Sub AllAppsMouseLeave(sender As Object, e As EventArgs)
        Dim needRepaint = (allAppsHoveredIdx >= 0 OrElse letterGridHoveredIdx >= 0)
        allAppsHoveredIdx = -1
        letterGridHoveredIdx = -1
        If needRepaint Then pnlAllApps.Invalidate()
    End Sub

    Private Sub AllAppsWheel(sender As Object, e As MouseEventArgs)
        If showLetterGrid Then
            showLetterGrid = False
            letterGridHoveredIdx = -1
        End If
        allAppsScrollOffset -= e.Delta \ 4
        allAppsScrollOffset = Math.Max(0, allAppsScrollOffset)
        ' Clamp to content height
        Dim totalHeight = EstimateAllAppsHeight()
        Dim visibleH = pnlAllApps.ClientSize.Height - 34
        allAppsScrollOffset = Math.Min(allAppsScrollOffset, Math.Max(0, totalHeight - visibleH))
        pnlAllApps.Invalidate()
    End Sub

    Private Function EstimateAllAppsHeight() As Integer
        Dim query = txtSearch.Text.Trim()
        Dim filtered = allEntries.Where(Function(a) String.IsNullOrEmpty(query) OrElse a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(Function(a) a.Name).ToList()
        Dim h = 0
        Dim currentLetter As String = ""
        For Each entry In filtered
            Dim firstLetter = entry.Name.Substring(0, 1).ToUpper()
            If Not Char.IsLetter(firstLetter(0)) Then firstLetter = "#"
            If firstLetter <> currentLetter Then
                currentLetter = firstLetter
                h += ALLAPPS_LETTER_H
            End If
            h += ALLAPPS_ROW_H
        Next
        Return h
    End Function

    ' --- All Apps right-click context menu ---
    Private Sub AllAppsRightClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Right Then Return
        Dim hit = GetAllAppsEntryAt(e.Location)
        If hit.Entry Is Nothing Then Return

        Dim entry = hit.Entry
        Dim cm = New ContextMenuStrip()
        '    cm.Renderer = New ToolStripProfessionalRenderer(New DarkColorTable())

        Dim isPinned = pinnedEntries.Any(Function(p) p.Key = entry.Key)
        If isPinned Then
            cm.Items.Add("Unpin from Start", Nothing, Sub(s, ev)
                                                          pinnedEntries.RemoveAll(Function(p) p.Key = entry.Key)
                                                          SavePins()
                                                          RebuildPinnedHeight()
                                                          pnlPinned.Invalidate()
                                                      End Sub)
        Else
            cm.Items.Add("Pin to Start", Nothing, Sub(s, ev)
                                                      If pinnedEntries.Count < 18 Then
                                                          pinnedEntries.Add(entry)
                                                          SavePins()
                                                          RebuildPinnedHeight()
                                                          pnlPinned.Invalidate()
                                                      End If
                                                  End Sub)
        End If

        ' Add to category submenu
        Dim catMenu = New ToolStripMenuItem("Add to category")
        For Each cat In categories
            If Not customCategoryNames.Contains(cat.Name) Then Continue For
            Dim catName = cat.Name ' capture for lambda
            Dim alreadyIn = cat.Entries.Any(Function(en) en.Key = entry.Key)
            Dim item = catMenu.DropDownItems.Add(catName)
            If alreadyIn Then
                item.Enabled = False
                item.Text = catName & " (already added)"
            Else
                AddHandler item.Click, Sub(s, ev) AddAppToCustomCategory(catName, entry)
            End If
        Next
        If catMenu.DropDownItems.Count > 0 Then
            catMenu.DropDownItems.Add(New ToolStripSeparator())
        End If
        catMenu.DropDownItems.Add("New category...", Nothing, Sub(s, ev)
                                                                  Dim name = InputBox("Category name:", "New Custom Category")
                                                                  If Not String.IsNullOrEmpty(name) Then
                                                                      CreateCustomCategory(name)
                                                                      AddAppToCustomCategory(name, entry)
                                                                  End If
                                                              End Sub)
        cm.Items.Add(catMenu)
        cm.Items.Add(New ToolStripSeparator())

        cm.Items.Add("Open", Nothing, Sub(s, ev) LaunchEntry(entry))
        cm.Items.Add("Open file location", Nothing, Sub(s, ev)
                                                        If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso IO.File.Exists(entry.ShortcutPath) Then
                                                            Process.Start("explorer.exe", "/select,""" & entry.ShortcutPath & """")
                                                        End If
                                                    End Sub)

        cm.Show(pnlAllApps, e.Location)
    End Sub

    Private Sub RebuildPinnedHeight()
        ' Recalculate the pinned panel height based on how many entries there are
        Dim rows = Math.Max(1, CInt(Math.Ceiling(pinnedEntries.Count / Math.Floor(LEFT_CONTENT_W / (CARD_W + CARD_GAP)))))
        Dim neededH = 36 + rows * (CARD_H + CARD_GAP) + 10
        If neededH <> pnlPinned.Height Then pnlPinned.Height = neededH
    End Sub

    ' --- User bar interaction ---
    Private Sub UserBarClick(sender As Object, e As MouseEventArgs)
        ' Mica toggle switch area (pill + label)
        Dim toggleX = pnlUserBar.Width - 160
        If e.X >= toggleX AndAlso e.X <= toggleX + 80 AndAlso e.Y >= 16 AndAlso e.Y <= 44 Then
            ToggleMica()
            Return
        End If
        ' Power button area (right side)
        If e.X > pnlUserBar.Width - 60 Then
            ' Show a simple power menu
            Dim cm = New ContextMenuStrip()
            cm.Items.Add("Sleep", Nothing, Sub(s, ev) Application.SetSuspendState(PowerState.Suspend, False, False))
            cm.Items.Add("Shut down", Nothing, Sub(s, ev)
                                                   Try : Process.Start("shutdown", "/s /t 0") : Catch : End Try
                                               End Sub)
            cm.Items.Add("Restart", Nothing, Sub(s, ev)
                                                 Try : Process.Start("shutdown", "/r /t 0") : Catch : End Try
                                             End Sub)
            cm.Items.Add("Sign out", Nothing, Sub(s, ev)
                                                  Try : Process.Start("logoff") : Catch : End Try
                                              End Sub)
            cm.Items.Add("-", Nothing, Sub(s, ev)
                                           Try : Catch : End Try
                                       End Sub)

            cm.Items.Add("Exit Rice", Nothing, Sub(s, ev) Application.Exit())

            cm.Show(pnlUserBar, e.Location)
        End If
    End Sub
#End Region

    Private Sub ToggleMica()
        micaEnabled = Not micaEnabled
        If micaEnabled Then
            ' Add WS_CAPTION so DWM can render the backdrop
            Dim style = GetWindowLong(Me.Handle, GWL_STYLE)
            SetWindowLong(Me.Handle, GWL_STYLE, style Or WS_CAPTION)
            SetWindowPos(Me.Handle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
            ApplyMicaBackdrop()
        Else
            ' Disable Mica
            micaApplied = False
            Try
                ' Remove backdrop
                Dim backdropType As Integer = 0 ' DWMSBT_NONE
                DwmSetWindowAttribute(Me.Handle, DWMWA_SYSTEMBACKDROP_TYPE, backdropType, 4)

                ' Also try the legacy attribute
                Dim micaOff As Integer = 0
                DwmSetWindowAttribute(Me.Handle, DWMWA_MICA_EFFECT, micaOff, 4)

                ' Retract DWM frame so it stops being see-through
                Dim m As New MARGINS() With {.Left = 0, .Right = 0, .Top = 0, .Bottom = 0}
                DwmExtendFrameIntoClientArea(Me.Handle, m)

                ' Remove WS_CAPTION so DWM stops managing the frame
                Dim style = GetWindowLong(Me.Handle, GWL_STYLE)
                SetWindowLong(Me.Handle, GWL_STYLE, style And Not WS_CAPTION)
                SetWindowPos(Me.Handle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
            Catch
            End Try
        End If
        Me.Invalidate(True)
        pnlUserBar.Invalidate()
    End Sub

    Private Sub InitializeComponent()
    End Sub

    Private Sub frmStart_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub
End Class
