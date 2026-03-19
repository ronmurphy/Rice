Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class frmStart
    Inherits Form

    Private lstApps As ListView
    Private imgList As ImageList
    Private txtSearch As TextBox
    Private flowTiles As FlowLayoutPanel
    Private flowRight As FlowLayoutPanel
    Private chkRightTiles As CheckBox
    Private rightAsTiles As Boolean = False
    Private leftPanel As Panel
    Private rightPanel As Panel
    Private allEntries As New List(Of AppEntry)()

    Private Class AppEntry
        Public Property Name As String
        Public Property Key As String
        Public Property ShortcutPath As String
        Public Property Group As String
    End Class

    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.Width = 800
        Me.Height = 700
        Me.Text = "Start"
        Me.BackColor = Color.White

        txtSearch = New TextBox()
        txtSearch.Dock = DockStyle.Top
        txtSearch.Height = 32
        txtSearch.PlaceholderText = "Search apps..."
        AddHandler txtSearch.TextChanged, AddressOf OnSearchTextChanged
        Me.Controls.Add(txtSearch)

        Dim mainPanel As New Panel()
        mainPanel.Dock = DockStyle.Fill
        Me.Controls.Add(mainPanel)

        ' Left navigation panel + right content panel layout
        leftPanel = New Panel()
        leftPanel.Dock = DockStyle.Left
        leftPanel.Width = 220
        leftPanel.Padding = New Padding(8)

        rightPanel = New Panel()
        rightPanel.Dock = DockStyle.Fill

        imgList = New ImageList()
        imgList.ImageSize = New Size(32, 32)

        lstApps = New ListView()
        lstApps.Dock = DockStyle.Left
        lstApps.Width = 420
        lstApps.View = View.Tile
        lstApps.TileSize = New Size(380, 40)
        lstApps.LargeImageList = imgList
        lstApps.MultiSelect = False
        lstApps.ShowGroups = True

        flowTiles = New FlowLayoutPanel()
        flowTiles.Dock = DockStyle.Top
        flowTiles.Height = 260
        flowTiles.FlowDirection = FlowDirection.LeftToRight
        flowTiles.WrapContents = True
        flowTiles.Padding = New Padding(8)

        ' Right side panel for quick icons or tiles
        flowRight = New FlowLayoutPanel()
        flowRight.Dock = DockStyle.Right
        flowRight.Width = 200
        flowRight.FlowDirection = FlowDirection.TopDown
        flowRight.WrapContents = True
        flowRight.Padding = New Padding(6)
        flowRight.AutoScroll = True
        flowRight.BorderStyle = BorderStyle.FixedSingle

        chkRightTiles = New CheckBox()
        chkRightTiles.Text = "Right side tiles"
        chkRightTiles.Dock = DockStyle.Top
        chkRightTiles.AutoSize = True
        chkRightTiles.Padding = New Padding(4)
        chkRightTiles.Checked = rightAsTiles
        AddHandler chkRightTiles.CheckedChanged, AddressOf OnRightTilesToggled

        ' Build left navigation and right content
        leftPanel.Controls.Add(flowRight)
        flowRight.Controls.Add(chkRightTiles)
        leftPanel.Controls.Add(txtSearch)

        rightPanel.Controls.Add(flowTiles)
        lstApps.Dock = DockStyle.Fill
        rightPanel.Controls.Add(lstApps)

        mainPanel.Controls.Add(rightPanel)
        mainPanel.Controls.Add(leftPanel)

        LoadStartMenuApps()
        LoadPinnedTiles()

        AddHandler lstApps.ItemActivate, AddressOf OnAppActivated
    End Sub

    ' Interop for shell image factory to get UWP app icons
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

    Private Sub MakeBlackTransparent(bmp As Bitmap)
        If bmp.PixelFormat <> PixelFormat.Format32bppPArgb AndAlso bmp.PixelFormat <> PixelFormat.Format32bppArgb Then
            ' ensure 32bpp
            Dim copy = New Bitmap(bmp.Width, bmp.Height, PixelFormat.Format32bppPArgb)
            Using g = Graphics.FromImage(copy)
                g.DrawImage(bmp, New Rectangle(0, 0, copy.Width, copy.Height))
            End Using
            bmp.Dispose()
            bmp = copy
        End If

        Dim rect = New Rectangle(0, 0, bmp.Width, bmp.Height)
        Dim data = bmp.LockBits(rect, Imaging.ImageLockMode.ReadWrite, bmp.PixelFormat)
        Try
            Dim bytes = Math.Abs(data.Stride) * bmp.Height
            Dim buffer(bytes - 1) As Byte
            Marshal.Copy(data.Scan0, buffer, 0, buffer.Length)
            For i = 0 To buffer.Length - 4 Step 4
                Dim b = buffer(i)
                Dim g = buffer(i + 1)
                Dim r = buffer(i + 2)
                ' if pixel is pure black (0,0,0) make it transparent
                If b = 0 AndAlso g = 0 AndAlso r = 0 Then
                    buffer(i + 3) = 0 ' alpha
                End If
            Next
            Marshal.Copy(buffer, 0, data.Scan0, buffer.Length)
        Finally
            bmp.UnlockBits(data)
        End Try
    End Sub

    <DllImport("gdi32.dll")>
    Private Shared Function DeleteObject(hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Function GetShellIconBitmap(parsingName As String, size As Integer) As Bitmap
        Try
            Dim iid As New Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")
            Dim pUnk As IntPtr = IntPtr.Zero
            Dim hr = SHCreateItemFromParsingName(parsingName, IntPtr.Zero, iid, pUnk)
            If hr <> 0 OrElse pUnk = IntPtr.Zero Then Return Nothing

            Dim factory = CType(Marshal.GetObjectForIUnknown(pUnk), IShellItemImageFactory)
            Marshal.Release(pUnk)

            Dim hbm As IntPtr = IntPtr.Zero
            Dim s As New SHSIZE(size, size)
            Const SIIGBF_RESIZETOFIT As UInteger = 0
            Dim res = factory.GetImage(s, SIIGBF_RESIZETOFIT, hbm)
            If res = 0 AndAlso hbm <> IntPtr.Zero Then
                Dim bmp As Bitmap = Nothing
                Try
                    bmp = Bitmap.FromHbitmap(hbm)
                    ' convert to 32bpp to avoid palette issues
                    Dim copy = New Bitmap(bmp.Width, bmp.Height, PixelFormat.Format32bppPArgb)
                    Using g = Graphics.FromImage(copy)
                        g.DrawImage(bmp, New Rectangle(0, 0, copy.Width, copy.Height))
                    End Using
                    Try
                        MakeBlackTransparent(copy)
                    Catch
                    End Try
                    Return copy
                Finally
                    If bmp IsNot Nothing Then bmp.Dispose()
                    DeleteObject(hbm)
                End Try
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Sub OnSearchTextChanged(sender As Object, e As EventArgs)
        Dim q = txtSearch.Text.Trim()
        ApplyFilter(q)
    End Sub

    Private Sub ApplyFilter(query As String)
        lstApps.BeginUpdate()
        lstApps.Items.Clear()
        lstApps.Groups.Clear()

        Dim filtered = allEntries.Where(Function(a) String.IsNullOrEmpty(query) OrElse a.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).ToList()
        Dim groups = filtered.Select(Function(a) a.Group).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(Function(x) x).ToList()
        For Each g In groups
            lstApps.Groups.Add(New ListViewGroup(g, g))
        Next

        For Each e In filtered.OrderBy(Function(x) x.Name)
            Dim item = New ListViewItem(e.Name)
            item.Tag = e.ShortcutPath
            item.ImageKey = e.Key
            Dim grp = lstApps.Groups.Cast(Of ListViewGroup)().FirstOrDefault(Function(gg) String.Equals(gg.Header, e.Group, StringComparison.OrdinalIgnoreCase))
            If grp IsNot Nothing Then item.Group = grp
            lstApps.Items.Add(item)
        Next

        lstApps.EndUpdate()
    End Sub

    Private Sub LoadStartMenuApps()
        ' Use StartMenu roots (which include Programs as a subfolder) plus Desktop.
        Dim roots As New List(Of String)()
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
        roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))

        ' Only scan for .lnk files — CreateShortcut only works on .lnk
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

        ' Group roots for determining folder group names (longest match first)
        Dim groupRoots = New String() {
            Environment.GetFolderPath(Environment.SpecialFolder.Programs),
            Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms),
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
            Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
        }.Where(Function(r) Not String.IsNullOrEmpty(r)).OrderByDescending(Function(r) r.Length).ToArray()

        ' Resolve shortcuts and deduplicate by stable key (exe path or appsfolder id)
        Dim resolved As New Dictionary(Of String, AppEntry)(StringComparer.OrdinalIgnoreCase)
        For Each f In files
            Try
                Dim shell = CreateObject("WScript.Shell")
                Dim lnk = shell.CreateShortcut(f)
                Dim target = If(lnk.TargetPath, String.Empty)
                Dim args = If(lnk.Arguments, String.Empty)

                ' Display name is always the shortcut filename
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
                        Try
                            key = Path.GetFullPath(target)
                        Catch
                            key = target
                        End Try
                    End If
                ElseIf Not String.IsNullOrEmpty(args) AndAlso args.IndexOf("appsfolder", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    key = args.Trim()
                End If

                ' Fallback: use the shortcut file path as key (for .lnk with empty target, e.g. File Explorer, Control Panel)
                If String.IsNullOrEmpty(key) Then
                    key = f
                End If

                ' Determine group from shortcut folder
                Dim groupName = "Programs"
                Try
                    Dim dir = Path.GetDirectoryName(f)
                    For Each r In groupRoots
                        If dir.StartsWith(r, StringComparison.OrdinalIgnoreCase) Then
                            Dim rel = dir.Substring(r.Length).TrimStart(Path.DirectorySeparatorChar)
                            If Not String.IsNullOrEmpty(rel) Then
                                groupName = rel.Replace(Path.DirectorySeparatorChar, " "c)
                            End If
                            Exit For
                        End If
                    Next
                Catch
                End Try

                ' Prefer user StartMenu entry over common when same key exists
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

        ' Also enumerate installed Store/UWP apps via shell:AppsFolder
        ' These don't have .lnk files but are visible in the Windows Start menu
        Try
            Dim shellApp = CreateObject("Shell.Application")
            Dim appsFolder = shellApp.NameSpace("shell:AppsFolder")
            If appsFolder IsNot Nothing Then
                For Each item In appsFolder.Items()
                    Try
                        Dim appName As String = item.Name
                        Dim appPath As String = item.Path
                        If String.IsNullOrEmpty(appName) OrElse String.IsNullOrEmpty(appPath) Then Continue For

                        ' Skip if we already have this app from .lnk scanning (by name or by appsfolder key)
                        If resolved.ContainsKey(appPath) Then Continue For

                        ' Build a launch path: shell:AppsFolder\{appPath}
                        Dim launchPath = "shell:AppsFolder\" & appPath
                        resolved(appPath) = New AppEntry() With {
                            .Name = appName,
                            .Key = appPath,
                            .ShortcutPath = launchPath,
                            .Group = "Apps"
                        }
                    Catch
                    End Try
                Next
            End If
        Catch
        End Try

        ' Populate internal list, deduplicating by display name
        allEntries.Clear()
        Dim entries = resolved.Values.ToList()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each entry In entries.OrderBy(Function(x) x.Name)
            If seen.Contains(entry.Name) Then Continue For
            seen.Add(entry.Name)
            allEntries.Add(entry)
            ' Add icon for the entry
            Try
                If imgList.Images.ContainsKey(entry.Key) Then
                    ' already added
                ElseIf File.Exists(entry.Key) Then
                    Dim ico = Icon.ExtractAssociatedIcon(entry.Key)
                    imgList.Images.Add(entry.Key, ico.ToBitmap())
                Else
                    ' For Store apps and shortcuts with non-file keys, try a few fallbacks:
                    Dim gotIcon = False
                    ' 1) Extract icon from the shortcut file itself (if present)
                    If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso entry.ShortcutPath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) AndAlso File.Exists(entry.ShortcutPath) Then
                        Try
                            Dim ico = Icon.ExtractAssociatedIcon(entry.ShortcutPath)
                            If ico IsNot Nothing Then
                                imgList.Images.Add(entry.Key, ico.ToBitmap())
                                gotIcon = True
                            End If
                        Catch
                        End Try
                    End If
                    ' 2) If this is a shell:AppsFolder entry (UWP/Store app), try Shell image factory
                    If Not gotIcon Then
                        Try
                            Dim parsingName As String = Nothing
                            If Not String.IsNullOrEmpty(entry.ShortcutPath) AndAlso entry.ShortcutPath.StartsWith("shell:AppsFolder", StringComparison.OrdinalIgnoreCase) Then
                                parsingName = entry.ShortcutPath
                            Else
                                parsingName = "shell:AppsFolder\" & entry.Key
                            End If
                            Dim bmpUwp = GetShellIconBitmap(parsingName, imgList.ImageSize.Width)
                            If bmpUwp IsNot Nothing Then
                                imgList.Images.Add(entry.Key, bmpUwp)
                                gotIcon = True
                            End If
                        Catch
                        End Try
                    End If
                    ' 3) Final fallback: generic placeholder
                    If Not gotIcon Then
                        Dim bmp = New Bitmap(imgList.ImageSize.Width, imgList.ImageSize.Height)
                        Using gr = Graphics.FromImage(bmp)
                            gr.Clear(Color.LightGray)
                        End Using
                        imgList.Images.Add(entry.Key, bmp)
                    End If
                End If
            Catch
            End Try
        Next

        ApplyFilter(txtSearch.Text.Trim())
    End Sub

    Private Sub LoadPinnedTiles()
        ' Try to load pinned shortcuts from taskbar pinned folder as tiles
        Try
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim taskband = IO.Path.Combine(appData, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
            Dim shortcuts As New List(Of String)
            If IO.Directory.Exists(taskband) Then shortcuts.AddRange(IO.Directory.GetFiles(taskband, "*.lnk"))

            If shortcuts.Count = 0 Then
                ' fallback: use first few entries from allEntries
                For Each e In allEntries.Take(6)
                    AddTile(e)
                Next
            Else
                For Each s In shortcuts.Distinct().Take(12)
                    Try
                        Dim key = ResolveShortcutKey(s)
                        Dim name = Path.GetFileNameWithoutExtension(s)
                        Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                        If entry IsNot Nothing Then
                            AddTile(entry)
                        Else
                            Dim entry2 As New AppEntry() With {.Name = name, .Key = key, .ShortcutPath = s, .Group = "Pinned"}
                            AddTile(entry2)
                        End If
                    Catch
                    End Try
                Next
            End If
            ' Also populate right-side quick area
            PopulateRightSide(shortcuts)

            ' Quick fix: hide center tiles area when empty
            Try
                flowTiles.Visible = (flowTiles.Controls.Count > 0)
                If rightPanel IsNot Nothing Then rightPanel.Refresh()
            Catch
            End Try
        Catch
        End Try
    End Sub

    Private Sub PopulateRightSide(shortcuts As List(Of String))
        Try
            flowRight.SuspendLayout()
            ' Remove any existing buttons except the toggle
            For i = flowRight.Controls.Count - 1 To 0 Step -1
                Dim c = flowRight.Controls(i)
                If c Is chkRightTiles Then Continue For
                flowRight.Controls.RemoveAt(i)
                c.Dispose()
            Next

            If rightAsTiles Then
                ' larger tiles vertically stacked
                If shortcuts IsNot Nothing AndAlso shortcuts.Count > 0 Then
                    For Each s In shortcuts.Distinct().Take(6)
                        Try
                            Dim key = ResolveShortcutKey(s)
                            Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                            If entry Is Nothing Then
                                entry = New AppEntry() With {.Name = Path.GetFileNameWithoutExtension(s), .Key = key, .ShortcutPath = s, .Group = "Pinned"}
                            End If
                            Dim btn = New Button()
                            Dim targetW = Math.Max(72, flowRight.ClientSize.Width - 12)
                            btn.Width = targetW
                            btn.Height = 64
                            btn.Text = entry.Name
                            btn.TextAlign = ContentAlignment.MiddleLeft
                            btn.ImageAlign = ContentAlignment.MiddleLeft
                            btn.FlatStyle = FlatStyle.Flat
                            btn.FlatAppearance.BorderSize = 0
                            btn.Margin = New Padding(4)
                            btn.BackColor = Color.Transparent
                            If imgList.Images.ContainsKey(entry.Key) Then btn.Image = imgList.Images(entry.Key)
                            AddHandler btn.Click, Sub(sa, ea)
                                                      Try
                                                          If Not String.IsNullOrEmpty(entry.ShortcutPath) Then Process.Start(New ProcessStartInfo(entry.ShortcutPath) With {.UseShellExecute = True})
                                                      Catch
                                                      End Try
                                                  End Sub
                            ' insert after the toggle checkbox
                            flowRight.Controls.Add(btn)
                        Catch
                        End Try
                    Next
                Else
                    For Each e In allEntries.Take(6)
                        AddRightIcon(e)
                    Next
                End If
            Else
                ' icons: create small square buttons
                If shortcuts IsNot Nothing AndAlso shortcuts.Count > 0 Then
                    For Each s In shortcuts.Distinct().Take(16)
                        Try
                            Dim key = ResolveShortcutKey(s)
                            Dim entry = allEntries.FirstOrDefault(Function(a) String.Equals(a.Key, key, StringComparison.OrdinalIgnoreCase))
                            If entry Is Nothing Then
                                entry = New AppEntry() With {.Name = Path.GetFileNameWithoutExtension(s), .Key = key, .ShortcutPath = s, .Group = "Pinned"}
                            End If
                            AddRightIcon(entry)
                        Catch
                        End Try
                    Next
                Else
                    For Each e In allEntries.Take(8)
                        AddRightIcon(e)
                    Next
                End If
            End If
        Finally
            flowRight.ResumeLayout()
        End Try
    End Sub

    Private Sub AddRightIcon(entry As AppEntry)
        Try
            Dim btn = New Button()
            btn.Width = 36
            btn.Height = 36
            btn.Margin = New Padding(4)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.Text = String.Empty
            btn.BackColor = Color.Transparent
            If imgList.Images.ContainsKey(entry.Key) Then
                btn.BackgroundImage = imgList.Images(entry.Key)
                btn.BackgroundImageLayout = ImageLayout.Zoom
            Else
                btn.Text = entry.Name.Substring(0, Math.Min(3, entry.Name.Length))
            End If
            AddHandler btn.Click, Sub(sa, ea)
                                      Try
                                          If Not String.IsNullOrEmpty(entry.ShortcutPath) Then Process.Start(New ProcessStartInfo(entry.ShortcutPath) With {.UseShellExecute = True})
                                      Catch
                                      End Try
                                  End Sub
            flowRight.Controls.Add(btn)
        Catch
        End Try
    End Sub

    Private Sub OnRightTilesToggled(sender As Object, e As EventArgs)
        rightAsTiles = chkRightTiles.Checked
        ' refresh right side using previously discovered pinned shortcuts if available
        Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim taskband = IO.Path.Combine(appData, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
        Dim shortcuts As New List(Of String)
        If IO.Directory.Exists(taskband) Then shortcuts.AddRange(IO.Directory.GetFiles(taskband, "*.lnk"))
        PopulateRightSide(shortcuts)
    End Sub

    Private Sub AddTile(entry As AppEntry)
        Try
            Dim btn = New Button()
            btn.Width = 120
            btn.Height = 80
            btn.Text = entry.Name
            btn.TextAlign = ContentAlignment.BottomCenter
            btn.ImageAlign = ContentAlignment.TopCenter
            btn.FlatStyle = FlatStyle.Flat
            btn.Margin = New Padding(6)
            Try
                If imgList.Images.ContainsKey(entry.Key) Then
                    btn.Image = imgList.Images(entry.Key)
                End If
            Catch
            End Try
            AddHandler btn.Click, Sub(s, e)
                                      Try
                                          If Not String.IsNullOrEmpty(entry.ShortcutPath) Then
                                              Process.Start(New ProcessStartInfo(entry.ShortcutPath) With {.UseShellExecute = True})
                                          End If
                                      Catch
                                      End Try
                                  End Sub
            flowTiles.Controls.Add(btn)
        Catch
        End Try
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

    Private Sub InitializeComponent()

    End Sub

    Private Sub OnAppActivated(sender As Object, e As EventArgs)
        If lstApps.SelectedItems.Count = 0 Then Return
        Dim item = lstApps.SelectedItems(0)
        Dim shortcut = TryCast(item.Tag, String)
        If String.IsNullOrEmpty(shortcut) Then Return
        Try
            Process.Start(New ProcessStartInfo(shortcut) With {.UseShellExecute = True})
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Failed to launch: " & ex.Message)
        End Try
    End Sub
End Class
