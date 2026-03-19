Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices

Public Class frmStart
    Inherits Form

    Private lstApps As ListView
    Private imgList As ImageList
    Private txtSearch As TextBox
    Private flowTiles As FlowLayoutPanel
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
        flowTiles.Dock = DockStyle.Fill
        flowTiles.FlowDirection = FlowDirection.LeftToRight
        flowTiles.WrapContents = True
        flowTiles.Padding = New Padding(8)

        mainPanel.Controls.Add(flowTiles)
        mainPanel.Controls.Add(lstApps)

        LoadStartMenuApps()
        LoadPinnedTiles()

        AddHandler lstApps.ItemActivate, AddressOf OnAppActivated
    End Sub

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
                    ' For Store apps and shortcuts with non-file keys, try extracting icon from the shortcut itself
                    Dim gotIcon = False
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
                    If Not gotIcon Then
                        Dim bmp = New Bitmap(32, 32)
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
        Catch
        End Try
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
