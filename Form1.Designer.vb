<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private flowLeft As FlowLayoutPanel
    Private centerPanel As Panel
    Private lblClock As Label
    Private timerClock As System.Windows.Forms.Timer
    Private panelUtilities As FlowLayoutPanel
    Private panelTray As FlowLayoutPanel
    Private panelClickArea As Panel
    Private overflowPanelInCenter As FlowLayoutPanel
    Private centerContainer As Panel
    Private panelRight As Panel

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        flowLeft = New FlowLayoutPanel()
        centerPanel = New Panel()
        lblClock = New Label()
        overflowPanelInCenter = New FlowLayoutPanel()
        centerContainer = New Panel()
        panelClickArea = New Panel()
        timerClock = New Timer(components)
        panelUtilities = New FlowLayoutPanel()
        panelTray = New FlowLayoutPanel()
        panelRight = New Panel()
        centerPanel.SuspendLayout()
        panelRight.SuspendLayout()
        SuspendLayout()
        ' 
        ' flowLeft
        ' 
        flowLeft.AutoScroll = True
        flowLeft.Dock = DockStyle.Left
        flowLeft.Location = New Point(0, 0)
        flowLeft.Name = "flowLeft"
        flowLeft.Padding = New Padding(4, 4, 4, 4)
        flowLeft.Size = New Size(300, 48)
        flowLeft.TabIndex = 2
        flowLeft.WrapContents = False
        ' 
        ' centerPanel
        ' 
        centerPanel.Controls.Add(overflowPanelInCenter)
        centerPanel.Controls.Add(centerContainer)
        centerPanel.Dock = DockStyle.Fill
        centerPanel.Location = New Point(300, 0)
        centerPanel.Name = "centerPanel"
        centerPanel.Size = New Size(500, 48)
        centerPanel.TabIndex = 1
        ' 
        ' centerContainer
        '
        centerContainer.Location = New Point(100, 4)
        centerContainer.Name = "centerContainer"
        centerContainer.Size = New Size(300, 40)
        centerContainer.TabIndex = 3

        ' lblClock
        ' 
        lblClock.Dock = DockStyle.Fill
        lblClock.Font = New Font("Segoe UI", 12F)
        lblClock.Name = "lblClock"
        lblClock.TabIndex = 0
        lblClock.TextAlign = ContentAlignment.MiddleCenter

        ' add lblClock and panelClickArea into centerContainer
        centerContainer.Controls.Add(lblClock)
        centerContainer.Controls.Add(panelClickArea)
        ' 
        ' overflowPanelInCenter
        ' 
        overflowPanelInCenter.Dock = DockStyle.Left
        overflowPanelInCenter.Location = New Point(0, 0)
        overflowPanelInCenter.Name = "overflowPanelInCenter"
        overflowPanelInCenter.Padding = New Padding(2)
        overflowPanelInCenter.Size = New Size(200, 48)
        overflowPanelInCenter.TabIndex = 1
        overflowPanelInCenter.WrapContents = False
        ' 
        ' panelClickArea
        ' 
        panelClickArea.Anchor = AnchorStyles.None
        panelClickArea.BackColor = Color.Transparent
        panelClickArea.Location = New Point(150, -26)
        panelClickArea.Name = "panelClickArea"
        panelClickArea.Size = New Size(200, 40)
        panelClickArea.TabIndex = 2
        ' 
        ' timerClock
        ' 
        timerClock.Interval = 1000
        '
        ' panelUtilities
        ' 
        panelUtilities.Dock = DockStyle.Fill
        panelUtilities.Location = New Point(0, 0)
        panelUtilities.Name = "panelUtilities"
        panelUtilities.Padding = New Padding(4, 8, 4, 8)
        panelUtilities.Size = New Size(300, 48)
        panelUtilities.TabIndex = 1
        panelUtilities.WrapContents = False
        ' 
        ' panelTray
        ' 
        panelTray.Dock = DockStyle.Right
        panelTray.Location = New Point(160, 0)
        panelTray.Name = "panelTray"
        panelTray.Padding = New Padding(4, 4, 4, 4)
        panelTray.Size = New Size(140, 48)
        panelTray.TabIndex = 0
        panelTray.WrapContents = False
        ' 
        ' panelRight
        ' 
        panelRight.Controls.Add(panelTray)
        panelRight.Controls.Add(panelUtilities)
        panelRight.Dock = DockStyle.Right
        panelRight.Location = New Point(500, 0)
        panelRight.Name = "panelRight"
        panelRight.Size = New Size(300, 48)
        panelRight.TabIndex = 0
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 48)
        Controls.Add(panelRight)
        Controls.Add(centerPanel)
        Controls.Add(flowLeft)
        FormBorderStyle = FormBorderStyle.None
        Name = "Form1"
        Text = "frmRice"
        centerPanel.ResumeLayout(False)
        panelRight.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

End Class
