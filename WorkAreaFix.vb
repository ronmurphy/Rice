Imports System.Runtime.InteropServices

Module WorkAreaFix

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll")>
    Private Function SystemParametersInfo(
        uiAction As UInteger,
        uiParam As UInteger,
        ByRef pvParam As RECT,
        fWinIni As UInteger) As Boolean
    End Function

    Private Const SPI_GETWORKAREA As UInteger = &H30
    Private Const SPI_SETWORKAREA As UInteger = &H2F
    Private Const SPIF_UPDATEINIFILE As UInteger = &H1
    Private Const SPIF_SENDCHANGE As UInteger = &H2

    ' Saved original work area so we can restore on exit
    Private _originalSaved As Boolean = False
    Private _originalWorkArea As RECT

    ''' <summary>
    ''' Save the current work area before we modify it, so we can restore on exit.
    ''' </summary>
    Public Sub SaveOriginalWorkArea()
        If Not _originalSaved Then
            SystemParametersInfo(SPI_GETWORKAREA, 0, _originalWorkArea, 0)
            _originalSaved = True
        End If
    End Sub

    ''' <summary>
    ''' Set work area to full screen minus a top offset (for the Rice bar).
    ''' </summary>
    Public Sub SetWorkAreaExcludeTop(topOffset As Integer)
        Dim scr = Screen.PrimaryScreen.Bounds

        Dim r As RECT
        r.Left = scr.Left
        r.Top = scr.Top + topOffset
        r.Right = scr.Right
        r.Bottom = scr.Bottom

        SystemParametersInfo(SPI_SETWORKAREA, 0, r,
                             SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)
    End Sub

    ''' <summary>
    ''' Expand work area to full screen (no reservations).
    ''' </summary>
    Public Sub ExpandWorkAreaToFullScreen()
        Dim scr = Screen.PrimaryScreen.Bounds

        Dim r As RECT
        r.Left = scr.Left
        r.Top = scr.Top
        r.Right = scr.Right
        r.Bottom = scr.Bottom

        SystemParametersInfo(SPI_SETWORKAREA, 0, r,
                             SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)
    End Sub

    ''' <summary>
    ''' Restore the original work area (call on exit).
    ''' </summary>
    Public Sub RestoreOriginalWorkArea()
        If _originalSaved Then
            SystemParametersInfo(SPI_SETWORKAREA, 0, _originalWorkArea,
                                 SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)
        End If
    End Sub

End Module
