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

    Private Const SPI_SETWORKAREA As UInteger = &H2F
    Private Const SPIF_UPDATEINIFILE As UInteger = &H1
    Private Const SPIF_SENDCHANGE As UInteger = &H2

    Public Sub ExpandWorkAreaToFullScreen()
        Dim wscreen = Screen.PrimaryScreen.Bounds

        Dim r As RECT
        r.Left = wscreen.Left
        r.Top = wscreen.Top
        r.Right = wscreen.Right
        r.Bottom = wscreen.Bottom

        SystemParametersInfo(SPI_SETWORKAREA, 0, r,
                             SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)
    End Sub

End Module
