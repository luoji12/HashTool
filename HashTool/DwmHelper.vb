Imports System
Imports System.Runtime.InteropServices

Namespace HashTool

    Friend Module DwmHelper

        ' --- Win11 DWM attributes ---
        Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Integer = 33
        Private Const DWMWA_BORDER_COLOR As Integer = 34
        Private Const DWMWA_SYSTEMBACKDROP_TYPE As Integer = 38

        ' Corner preference values
        Private Const DWMWCP_DEFAULT As Integer = 0
        Private Const DWMWCP_DONOTROUND As Integer = 1
        Private Const DWMWCP_ROUND As Integer = 2
        Private Const DWMWCP_ROUNDSMALL As Integer = 3

        ' Backdrop types (Win11)
        Private Const DWMSBT_MAINWINDOW As Integer = 2 ' Mica

        ' --- Win10 composition fallback ---
        Private Const WCA_ACCENT_POLICY As Integer = 19
        Private Const ACCENT_ENABLE_BLURBEHIND As Integer = 3
        Private Const ACCENT_ENABLE_ACRYLICBLURBEHIND As Integer = 4

        <DllImport("dwmapi.dll")>
        Private Function DwmSetWindowAttribute(hwnd As IntPtr, dwAttribute As Integer, ByRef pvAttribute As Integer, cbAttribute As Integer) As Integer
        End Function

        <StructLayout(LayoutKind.Sequential)>
        Private Structure ACCENT_POLICY
            Public AccentState As Integer
            Public AccentFlags As Integer
            Public GradientColor As Integer
            Public AnimationId As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure WINDOWCOMPOSITIONATTRIBDATA
            Public Attribute As Integer
            Public Data As IntPtr
            Public SizeOfData As Integer
        End Structure

        <DllImport("user32.dll")>
        Private Function SetWindowCompositionAttribute(hwnd As IntPtr, ByRef data As WINDOWCOMPOSITIONATTRIBDATA) As Integer
        End Function

        Friend Sub TryEnableBackdrop(hwnd As IntPtr)
            ' ✅ 关键：先把 Win11 系统“圆角/边框”关掉，保证跨系统一致
            TryDisableSystemRounding(hwnd)
            TrySetBorderColorTransparent(hwnd)

            ' Win11: Mica
            If TrySetSystemBackdrop(hwnd, DWMSBT_MAINWINDOW) Then Return

            ' Win10 fallback：先 Acrylic，再 Blur
            If Not TrySetAccent(hwnd, ACCENT_ENABLE_ACRYLICBLURBEHIND, &HCCFFFFFF) Then
                TrySetAccent(hwnd, ACCENT_ENABLE_BLURBEHIND, 0)
            End If
        End Sub

        Private Sub TryDisableSystemRounding(hwnd As IntPtr)
            Try
                Dim pref As Integer = DWMWCP_DONOTROUND
                DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, pref, Marshal.SizeOf(Of Integer)())
            Catch
            End Try
        End Sub

        Private Sub TrySetBorderColorTransparent(hwnd As IntPtr)
            Try
                ' DWMWA_BORDER_COLOR uses COLORREF (0x00BBGGRR). 0 means "default".
                ' We want fully transparent border — Win11 accepts 0x00000000 as "no visible border" in practice.
                Dim color As Integer = 0
                DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR, color, Marshal.SizeOf(Of Integer)())
            Catch
            End Try
        End Sub

        Private Function TrySetSystemBackdrop(hwnd As IntPtr, backdropType As Integer) As Boolean
            Try
                Dim hr = DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, backdropType, Marshal.SizeOf(Of Integer)())
                Return hr = 0
            Catch
                Return False
            End Try
        End Function

        Private Function TrySetAccent(hwnd As IntPtr, accentState As Integer, rgba As Integer) As Boolean
            Try
                Dim policy As New ACCENT_POLICY With {
                    .AccentState = accentState,
                    .AccentFlags = 2,
                    .GradientColor = rgba, ' AABBGGRR
                    .AnimationId = 0
                }

                Dim size = Marshal.SizeOf(Of ACCENT_POLICY)()
                Dim ptr = Marshal.AllocHGlobal(size)
                Marshal.StructureToPtr(policy, ptr, False)

                Dim data As New WINDOWCOMPOSITIONATTRIBDATA With {
                    .Attribute = WCA_ACCENT_POLICY,
                    .Data = ptr,
                    .SizeOfData = size
                }

                Dim ret = SetWindowCompositionAttribute(hwnd, data)
                Marshal.FreeHGlobal(ptr)

                Return ret <> 0
            Catch
                Return False
            End Try
        End Function

    End Module

End Namespace
