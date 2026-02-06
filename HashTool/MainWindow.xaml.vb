Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace HashTool

    Partial Public Class MainWindow
        Inherits Window

        ' =========================
        ' Win32: Region 真圆角
        ' =========================
        <DllImport("gdi32.dll")>
        Private Shared Function CreateRoundRectRgn(
            nLeftRect As Integer, nTopRect As Integer,
            nRightRect As Integer, nBottomRect As Integer,
            nWidthEllipse As Integer, nHeightEllipse As Integer) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetWindowRgn(hWnd As IntPtr, hRgn As IntPtr, bRedraw As Boolean) As Integer
        End Function

        <DllImport("gdi32.dll")>
        Private Shared Function DeleteObject(hObject As IntPtr) As Boolean
        End Function

        ' ✅ 改名：避免和 System.Windows.Rect 混淆
        <StructLayout(LayoutKind.Sequential)>
        Private Structure WRECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        <DllImport("user32.dll")>
        Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As WRECT) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetDpiForWindow(hWnd As IntPtr) As UInteger
        End Function

        Private Const WINDOW_RADIUS_DIP As Double = 10

        ' =========================
        ' Hook：WindowChrome 下保持 Region
        ' =========================
        Private _source As HwndSource = Nothing
        Private _pendingApply As Boolean = False

        Private Const WM_SIZE As Integer = &H5
        Private Const WM_WINDOWPOSCHANGED As Integer = &H47
        Private Const WM_DPICHANGED As Integer = &H2E0
        Private Const WM_SHOWWINDOW As Integer = &H18

        ' -------------------------
        ' Window lifecycle
        ' -------------------------
        Private Sub Window_SourceInitialized(sender As Object, e As EventArgs)
            Dim hwnd As IntPtr = New WindowInteropHelper(Me).Handle

            'DwmHelper.TryEnableBackdrop(hwnd)

            _source = TryCast(HwndSource.FromHwnd(hwnd), HwndSource)
            If _source IsNot Nothing Then
                _source.AddHook(New HwndSourceHook(AddressOf WndProc))
            End If

            UpdateMaxButtonGlyph()
            UpdateCornerForState()
            UpdateRootClip()
            ScheduleApplyRegion()
        End Sub

        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
            UpdateMaxButtonGlyph()
            UpdateCornerForState()
            UpdateRootClip()
            ScheduleApplyRegion()
        End Sub

        Private Sub Window_StateChanged(sender As Object, e As EventArgs)
            UpdateMaxButtonGlyph()
            UpdateCornerForState()
            UpdateRootClip()
            ScheduleApplyRegion()
        End Sub

        Private Sub RootCard_SizeChanged(sender As Object, e As SizeChangedEventArgs)
            UpdateRootClip()
            ScheduleApplyRegion()
        End Sub

        ' -------------------------
        ' WndProc hook
        ' -------------------------
        Private Function WndProc(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
            Select Case msg
                Case WM_SIZE, WM_WINDOWPOSCHANGED, WM_DPICHANGED, WM_SHOWWINDOW
                    ScheduleApplyRegion()
            End Select
            Return IntPtr.Zero
        End Function

        Private Sub ScheduleApplyRegion()
            ' Transparency mode: no Win32 region needed.
        End Sub

        ' ✅ 真圆角（用 GetWindowRect 获取真实像素尺寸）
        Private Sub ApplyWindowRoundRegion()
            ' No-op.
        End Sub

        ' -------------------------
        ' Title bar drag / maximize
        ' -------------------------
        Private Sub TitleBar_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            If e.ChangedButton <> MouseButton.Left Then Return

            If e.ClickCount = 2 Then
                ToggleMaximize()
                Return
            End If

            Try
                DragMove()
            Catch
            End Try
        End Sub

        Private Sub BtnMin_Click(sender As Object, e As RoutedEventArgs)
            WindowState = WindowState.Minimized
        End Sub

        Private Sub BtnMax_Click(sender As Object, e As RoutedEventArgs)
            ToggleMaximize()
        End Sub

        Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub

        Private Sub ToggleMaximize()
            WindowState = If(WindowState = WindowState.Maximized, WindowState.Normal, WindowState.Maximized)
        End Sub

        Private Sub UpdateMaxButtonGlyph()
            If BtnMaxGlyph Is Nothing Then Return
            BtnMaxGlyph.Text = If(WindowState = WindowState.Maximized, "❐", "☐")
        End Sub

        Private Sub UpdateCornerForState()
            If RootCard Is Nothing Then Return
            RootCard.CornerRadius = If(WindowState = WindowState.Maximized, New CornerRadius(0), New CornerRadius(10))
        End Sub

        ' ✅ 注意：这里用的是 WPF 的 System.Windows.Rect（避免和 WRECT 混）
        Private Sub UpdateRootClip()
            If RootCard Is Nothing Then Return

            Dim r As Double = If(WindowState = WindowState.Maximized, 0, 9)

            Dim w As Double = Math.Max(0, RootCard.ActualWidth)
            Dim h As Double = Math.Max(0, RootCard.ActualHeight)
            If w <= 0 OrElse h <= 0 Then Return

            Dim rect As New System.Windows.Rect(1.0, 1.0, Math.Max(0, w - 2.0), Math.Max(0, h - 2.0))
            RootCard.Clip = New RectangleGeometry(rect, r, r)
        End Sub

        ' -------------------------
        ' Drag & Drop
        ' -------------------------
        Private Sub Window_PreviewDragOver(sender As Object, e As DragEventArgs)
            If e.Data IsNot Nothing AndAlso e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effects = DragDropEffects.Copy
                DragOverlay.Visibility = Visibility.Visible
            Else
                e.Effects = DragDropEffects.None
                DragOverlay.Visibility = Visibility.Collapsed
            End If
            e.Handled = True
        End Sub

        Private Sub Window_PreviewDrop(sender As Object, e As DragEventArgs)
            DragOverlay.Visibility = Visibility.Collapsed

            If e.Data Is Nothing OrElse Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

            Dim arr = TryCast(e.Data.GetData(DataFormats.FileDrop), String())
            If arr Is Nothing OrElse arr.Length = 0 Then Return

            Dim files As New List(Of String)()

            For Each p In arr
                Try
                    If File.Exists(p) Then
                        files.Add(p)
                    ElseIf Directory.Exists(p) Then
                        files.AddRange(Directory.GetFiles(p, "*", SearchOption.AllDirectories))
                    End If
                Catch
                End Try
            Next

            If files.Count = 0 Then Return

            HashUI.AddFiles(files)
            e.Handled = True
        End Sub

    End Class

End Namespace
