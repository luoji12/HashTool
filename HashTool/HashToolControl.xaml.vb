Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports System.Windows.Media
Imports Microsoft.Win32

Namespace HashTool

    Partial Public Class HashToolControl
        Inherits UserControl

        Private _dirtyCb As NativeMethods.HT_OnDirty ' 防 GC
        Private ReadOnly _uiTimer As DispatcherTimer

        Private _pendingRefresh As Boolean
        Private _wasRunning As Boolean = False
        Private _algoGuard As Boolean = False
        Private _inited As Boolean

        Private ReadOnly _maxThreads As Integer
        Private _lastOutText As String = ""

        Public Sub New()
            InitializeComponent()

            _maxThreads = Math.Max(1, Environment.ProcessorCount)
            TxtThreads.Text = _maxThreads.ToString()
            TxtStatus.Text = $"准备就绪（最大线程数：{_maxThreads}）"

            InitNative()

            _uiTimer = New DispatcherTimer(DispatcherPriority.Background)
            _uiTimer.Interval = TimeSpan.FromMilliseconds(80)
            AddHandler _uiTimer.Tick, AddressOf UiTimer_Tick
            _uiTimer.Start()

            AddHandler Me.Unloaded, Sub() ShutdownNative()
            AddHandler TxtThreads.LostFocus, Sub() NormalizeThreadText()
        End Sub

        ' ====== 给 MainWindow 拖拽转发用 ======
        Public Sub AddFiles(paths As IEnumerable(Of String))
            AddPaths(paths)
        End Sub

        Private Sub InitNative()
            If _inited Then Return

            Try
                _dirtyCb = AddressOf OnDirty
                Dim ok As Boolean = NativeMethods.HT_InitB(_dirtyCb, IntPtr.Zero)
                _inited = ok

                If Not ok Then
                    TxtStatus.Text = "初始化失败（HT_Init 返回 FALSE）"
                    Return
                End If

                ApplyThreadCount(True)
                _pendingRefresh = True
                RefreshUiOnce(forceText:=True)

            Catch ex As DllNotFoundException
                TxtStatus.Text = "找不到 HashTool.Core.dll（请确认部署 x64 DLL 并可被加载）"
                TxtOut.Text = ex.ToString()
            Catch ex As Exception
                TxtStatus.Text = "初始化异常"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

        Private Sub ShutdownNative()
            If Not _inited Then Return
            _inited = False
            Try
                NativeMethods.HT_Shutdown()
            Catch
            End Try
        End Sub

        Private Sub OnDirty(user As IntPtr)
            ' Core 通知有变化
            _pendingRefresh = True
        End Sub

        Private Sub UiTimer_Tick(sender As Object, e As EventArgs)
            If Not _inited Then Return

            ' ✅ 关键修复：不要再用 “没 dirty/没 poll 就 return” 的 gating
            ' 否则会复现 “上一个结果要等下一个文件进入后才显示”
            Dim hadDirty As Boolean = _pendingRefresh
            _pendingRefresh = False

            RefreshUiOnce(forceText:=hadDirty)
        End Sub

        ' ====== ScrollViewer helper：判断是否贴底 ======
        Private Shared Function FindDescendant(Of T As DependencyObject)(root As DependencyObject) As T
            If root Is Nothing Then Return Nothing
            Dim n As Integer = VisualTreeHelper.GetChildrenCount(root)
            For i As Integer = 0 To n - 1
                Dim c = VisualTreeHelper.GetChild(root, i)
                Dim hit = TryCast(c, T)
                If hit IsNot Nothing Then Return hit
                Dim deep = FindDescendant(Of T)(c)
                If deep IsNot Nothing Then Return deep
            Next
            Return Nothing
        End Function

        Private Function GetTextBoxScrollViewer(tb As TextBox) As ScrollViewer
            If tb Is Nothing Then Return Nothing
            tb.ApplyTemplate()
            Return FindDescendant(Of ScrollViewer)(tb)
        End Function

        Private Function IsTextBoxStickingToBottom(tb As TextBox) As Boolean
            Dim sv = GetTextBoxScrollViewer(tb)
            If sv Is Nothing Then Return True
            If sv.ScrollableHeight <= 0 Then Return True
            Return sv.VerticalOffset >= sv.ScrollableHeight - 1.0
        End Function

        ' ====== 核心刷新（修复：完成后速度不归零 + 输出延迟） ======
        Private Sub RefreshUiOnce(Optional forceText As Boolean = False)
            If Not _inited Then Return

            Try
                Dim s As NativeMethods.HT_Summary
                NativeMethods.HT_GetSummary(s)

                ' ---- 进度 & 速度（完成/空闲时强制显示 0）----
                Dim isRunningNow As Boolean = (s.runningCount > 0)
                Dim isDone As Boolean = (s.totalBytes > 0 AndAlso s.doneBytes >= s.totalBytes) OrElse (s.percent >= 100)

                Dim shownPercent As Integer = If(isDone, 100, Math.Max(0, Math.Min(100, s.percent)))
                Dim shownMbps As Double = If(isRunningNow AndAlso Not isDone, s.mbps, 0.0)

                Prog.Value = shownPercent

                If isDone AndAlso Not isRunningNow Then
                    TxtStatus.Text = $"完成 · {shownMbps:F1} MB/s · 0 running · {s.poolThreads} threads"
                ElseIf isDone AndAlso isRunningNow Then
                    ' 有些 Core 会在 100% 后做收尾
                    TxtStatus.Text = $"收尾中 · {shownMbps:F1} MB/s · {s.runningCount} running · {s.poolThreads} threads"
                Else
                    TxtStatus.Text = $"{shownPercent}% · {shownMbps:F1} MB/s · {s.runningCount} running · {s.poolThreads} threads"
                End If

                Dim justBecameIdle As Boolean = (_wasRunning AndAlso Not isRunningNow)
                _wasRunning = isRunningNow

                If justBecameIdle Then
                    forceText = True
                End If

                ' ---- 文本输出：每次 Tick 都允许读取（靠对比 _lastOutText 去抖）----
                Dim len As Integer = NativeMethods.HT_GetTextLength()
                If len <= 0 Then
                    If _lastOutText <> "" Then
                        Dim stick0 As Boolean = IsTextBoxStickingToBottom(TxtOut)
                        TxtOut.Clear()
                        _lastOutText = ""
                        If stick0 Then TxtOut.ScrollToEnd()
                    End If
                    Return
                End If

                Dim sb As New StringBuilder(len + 1)
                NativeMethods.HT_GetText(sb, sb.Capacity)
                Dim newText As String = sb.ToString()

                If newText = _lastOutText AndAlso Not forceText Then
                    Return
                End If

                Dim stickToBottom As Boolean = IsTextBoxStickingToBottom(TxtOut)

                TxtOut.Text = newText
                _lastOutText = newText

                If stickToBottom Then
                    TxtOut.ScrollToEnd()
                End If

            Catch ex As Exception
                TxtStatus.Text = "刷新异常"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

        ' ====== 线程数 ======
        Private Function ClampThreads(n As Integer) As Integer
            If n < 1 Then n = 1
            If n > _maxThreads Then n = _maxThreads
            Return n
        End Function

        Private Sub NormalizeThreadText()
            Dim n As Integer
            If Not Integer.TryParse(If(TxtThreads.Text, "").Trim(), n) Then n = _maxThreads
            n = ClampThreads(n)
            TxtThreads.Text = n.ToString()
        End Sub

        Private Sub ApplyThreadCount(clamp As Boolean)
            If Not _inited Then Return

            Dim n As Integer
            If Not Integer.TryParse(If(TxtThreads.Text, "").Trim(), n) Then n = _maxThreads
            If clamp Then n = ClampThreads(n)

            Try
                NativeMethods.HT_SetThreadCount(n)
                TxtThreads.Text = n.ToString()
            Catch ex As Exception
                TxtStatus.Text = "设置线程失败"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

        ' ✅ Spinner 上/下
        Private Sub BtnThrUp_Click(sender As Object, e As RoutedEventArgs)
            Dim n As Integer
            If Not Integer.TryParse(If(TxtThreads.Text, "").Trim(), n) Then n = _maxThreads
            n = ClampThreads(n + 1)
            TxtThreads.Text = n.ToString()
            ApplyThreadCount(True)
            _pendingRefresh = True
        End Sub

        Private Sub BtnThrDown_Click(sender As Object, e As RoutedEventArgs)
            Dim n As Integer
            If Not Integer.TryParse(If(TxtThreads.Text, "").Trim(), n) Then n = _maxThreads
            n = ClampThreads(n - 1)
            TxtThreads.Text = n.ToString()
            ApplyThreadCount(True)
            _pendingRefresh = True
        End Sub

        Private Sub BtnApplyThr_Click(sender As Object, e As RoutedEventArgs)
            NormalizeThreadText()
            ApplyThreadCount(True)
            _pendingRefresh = True
            RefreshUiOnce(forceText:=True)
        End Sub

        
        ' ====== 算法选择：MD5 与 SHA256 至少选中其一（可都选） ======
        Private Sub Algo_Checked(sender As Object, e As RoutedEventArgs)
            ' Checked 时无需处理；仅在取消勾选导致全不选时回滚
        End Sub

        Private Sub Algo_Unchecked(sender As Object, e As RoutedEventArgs)
            If _algoGuard Then Return

            Dim md5 As Boolean = (ChkMd5.IsChecked = True)
            Dim sha As Boolean = (ChkSha.IsChecked = True)

            If (Not md5) AndAlso (Not sha) Then
                _algoGuard = True
                Try
                    Dim cb = TryCast(sender, CheckBox)
                    If cb IsNot Nothing Then cb.IsChecked = True
                Finally
                    _algoGuard = False
                End Try
            End If
        End Sub

' ====== 添加文件（按钮/拖拽共用） ======
        Private Sub AddPaths(paths As IEnumerable(Of String))
            If Not _inited Then Return

            Dim md5 As Boolean = (ChkMd5.IsChecked = True)
            Dim sha As Boolean = (ChkSha.IsChecked = True)

            Dim added As Integer = 0
            Dim failed As Integer = 0

            For Each p In paths
                Try
                    If NativeMethods.HT_AddFileB(p, md5, sha) Then
                        added += 1
                    Else
                        failed += 1
                    End If
                Catch
                    failed += 1
                End Try
            Next

            TxtStatus.Text = $"已添加 {added} 个文件" & If(failed > 0, $"，失败 {failed}", "")
            _pendingRefresh = True
            RefreshUiOnce(forceText:=True)
        End Sub

        Private Sub BtnOpen_Click(sender As Object, e As RoutedEventArgs)
            If Not _inited Then Return

            Dim dlg As New OpenFileDialog With {
                .Title = "选择文件",
                .Multiselect = True,
                .CheckFileExists = True
            }
            If dlg.ShowDialog() <> True Then Return

            AddPaths(dlg.FileNames)
        End Sub

        Private Sub BtnCopy_Click(sender As Object, e As RoutedEventArgs)
            Try
                Clipboard.SetText(If(TxtOut.Text, ""))
                TxtStatus.Text = "已复制到剪贴板"
            Catch ex As Exception
                TxtStatus.Text = "复制失败"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

        Private Sub BtnExport_Click(sender As Object, e As RoutedEventArgs)
            Dim dlg As New SaveFileDialog With {
                .Title = "导出结果",
                .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*",
                .FileName = "hashes.txt",
                .AddExtension = True
            }
            If dlg.ShowDialog() <> True Then Return

            Try
                File.WriteAllText(dlg.FileName, If(TxtOut.Text, ""), New UTF8Encoding(False))
                TxtStatus.Text = $"已导出：{Path.GetFileName(dlg.FileName)}"
            Catch ex As Exception
                TxtStatus.Text = "导出失败"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

        Private Sub BtnClear_Click(sender As Object, e As RoutedEventArgs)
            If Not _inited Then
                Prog.Value = 0
                TxtOut.Clear()
                TxtStatus.Text = $"准备就绪（最大线程数：{_maxThreads}）"
                _lastOutText = ""
                _pendingRefresh = False
                _wasRunning = False
                Return
            End If

            Try
                Dim ok As Boolean = NativeMethods.HT_ClearAllB()
                If Not ok Then
                    TxtStatus.Text = "仍有任务在运行，无法清空"
                    Return
                End If

                Prog.Value = 0
                TxtOut.Clear()
                TxtStatus.Text = "已清空"

                _lastOutText = ""
                _wasRunning = False

                _pendingRefresh = True
                RefreshUiOnce(forceText:=True)

            Catch ex As Exception
                TxtStatus.Text = "清空异常"
                TxtOut.Text = ex.ToString()
            End Try
        End Sub

    End Class

End Namespace
