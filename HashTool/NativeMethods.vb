Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO
Imports System.Reflection


Friend Module NativeMethods

    <UnmanagedFunctionPointer(CallingConvention.StdCall)>
    Friend Delegate Sub HT_OnDirty(user As IntPtr)

    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Friend Structure HT_Summary
        Public percent As Integer
        Public totalBytes As ULong
        Public doneBytes As ULong
        Public mbps As Double
        Public runningCount As Integer
        Public poolThreads As Integer
    End Structure

    Private Const DllName As String = "HashTool.Core.dll"

    ' ====== 单文件打包支持：从资源释放并预加载 HashTool.Core.dll ======
    Private _coreLoaded As Boolean = False
    Private ReadOnly _coreLock As New Object()

    <DllImport("kernel32", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Function LoadLibrary(lpFileName As String) As IntPtr
    End Function

    Private Sub EnsureCoreLoaded()
        If _coreLoaded Then Return

        SyncLock _coreLock
            If _coreLoaded Then Return

            Dim dllPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DllName)

            ' 1) 若与 EXE 同目录有 DLL，则直接加载
            If File.Exists(dllPath) Then
                If LoadLibrary(dllPath) = IntPtr.Zero Then
                    Throw New DllNotFoundException($"无法加载 {DllName}（{dllPath}），Win32Error={Marshal.GetLastWin32Error()}")
                End If
                _coreLoaded = True
                Return
            End If

            ' 2) 尝试从嵌入资源中释放 DLL
            Dim asm = Assembly.GetExecutingAssembly()
            Dim resName As String = Nothing
            For Each n In asm.GetManifestResourceNames()
                If n.EndsWith(DllName, StringComparison.OrdinalIgnoreCase) Then
                    resName = n
                    Exit For
                End If
            Next

            If resName IsNot Nothing Then
                Dim outDir As String = Path.Combine(Path.GetTempPath(), "HashTool")
                Directory.CreateDirectory(outDir)
                dllPath = Path.Combine(outDir, DllName)

                Using s = asm.GetManifestResourceStream(resName)
                    If s Is Nothing Then
                        Throw New DllNotFoundException($"无法读取嵌入资源 {resName}（{DllName}）")
                    End If
                    Using ms As New MemoryStream()
                        s.CopyTo(ms)
                        File.WriteAllBytes(dllPath, ms.ToArray())
                    End Using
                End Using

                If LoadLibrary(dllPath) = IntPtr.Zero Then
                    Throw New DllNotFoundException($"无法加载已释放的 {DllName}（{dllPath}），Win32Error={Marshal.GetLastWin32Error()}")
                End If

                _coreLoaded = True
                Return
            End If

            ' 3) 若未嵌入资源，则走系统默认探测
            If LoadLibrary(DllName) = IntPtr.Zero Then
                Throw New DllNotFoundException($"找不到 {DllName}（未在输出目录找到，也未作为嵌入资源打包），Win32Error={Marshal.GetLastWin32Error()}")
            End If

            _coreLoaded = True
        End SyncLock
    End Sub


    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Function HT_Init(cb As HT_OnDirty, user As IntPtr) As Integer
    End Function

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Sub HT_Shutdown()
    End Sub

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Sub HT_SetThreadCount(n As Integer)
    End Sub

    <DllImport(DllName, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)>
    Friend Function HT_AddFile(
        <MarshalAs(UnmanagedType.LPWStr)> path As String,
        md5 As Integer,
        sha256 As Integer
    ) As Integer
    End Function

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Sub HT_CancelAll()
    End Sub

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Function HT_ClearAll() As Integer
    End Function

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Sub HT_GetSummary(ByRef summary As HT_Summary)
    End Sub

    <DllImport(DllName, CallingConvention:=CallingConvention.StdCall)>
    Friend Function HT_GetTextLength() As Integer
    End Function

    <DllImport(DllName, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)>
    Friend Function HT_GetText(buf As StringBuilder, cch As Integer) As Integer
    End Function

    ' ====== VB 友好包装：避免 True=-1 ======

    Friend Function HT_InitB(cb As HT_OnDirty, user As IntPtr) As Boolean
        EnsureCoreLoaded()
        Return HT_Init(cb, user) <> 0
    End Function

    Friend Function HT_AddFileB(path As String, md5 As Boolean, sha256 As Boolean) As Boolean
        Dim bMd5 As Integer = If(md5, 1, 0)
        Dim bSha As Integer = If(sha256, 1, 0)
        Return HT_AddFile(path, bMd5, bSha) <> 0
    End Function

    Friend Function HT_ClearAllB() As Boolean
        EnsureCoreLoaded()
        Return HT_ClearAll() <> 0
    End Function

End Module
