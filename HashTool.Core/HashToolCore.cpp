
#include "HashToolCore.h"

#include <commctrl.h>
#include <shellapi.h>
#include <strsafe.h>
#include <uxtheme.h>
#include <stdint.h>
#include <stdarg.h>
#include <wchar.h>
#include <bcrypt.h>

#pragma comment(lib, "Bcrypt.lib")
#pragma comment(lib, "Version.lib")
#pragma comment(lib, "Kernel32.lib")

#ifndef _countof
#define _countof(a) (sizeof(a)/sizeof((a)[0]))
#endif

// ---------------- 原常量保留 ----------------
#define MAX_TASKS 250

static const DWORD UI_THROTTLE_MS_WORKER = 100;
static const DWORD UI_TEXT_PAINT_MIN_MS = 250;
static const DWORD UI_SUMMARY_MIN_MS = 100;

static const DWORD IO_BUF_SIZE = 8 * 1024 * 1024;
static const DWORD EDIT_LIMIT_TEXT = 10 * 1024 * 1024;

// ---------------- 结构体（保持原样） ----------------
typedef struct {
	WCHAR szFilePath[MAX_PATH];

	ULONGLONG ullFileSize;
	ULONGLONG ullFileSizeInit;
	FILETIME  ftModify;
	WCHAR     szFileVersion[64];

	WCHAR szMD5Value[33];
	WCHAR szSHA256Value[65];
	BOOL  bCalcMD5;
	BOOL  bCalcSHA256;

	volatile LONG bFinished;
	volatile LONG bCanceled;
	BOOL  bSuccessMD5;
	BOOL  bSuccessSHA256;

	volatile LONGLONG llDoneBytes;

	ULONGLONG ullStartTick;
	ULONGLONG ullEndTick;

	volatile LONG lLastUiPctNotified; // init = -1

	PTP_WORK work;
} FILE_HASH_TASK;

typedef struct {
	WCHAR szFilePath[MAX_PATH];
	ULONGLONG ullFileSize;
	ULONGLONG ullFileSizeInit;
	FILETIME  ftModify;
	WCHAR     szFileVersion[64];

	WCHAR szMD5Value[33];
	WCHAR szSHA256Value[65];
	BOOL  bCalcMD5;
	BOOL  bCalcSHA256;

	LONG bFinished;
	LONG bCanceled;
	BOOL bSuccessMD5;
	BOOL bSuccessSHA256;

	LONGLONG llDoneBytes;
	ULONGLONG ullStartTick;
	ULONGLONG ullEndTick;
} TASK_SNAPSHOT;

// ---------------- 全局（保留核心状态） ----------------
static FILE_HASH_TASK g_Tasks[MAX_TASKS];
static int  g_nTaskCount = 0;

static CRITICAL_SECTION g_csTasks;

static volatile LONG g_lRunningCount = 0;
static volatile LONG g_lCancelAll = 0;

static volatile LONGLONG g_llTotalBytesAll = 0;
static volatile LONGLONG g_llDoneBytesAll = 0;

static volatile ULONGLONG g_ullOverallStartTick = 0;

static volatile LONG g_bTextDirty = 0;

static DWORD g_dwLastTextBuildTick = 0;
static DWORD g_dwLastSummaryTick = 0;

// 动态文本缓冲（保留）
static WCHAR* g_pTextBuf = NULL;
static size_t g_cchTextCap = 0;

// Threadpool（保留）
static PTP_POOL g_pool = NULL;
static TP_CALLBACK_ENVIRON g_callEnv;
static LONG g_lPoolThreads = 0;
static PTP_CLEANUP_GROUP g_cleanup = NULL;

// CNG providers（保留）
static BCRYPT_ALG_HANDLE g_hAlgMD5 = NULL;
static BCRYPT_ALG_HANDLE g_hAlgSHA256 = NULL;
static DWORD g_dwObjLenMD5 = 0, g_dwObjLenSHA = 0;
static DWORD g_dwHashLenMD5 = 0, g_dwHashLenSHA = 0;

// UI dirty callback（新增）
static HT_OnDirty g_cbDirty = NULL;
static void* g_cbUser = NULL;

static ULONGLONG NowTick64(void) { return GetTickCount64(); }

// ---------------- 工具函数（原样） ----------------
static void BinToHexUpper(const BYTE* pBin, DWORD dwLen, WCHAR* pHex, size_t cchHex)
{
	if (!pBin || !pHex || cchHex < (size_t)dwLen * 2 + 1) return;
	static const WCHAR kHex[] = L"0123456789ABCDEF";
	for (DWORD i = 0; i < dwLen; i++) {
		pHex[i * 2] = kHex[(pBin[i] >> 4) & 0x0F];
		pHex[i * 2 + 1] = kHex[pBin[i] & 0x0F];
	}
	pHex[dwLen * 2] = L'\0';
}

static void GetFileVersionStr(const WCHAR* szFilePath, WCHAR* szVersion, size_t cchVersion)
{
	if (!szFilePath || !szVersion || cchVersion == 0) return;
	szVersion[0] = L'\0';

	DWORD dwHandle = 0;
	DWORD dwSize = GetFileVersionInfoSizeW(szFilePath, &dwHandle);
	if (dwSize == 0) return;

	BYTE* pBuf = (BYTE*)LocalAlloc(LPTR, dwSize);
	if (!pBuf) return;

	if (GetFileVersionInfoW(szFilePath, dwHandle, dwSize, pBuf)) {
		VS_FIXEDFILEINFO* pInfo = NULL;
		UINT uLen = 0;
		if (VerQueryValueW(pBuf, L"\\", (LPVOID*)&pInfo, &uLen) && pInfo) {
			StringCchPrintfW(szVersion, cchVersion, L"%u.%u.%u.%u",
				HIWORD(pInfo->dwFileVersionMS),
				LOWORD(pInfo->dwFileVersionMS),
				HIWORD(pInfo->dwFileVersionLS),
				LOWORD(pInfo->dwFileVersionLS));
		}
	}
	LocalFree(pBuf);
}

static void FileTimeToLocalStr(const FILETIME* pFt, WCHAR* szTimeStr, size_t cch)
{
	if (!pFt || !szTimeStr || cch == 0) return;
	szTimeStr[0] = L'\0';

	FILETIME ftLocal = { 0 };
	SYSTEMTIME st = { 0 };
	if (!FileTimeToLocalFileTime(pFt, &ftLocal)) return;
	if (!FileTimeToSystemTime(&ftLocal, &st)) return;

	StringCchPrintfW(szTimeStr, cch, L"%04u-%02u-%02u %02u:%02u:%02u",
		st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

static void FormatBytes(ULONGLONG b, WCHAR* out, size_t cch)
{
	if (!out || cch == 0) return;
	const double kb = 1024.0, mb = kb * 1024.0, gb = mb * 1024.0;
	if (b >= (ULONGLONG)gb) StringCchPrintfW(out, cch, L"%.2f GB", (double)b / gb);
	else if (b >= (ULONGLONG)mb) StringCchPrintfW(out, cch, L"%.2f MB", (double)b / mb);
	else if (b >= (ULONGLONG)kb) StringCchPrintfW(out, cch, L"%.2f KB", (double)b / kb);
	else StringCchPrintfW(out, cch, L"%I64u B", (unsigned long long)b);
}

static void FormatSpeedMBps(double mbps, WCHAR* out, size_t cch)
{
	if (!out || cch == 0) return;
	if (mbps < 0) mbps = 0;
	StringCchPrintfW(out, cch, L"%.2f MB/s", mbps);
}

static void FormatSeconds(double sec, WCHAR* out, size_t cch)
{
	if (!out || cch == 0) return;
	if (sec < 0) sec = 0;
	StringCchPrintfW(out, cch, L"%.2f s", sec);
}

static void EnsureOverallStart(void)
{
	if (g_ullOverallStartTick == 0) g_ullOverallStartTick = NowTick64();
}

// ---------------- “通知 UI”替代（不改动逻辑，只换出口） ----------------
static void SignalUiDirty(void)
{
	if (g_cbDirty) g_cbDirty(g_cbUser);
}

// 原来的 RequestUiUpdate/MarkTextDirtyAndRequest
static void RequestUiUpdate(void) { SignalUiDirty(); }

static void MarkTextDirtyAndRequest(void)
{
	InterlockedExchange(&g_bTextDirty, 1);
	RequestUiUpdate();
}

// ---------------- 动态文本缓冲（原样） ----------------
static BOOL EnsureTextCapacity(size_t cchNeed)
{
	if (cchNeed <= g_cchTextCap && g_pTextBuf) return TRUE;

	size_t newCap = (g_cchTextCap ? g_cchTextCap : 65536);
	while (newCap < cchNeed) {
		if (newCap > (size_t)16 * 1024 * 1024) { // 16M chars
			newCap = cchNeed;
			break;
		}
		newCap *= 2;
	}

	WCHAR* pNew = NULL;
	if (g_pTextBuf) {
		pNew = (WCHAR*)HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, g_pTextBuf, newCap * sizeof(WCHAR));
	}
	else {
		pNew = (WCHAR*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, newCap * sizeof(WCHAR));
	}
	if (!pNew) return FALSE;

	g_pTextBuf = pNew;
	g_cchTextCap = newCap;
	return TRUE;
}

static void TextReset(void)
{
	if (g_pTextBuf && g_cchTextCap) g_pTextBuf[0] = L'\0';
}

static void AppendLineDyn(WCHAR** pDst, size_t* pcchRemain, const WCHAR* fmt, ...)
{
	if (!pDst || !pcchRemain || !fmt) return;

	for (;;) {
		if (*pDst == NULL || *pcchRemain == 0) return;

		WCHAR tmp[2048];
		va_list ap; va_start(ap, fmt);
		HRESULT hr = StringCchVPrintfW(tmp, _countof(tmp), fmt, ap);
		va_end(ap);
		if (FAILED(hr)) return;

		size_t need = wcslen(tmp);
		if (need + 1 <= *pcchRemain) {
			StringCchCatExW(*pDst, *pcchRemain, tmp, pDst, pcchRemain, 0);
			return;
		}

		size_t used = (size_t)(*pDst - g_pTextBuf);
		size_t totalNeed = used + need + 1 + 4096;
		if (!EnsureTextCapacity(totalNeed)) return;

		*pDst = g_pTextBuf + used;
		*pcchRemain = g_cchTextCap - used;
	}
}

// ---------------- CNG 初始化（原样） ----------------
static BOOL InitCngProviders(void)
{
	NTSTATUS st;
	DWORD cb = 0;

	if (!g_hAlgMD5) {
		st = BCryptOpenAlgorithmProvider(&g_hAlgMD5, BCRYPT_MD5_ALGORITHM, NULL, 0);
		if (st != 0) return FALSE;

		st = BCryptGetProperty(g_hAlgMD5, BCRYPT_OBJECT_LENGTH, (PUCHAR)&g_dwObjLenMD5, sizeof(g_dwObjLenMD5), &cb, 0);
		if (st != 0) return FALSE;
		st = BCryptGetProperty(g_hAlgMD5, BCRYPT_HASH_LENGTH, (PUCHAR)&g_dwHashLenMD5, sizeof(g_dwHashLenMD5), &cb, 0);
		if (st != 0) return FALSE;
	}

	if (!g_hAlgSHA256) {
		st = BCryptOpenAlgorithmProvider(&g_hAlgSHA256, BCRYPT_SHA256_ALGORITHM, NULL, 0);
		if (st != 0) return FALSE;

		st = BCryptGetProperty(g_hAlgSHA256, BCRYPT_OBJECT_LENGTH, (PUCHAR)&g_dwObjLenSHA, sizeof(g_dwObjLenSHA), &cb, 0);
		if (st != 0) return FALSE;
		st = BCryptGetProperty(g_hAlgSHA256, BCRYPT_HASH_LENGTH, (PUCHAR)&g_dwHashLenSHA, sizeof(g_dwHashLenSHA), &cb, 0);
		if (st != 0) return FALSE;
	}
	return TRUE;
}

static void CleanupCngProviders(void)
{
	if (g_hAlgMD5) { BCryptCloseAlgorithmProvider(g_hAlgMD5, 0); g_hAlgMD5 = NULL; }
	if (g_hAlgSHA256) { BCryptCloseAlgorithmProvider(g_hAlgSHA256, 0); g_hAlgSHA256 = NULL; }
}

// ---------------- Hash 计算（原样：含 done 补齐） ----------------
static BOOL CalculateHashes_WithProgress(FILE_HASH_TASK* t)
{
	if (!t) return FALSE;

	BOOL ok = FALSE;
	HANDLE hFile = INVALID_HANDLE_VALUE;
	BYTE* buf = NULL;

	ULONGLONG done = 0;
	DWORD lastUi = 0;

	NTSTATUS st = 0;
	BCRYPT_HASH_HANDLE hMd5 = NULL, hSha = NULL;

	PUCHAR objMd5 = NULL, objSha = NULL;
	PUCHAR outMd5 = NULL, outSha = NULL;

	LARGE_INTEGER sz; sz.QuadPart = 0;

	hFile = CreateFileW(
		t->szFilePath,
		GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
		NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
		NULL
	);
	if (hFile == INVALID_HANDLE_VALUE) goto cleanup;

	if (GetFileSizeEx(hFile, &sz)) {
		t->ullFileSize = (ULONGLONG)sz.QuadPart;
	}

	if (t->bCalcMD5) {
		objMd5 = (PUCHAR)HeapAlloc(GetProcessHeap(), 0, g_dwObjLenMD5);
		outMd5 = (PUCHAR)HeapAlloc(GetProcessHeap(), 0, g_dwHashLenMD5);
		if (!objMd5 || !outMd5) goto cleanup;

		st = BCryptCreateHash(g_hAlgMD5, &hMd5, objMd5, g_dwObjLenMD5, NULL, 0, 0);
		if (st != 0) goto cleanup;
	}

	if (t->bCalcSHA256) {
		objSha = (PUCHAR)HeapAlloc(GetProcessHeap(), 0, g_dwObjLenSHA);
		outSha = (PUCHAR)HeapAlloc(GetProcessHeap(), 0, g_dwHashLenSHA);
		if (!objSha || !outSha) goto cleanup;

		st = BCryptCreateHash(g_hAlgSHA256, &hSha, objSha, g_dwObjLenSHA, NULL, 0, 0);
		if (st != 0) goto cleanup;
	}

	buf = (BYTE*)VirtualAlloc(NULL, IO_BUF_SIZE, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
	if (!buf) goto cleanup;

	done = 0;
	lastUi = GetTickCount();
	InterlockedExchange64(&t->llDoneBytes, 0);

	ok = TRUE;

	for (;;) {
		if (InterlockedCompareExchange(&g_lCancelAll, 0, 0) != 0) {
			InterlockedExchange(&t->bCanceled, 1);
			ok = FALSE;
			break;
		}

		DWORD dwRead = 0;
		BOOL br = ReadFile(hFile, buf, IO_BUF_SIZE, &dwRead, NULL);
		if (!br) { ok = FALSE; break; }
		if (dwRead == 0) break;

		if (hMd5) {
			st = BCryptHashData(hMd5, (PUCHAR)buf, dwRead, 0);
			if (st != 0) { ok = FALSE; break; }
		}
		if (hSha) {
			st = BCryptHashData(hSha, (PUCHAR)buf, dwRead, 0);
			if (st != 0) { ok = FALSE; break; }
		}

		done += dwRead;
		InterlockedExchange64(&t->llDoneBytes, (LONGLONG)done);
		InterlockedAdd64(&g_llDoneBytesAll, (LONGLONG)dwRead);

		DWORD now = GetTickCount();
		if (now - lastUi >= UI_THROTTLE_MS_WORKER) {
			lastUi = now;

			LONG pct = 0;
			ULONGLONG fs = t->ullFileSize;
			if (fs > 0) {
				double dp = (double)done * 100.0 / (double)fs;
				if (dp < 0) dp = 0;
				if (dp > 100) dp = 100;
				pct = (LONG)(dp + 0.5);
			}

			LONG lastPct = InterlockedCompareExchange(&t->lLastUiPctNotified, 0, 0);
			if (pct >= lastPct + 1 || pct == 100 || lastPct < 0) {
				InterlockedExchange(&t->lLastUiPctNotified, pct);
				InterlockedExchange(&g_bTextDirty, 1);
			}
			RequestUiUpdate();
		}
	}

	// ✅ 对账补齐（原样）
	if (InterlockedCompareExchange(&t->bCanceled, 0, 0) == 0) {
		ULONGLONG fileSize = t->ullFileSize;
		if (done > fileSize) done = fileSize;

		if (done < fileSize) {
			LONGLONG remain = (LONGLONG)(fileSize - done);
			InterlockedAdd64(&g_llDoneBytesAll, remain);
			done = fileSize;
			InterlockedExchange64(&t->llDoneBytes, (LONGLONG)done);
		}
	}

	if (ok) {
		if (hMd5) {
			st = BCryptFinishHash(hMd5, outMd5, g_dwHashLenMD5, 0);
			if (st == 0) {
				BinToHexUpper(outMd5, g_dwHashLenMD5, t->szMD5Value, _countof(t->szMD5Value));
				t->bSuccessMD5 = TRUE;
			}
			else {
				t->bSuccessMD5 = FALSE;
			}
		}
		if (hSha) {
			st = BCryptFinishHash(hSha, outSha, g_dwHashLenSHA, 0);
			if (st == 0) {
				BinToHexUpper(outSha, g_dwHashLenSHA, t->szSHA256Value, _countof(t->szSHA256Value));
				t->bSuccessSHA256 = TRUE;
			}
			else {
				t->bSuccessSHA256 = FALSE;
			}
		}
	}

cleanup:
	if (buf) VirtualFree(buf, 0, MEM_RELEASE);

	if (hMd5) BCryptDestroyHash(hMd5);
	if (hSha) BCryptDestroyHash(hSha);

	if (objMd5) HeapFree(GetProcessHeap(), 0, objMd5);
	if (objSha) HeapFree(GetProcessHeap(), 0, objSha);
	if (outMd5) HeapFree(GetProcessHeap(), 0, outMd5);
	if (outSha) HeapFree(GetProcessHeap(), 0, outSha);

	if (hFile != INVALID_HANDLE_VALUE) CloseHandle(hFile);

	InterlockedExchange(&g_bTextDirty, 1);
	RequestUiUpdate();
	return ok;
}

// ---------------- 线程池管理（原样） ----------------
static void EnsureThreadPool(void)
{
	if (g_pool) return;

	g_pool = CreateThreadpool(NULL);
	if (!g_pool) return;

	g_cleanup = CreateThreadpoolCleanupGroup();
	if (!g_cleanup) {
		CloseThreadpool(g_pool);
		g_pool = NULL;
		return;
	}

	InitializeThreadpoolEnvironment(&g_callEnv);
	SetThreadpoolCallbackPool(&g_callEnv, g_pool);
	SetThreadpoolCallbackCleanupGroup(&g_callEnv, g_cleanup, NULL);

	SYSTEM_INFO si; GetSystemInfo(&si);
	LONG n = (LONG)(si.dwNumberOfProcessors ? si.dwNumberOfProcessors : 1);
	if (n < 1) n = 1;
	g_lPoolThreads = n;

	SetThreadpoolThreadMinimum(g_pool, (DWORD)n);
	SetThreadpoolThreadMaximum(g_pool, (DWORD)n);
}

static void ApplyThreadPoolSize(LONG n)
{
	if (!g_pool) return;
	if (n < 1) n = 1;
	if (n > 64) n = 64;

	SetThreadpoolThreadMaximum(g_pool, (DWORD)n);
	SetThreadpoolThreadMinimum(g_pool, (DWORD)n);
	g_lPoolThreads = n;
}

// ---------------- WorkCallback（原样：total delta 修正唯一点） ----------------
static VOID CALLBACK WorkCallback(PTP_CALLBACK_INSTANCE Instance, PVOID Context, PTP_WORK Work)
{
	(void)Instance; (void)Work;
	FILE_HASH_TASK* t = (FILE_HASH_TASK*)Context;
	if (!t) return;

	t->ullStartTick = NowTick64();

	WIN32_FILE_ATTRIBUTE_DATA fad = { 0 };
	if (GetFileAttributesExW(t->szFilePath, GetFileExInfoStandard, &fad)) {
		t->ftModify = fad.ftLastWriteTime;
		GetFileVersionStr(t->szFilePath, t->szFileVersion, _countof(t->szFileVersion));
	}

	ULONGLONG realSize = t->ullFileSizeInit;
	HANDLE hf = CreateFileW(
		t->szFilePath, GENERIC_READ,
		FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
		NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
		NULL
	);
	if (hf != INVALID_HANDLE_VALUE) {
		LARGE_INTEGER li;
		if (GetFileSizeEx(hf, &li)) realSize = (ULONGLONG)li.QuadPart;
		CloseHandle(hf);
	}

	// ✅ total delta 修正：只在这里做（原样）
	if (realSize != t->ullFileSizeInit) {
		LONGLONG delta = (LONGLONG)realSize - (LONGLONG)t->ullFileSizeInit;
		InterlockedAdd64(&g_llTotalBytesAll, delta);
		t->ullFileSizeInit = realSize;
	}
	t->ullFileSize = realSize;

	(void)CalculateHashes_WithProgress(t);

	t->ullEndTick = NowTick64();
	InterlockedExchange(&t->bFinished, 1);

	InterlockedDecrement(&g_lRunningCount);
	MarkTextDirtyAndRequest();
}

// ---------------- 添加文件（原样：total 初值唯一点） ----------------
static BOOL TaskExists_Locked(const WCHAR* path)
{
	for (int i = 0; i < g_nTaskCount; i++) {
		if (_wcsicmp(g_Tasks[i].szFilePath, path) == 0) return TRUE;
	}
	return FALSE;
}

static void AddOneFile_Locked(const WCHAR* path, BOOL bMD5, BOOL bSHA)
{
	if (g_nTaskCount >= MAX_TASKS) return;
	if (TaskExists_Locked(path)) return;

	FILE_HASH_TASK* t = &g_Tasks[g_nTaskCount];
	ZeroMemory(t, sizeof(*t));

	StringCchCopyW(t->szFilePath, _countof(t->szFilePath), path);
	t->bCalcMD5 = bMD5;
	t->bCalcSHA256 = bSHA;

	InterlockedExchange(&t->lLastUiPctNotified, -1);

	ULONGLONG initSize = 0;
	WIN32_FILE_ATTRIBUTE_DATA fad = { 0 };
	if (GetFileAttributesExW(path, GetFileExInfoStandard, &fad)) {
		ULARGE_INTEGER u; u.LowPart = fad.nFileSizeLow; u.HighPart = fad.nFileSizeHigh;
		initSize = u.QuadPart;
		t->ftModify = fad.ftLastWriteTime;
	}
	t->ullFileSizeInit = initSize;
	t->ullFileSize = initSize;

	// ✅ total += initSize（第一处，原样）
	InterlockedAdd64(&g_llTotalBytesAll, (LONGLONG)initSize);

	t->work = CreateThreadpoolWork(WorkCallback, t, &g_callEnv);
	if (t->work) {
		InterlockedIncrement(&g_lRunningCount);
		SubmitThreadpoolWork(t->work);
		g_nTaskCount++;

		InterlockedExchange(&g_bTextDirty, 1);
	}
	else {
		InterlockedAdd64(&g_llTotalBytesAll, -(LONGLONG)initSize);
	}
}

// ---------------- 文本构建（由 UI 调用；逻辑与原拼接一致） ----------------
static void BuildTextIfDirty_Throttle(BOOL force)
{
	DWORD now = GetTickCount();
	if (!force) {
		if (now - g_dwLastTextBuildTick < UI_TEXT_PAINT_MIN_MS) return;
	}
	g_dwLastTextBuildTick = now;

	if (InterlockedExchange(&g_bTextDirty, 0) == 0 && !force) return;

	TASK_SNAPSHOT snap[MAX_TASKS];
	int nSnap = 0;

	EnterCriticalSection(&g_csTasks);
	nSnap = g_nTaskCount;
	if (nSnap > MAX_TASKS) nSnap = MAX_TASKS;

	for (int i = 0; i < nSnap; i++) {
		FILE_HASH_TASK* t = &g_Tasks[i];
		TASK_SNAPSHOT* s = &snap[i];

		ZeroMemory(s, sizeof(*s));
		StringCchCopyW(s->szFilePath, _countof(s->szFilePath), t->szFilePath);
		s->ullFileSize = t->ullFileSize;
		s->ullFileSizeInit = t->ullFileSizeInit;
		s->ftModify = t->ftModify;
		StringCchCopyW(s->szFileVersion, _countof(s->szFileVersion), t->szFileVersion);
		StringCchCopyW(s->szMD5Value, _countof(s->szMD5Value), t->szMD5Value);
		StringCchCopyW(s->szSHA256Value, _countof(s->szSHA256Value), t->szSHA256Value);
		s->bCalcMD5 = t->bCalcMD5;
		s->bCalcSHA256 = t->bCalcSHA256;
		s->bFinished = InterlockedCompareExchange(&t->bFinished, 0, 0);
		s->bCanceled = InterlockedCompareExchange(&t->bCanceled, 0, 0);
		s->bSuccessMD5 = t->bSuccessMD5;
		s->bSuccessSHA256 = t->bSuccessSHA256;
		s->llDoneBytes = InterlockedCompareExchange64(&t->llDoneBytes, 0, 0);
		s->ullStartTick = t->ullStartTick;
		s->ullEndTick = t->ullEndTick;
	}
	LeaveCriticalSection(&g_csTasks);

	size_t roughNeed = (size_t)(nSnap ? nSnap : 1) * 1024 + 4096;
	if (!EnsureTextCapacity(roughNeed)) return;

	TextReset();
	WCHAR* pDst = g_pTextBuf;
	size_t cchRemain = g_cchTextCap;

	ULONGLONG nowTick = NowTick64();
	ULONGLONG overallStart = g_ullOverallStartTick;

	for (int i = 0; i < nSnap; i++) {
		TASK_SNAPSHOT* t = &snap[i];

		if (i > 0) AppendLineDyn(&pDst, &cchRemain, L"──────────────────────────────────────\r\n");

		WCHAR timeStr[64] = { 0 };
		FileTimeToLocalStr(&t->ftModify, timeStr, _countof(timeStr));

		WCHAR sizeStr[64] = { 0 };
		FormatBytes(t->ullFileSize, sizeStr, _countof(sizeStr));

		LONGLONG doneBytes = t->llDoneBytes;
		if (doneBytes < 0) doneBytes = 0;

		WCHAR doneStr2[64] = { 0 };
		FormatBytes((ULONGLONG)doneBytes, doneStr2, _countof(doneStr2));

		int pct = 0;
		if (t->ullFileSize > 0) {
			pct = (int)((double)doneBytes * 100.0 / (double)t->ullFileSize + 0.5);
			if (pct > 100) pct = 100;
			if (pct < 0) pct = 0;
		}

		BOOL finished = (t->bFinished != 0);
		BOOL canceled = (t->bCanceled != 0);

		ULONGLONG start = t->ullStartTick ? t->ullStartTick : overallStart;
		ULONGLONG end = finished ? t->ullEndTick : nowTick;
		double fileSec = 0.0;
		if (start != 0 && end >= start) fileSec = (double)(end - start) / 1000.0;
		if (fileSec < 0.001) fileSec = 0.001;

		double fileMBps = ((double)doneBytes / (1024.0 * 1024.0)) / fileSec;
		WCHAR fileSpeedStr[64], fileTimeStr[64];
		FormatSpeedMBps(fileMBps, fileSpeedStr, _countof(fileSpeedStr));
		FormatSeconds(fileSec, fileTimeStr, _countof(fileTimeStr));

		AppendLineDyn(&pDst, &cchRemain, L"文件: %s\r\n", t->szFilePath);
		AppendLineDyn(&pDst, &cchRemain, L"大小: %s\r\n", sizeStr);
		AppendLineDyn(&pDst, &cchRemain, L"修改时间: %s\r\n", timeStr[0] ? timeStr : L"(未知)");
		if (t->szFileVersion[0]) AppendLineDyn(&pDst, &cchRemain, L"文件版本: %s\r\n", t->szFileVersion);

		if (t->bCalcMD5) {
			if (!finished) AppendLineDyn(&pDst, &cchRemain, L"MD5: 正在计算...\r\n");
			else AppendLineDyn(&pDst, &cchRemain, L"MD5: %s\r\n",
				(t->bSuccessMD5 ? t->szMD5Value : (canceled ? L"(取消)" : L"(失败)")));
		}
		if (t->bCalcSHA256) {
			if (!finished) AppendLineDyn(&pDst, &cchRemain, L"SHA256: 正在计算...\r\n");
			else AppendLineDyn(&pDst, &cchRemain, L"SHA256: %s\r\n",
				(t->bSuccessSHA256 ? t->szSHA256Value : (canceled ? L"(取消)" : L"(失败)")));
		}

		AppendLineDyn(&pDst, &cchRemain, L"\r\n");
	}
}

// ---------------- DLL 导出 API ----------------
BOOL __stdcall HT_Init(HT_OnDirty cb, void* user)
{
	g_cbDirty = cb;
	g_cbUser = user;

	InitializeCriticalSection(&g_csTasks);

	if (!InitCngProviders()) return FALSE;
	EnsureTextCapacity(131072);
	EnsureThreadPool();

	// 初始化文本
	InterlockedExchange(&g_bTextDirty, 1);
	BuildTextIfDirty_Throttle(TRUE);
	return TRUE;
}

void __stdcall HT_Shutdown()
{
	// 取消
	InterlockedExchange(&g_lCancelAll, 1);

	// 关闭线程池：cleanup group 统一取消/等待（原策略）
	if (g_cleanup) {
		CloseThreadpoolCleanupGroupMembers(g_cleanup, TRUE, NULL);
		CloseThreadpoolCleanupGroup(g_cleanup);
		g_cleanup = NULL;
	}
	if (g_pool) {
		CloseThreadpool(g_pool);
		g_pool = NULL;
	}
	DestroyThreadpoolEnvironment(&g_callEnv);

	CleanupCngProviders();

	if (g_pTextBuf) {
		HeapFree(GetProcessHeap(), 0, g_pTextBuf);
		g_pTextBuf = NULL;
		g_cchTextCap = 0;
	}

	DeleteCriticalSection(&g_csTasks);

	g_cbDirty = NULL;
	g_cbUser = NULL;
}

void __stdcall HT_SetThreadCount(int n)
{
	EnsureThreadPool();
	ApplyThreadPoolSize((LONG)n);
	RequestUiUpdate();
}

BOOL __stdcall HT_AddFile(const wchar_t* path, BOOL md5, BOOL sha256)
{
	if (!path || !path[0]) return FALSE;
	if (!md5 && !sha256) return FALSE;

	EnsureThreadPool();
	if (!g_pool) return FALSE;

	EnsureOverallStart();
	InterlockedExchange(&g_lCancelAll, 0);

	EnterCriticalSection(&g_csTasks);

	// 新批次：清零计数（保持你原 StartFiles 行为）
	if (g_nTaskCount == 0) {
		InterlockedExchange64(&g_llDoneBytesAll, 0);
		InterlockedExchange64(&g_llTotalBytesAll, 0);
		g_ullOverallStartTick = NowTick64();
	}
	else {
		// 追加：若上一轮已完成，将 done 对齐 total
		if (InterlockedCompareExchange(&g_lRunningCount, 0, 0) == 0) {
			LONGLONG total = InterlockedCompareExchange64(&g_llTotalBytesAll, 0, 0);
			InterlockedExchange64(&g_llDoneBytesAll, total);
		}
	}

	AddOneFile_Locked(path, md5, sha256);

	LeaveCriticalSection(&g_csTasks);

	MarkTextDirtyAndRequest();
	return TRUE;
}

void __stdcall HT_CancelAll()
{
	InterlockedExchange(&g_lCancelAll, 1);
	MarkTextDirtyAndRequest();
}

BOOL __stdcall HT_ClearAll()
{
	if (InterlockedCompareExchange(&g_lRunningCount, 0, 0) != 0) return FALSE;

	EnterCriticalSection(&g_csTasks);
	ZeroMemory(g_Tasks, sizeof(g_Tasks));
	g_nTaskCount = 0;
	LeaveCriticalSection(&g_csTasks);

	InterlockedExchange64(&g_llTotalBytesAll, 0);
	InterlockedExchange64(&g_llDoneBytesAll, 0);
	InterlockedExchange(&g_lCancelAll, 0);
	g_ullOverallStartTick = 0;

	if (g_pTextBuf) g_pTextBuf[0] = L'\0';
	InterlockedExchange(&g_bTextDirty, 0);

	RequestUiUpdate();
	return TRUE;
}

void __stdcall HT_GetSummary(HT_Summary* out)
{
	if (!out) return;
	ZeroMemory(out, sizeof(*out));

	LONGLONG total = InterlockedCompareExchange64(&g_llTotalBytesAll, 0, 0);
	LONGLONG done = InterlockedCompareExchange64(&g_llDoneBytesAll, 0, 0);
	if (done > total) done = total;

	ULONGLONG nowTick = NowTick64();
	ULONGLONG startTick = g_ullOverallStartTick;

	double elapsedSec = 0.0;
	if (startTick != 0 && nowTick >= startTick) elapsedSec = (double)(nowTick - startTick) / 1000.0;
	if (elapsedSec < 0.001) elapsedSec = 0.001;

	double mbps = ((double)done / (1024.0 * 1024.0)) / elapsedSec;

	int pct = 0;
	if (total > 0) pct = (int)((double)done * 100.0 / (double)total + 0.5);
	if (pct > 100) pct = 100;
	if (pct < 0) pct = 0;

	out->percent = pct;
	out->totalBytes = (uint64_t)(total < 0 ? 0 : total);
	out->doneBytes = (uint64_t)(done < 0 ? 0 : done);
	out->mbps = mbps;
	out->runningCount = (int)InterlockedCompareExchange(&g_lRunningCount, 0, 0);
	out->poolThreads = (int)g_lPoolThreads;
}

int __stdcall HT_GetTextLength()
{
	BuildTextIfDirty_Throttle(FALSE);
	if (!g_pTextBuf) return 0;
	return (int)wcslen(g_pTextBuf);
}

int __stdcall HT_GetText(wchar_t* buf, int cch)
{
	if (!buf || cch <= 0) return 0;

	BuildTextIfDirty_Throttle(FALSE);

	if (!g_pTextBuf) {
		buf[0] = L'\0';
		return 0;
	}

	int len = (int)wcslen(g_pTextBuf);
	int toCopy = (len < (cch - 1)) ? len : (cch - 1);
	if (toCopy > 0) {
		memcpy(buf, g_pTextBuf, (size_t)toCopy * sizeof(wchar_t));
	}
	buf[toCopy] = L'\0';
	return toCopy;
}
