#pragma once

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#ifndef WINVER
#define WINVER 0x0601
#endif

#include <windows.h>
#include <stdint.h>

#ifdef HASHTOOLCORE_EXPORTS
#define HT_API extern "C" __declspec(dllexport)
#else
#define HT_API extern "C" __declspec(dllimport)
#endif

typedef void(__stdcall* HT_OnDirty)(void* user);

typedef struct HT_Summary {
	int percent;                 // 0..100
	uint64_t totalBytes;
	uint64_t doneBytes;
	double mbps;
	int runningCount;
	int poolThreads;
} HT_Summary;

// 初始化/释放
HT_API BOOL  __stdcall HT_Init(HT_OnDirty cb, void* user);
HT_API void  __stdcall HT_Shutdown();

// 线程池线程数
HT_API void  __stdcall HT_SetThreadCount(int n);

// 任务控制
HT_API BOOL  __stdcall HT_AddFile(const wchar_t* path, BOOL md5, BOOL sha256);
HT_API void  __stdcall HT_CancelAll();
HT_API BOOL  __stdcall HT_ClearAll(); // running!=0 返回 FALSE

// 查询
HT_API void  __stdcall HT_GetSummary(HT_Summary* out);

// 文本获取（UTF-16）
HT_API int   __stdcall HT_GetTextLength();               // 字符数（不含 \0）
HT_API int   __stdcall HT_GetText(wchar_t* buf, int cch); // 返回写入字符数（不含 \0）
