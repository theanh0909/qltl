// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

// add headers that you want to pre-compile here
#include "framework.h"

#endif //PCH_H

#include "framework.h"
#include "afxdialogex.h"
#include "resource.h"
#include <windows.h>
#include <iostream>
#include <vector>
#include <string>
#include <stdio.h>
#include <stdlib.h>
#include <fstream>
#include <codecvt>
#include <sys/stat.h>
#include <direct.h>
#include <cstdio>
#include <tlhelp32.h>

#include <psapi.h>
#pragma comment( lib, "psapi.lib" )

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

int _FileExistsUnicode(CString pathFile);
char* _ConvertCstringToChars(CString szvalue);
CString _GetTimeNow(int ifulltime);
bool _IsProcessRunning(const wchar_t *processName);
CString _GetDirModules(DWORD processID, bool blFilter);