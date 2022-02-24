// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H

// add headers that you want to pre-compile here
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

#include <odbcinst.h>
//#pragma comment (lib, "Odbcinst")	// Link with Odbcinst.lib for 16-bit applications
//#pragma comment (lib, "Odbccp32")	// Link with Odbccp32.lib for 32-bit applications

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

#include "Define.h"
#include "msolib.h"
#include "Connection.h"
#include "Registry.h"
#include "Util.h"
#include "FileFinder.h"
#include "FileFilter.h"

#include "CFunction.h"

#include <afxcontrolbars.h>

#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif

// Application Excel
extern _ApplicationPtr xl;
extern _WorkbookPtr pWb;

// Thông tin sheet active
extern _WorksheetPtr pShActive;
extern RangePtr pRgActive;
extern CString szCodeActive, szNameActive;
extern COLORREF bkgColorActive;
extern int iRowActive, iColumnActive;

// Truy vấn dữ liệu Database
extern CConnection ObjConn;
extern CString SqlString;

// Hàm bổ trợ cho hàm sắp xếp qsort
// Ví dụ: qsort(vecItems.data(), (int)vecItems.size(), sizeof(<type of vector>), compareStructAZ);
// vecItems = Vector được sử dụng để sắp xếp dữ liệu từ A -> Z
// <type of vector> = Kiểu dữ liệu vector sử dụng
extern int compareStringAZ(const void * a, const void * b);

#endif //PCH_H