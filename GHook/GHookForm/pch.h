// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

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

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

#include "Define.h"
#include "msolib.h"
#include "Base64.h"
#include "Registry.h"
#include "SendKeys.h"
#include "Connection.h"
#include "SysImageList.h"
#include "ColorEdit.h"
#include "ListCtrlEx.h"
#include "Function.h"

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

// Microsoft Excel
extern _ApplicationPtr xl;

extern _WorksheetPtr pShActive;
extern RangePtr pRgActive;
extern CString sCodeActive;
extern CString sNameActive;
extern int iRowActive, iColumnActive;

extern CConnection ObjConn;
extern CString SqlString;

// Handle hộp thoại tìm kiếm chi phí
extern HWND hWndChiphi;

// Handle hộp thoại tìm kiếm quy trình
extern HWND hWndQuytrinh;

// Handle hộp thoại tìm kiếm tiêu chuẩn
extern HWND hWndTieuchuan;

// Handle hộp thoại chính (Excel)
extern HWND hWndMain;

// Giá trị trước và sau khi nhập vào ô cell
extern CString szCellBefore, szcellAfter;

// Ký tự nhập trên bàn phím
extern CString szKeyboard;

// Giá trị form hiển thị
// nIndexForm = 1 --> Hộp thoại tìm kiếm quy trình
// nIndexForm = 2 --> Hộp thoại tìm kiếm tiêu chuẩn
extern int nIndexForm;

// Giá trị sử dụng để phân biệt khi Hook Mouse
// nIndexMouse = 0 --> Giá trị mặc định (Global)
// nIndexMouse > 0 --> Chỉ sử dụng trong hộp thoại tìm kiếm tiêu chuẩn
extern int nIndexMouse;

// Phím bấm kế trước là phím F2
// jKeyBeforeF2 = 0 --> Bấm phím F2
// jKeyBeforeF2 = 1 --> Phím F2 là phím bấm liền trước (Cần quan tâm giá trị này)
// jKeyBeforeF2 > 1 --> Đã sử dụng nhiều phím khác sau phím F2
extern int jKeyBeforeF2;

// Phím bấm là phím ENTER
extern bool blKeyEnter;

// Tọa độ chuột & kích thước hộp thoại tìm kiếm tiêu chuẩn
extern int dPosX, dPosY, dWidthF, dHeightF;

// Fon chữ tiêu đề áp dụng cho tất cả các listcontrol
extern CFont m_FontHeaderList;

// Biến vị trí ListcontrolEx
extern int CLRow, CLColumn;

// Cột chứa hyperlink trên ListcontrolEx
extern int jColumnHand;

static const int CONST_TIMER_MAIN			= 0;
static const int CONST_TIMER_QUYTRINH		= 1;
static const int CONST_TIMER_TIEUCHUAN		= 2;
static const int CONST_TIMER_CHIPHI			= 3;

#endif //PCH_H