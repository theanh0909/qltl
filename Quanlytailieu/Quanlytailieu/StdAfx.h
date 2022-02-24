// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently

#pragma once
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

#import "..\\Library\\GProperties.tlb"

//#if defined _M_X64
//	#import "..\\Library\\x64\\GProperties.tlb"
//#else
//	#import "..\\Library\\x86\\GProperties.tlb"
//#endif

using namespace GProperties;

using namespace std;

#include "Define.h"
#include "msolib.h"
#include "Base64.h"
#include "EasySize.h"
#include "Registry.h"
#include "Util.h"
#include "FileFinder.h"
#include "FileFilter.h"
#include "ColorEdit.h"
#include "MenuIcon.h"
#include "SysImageList.h"
#include "PictureCtrl.h"
#include "ListCtrlEx.h"
#include "NumSpinCtrl.h"
#include "XPStyleButtonST.h"
#include "TextProgressCtrl.h"
#include "TaskbarNotifier.h"
#include "ColumnTreeCtrl.h"
#include "ColumnTreeView.h"
#include "ColumnTreeWnd.h"
#include "SplitterControl.h"
#include "ComboBoxExt.h"
#include "ComboBoxExtList.h"
#include "Function.h"

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // Exclude rarely-used stuff from Windows headers
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE classes
#include <afxodlgs.h>       // MFC OLE dialog classes
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>                      // MFC ODBC database classes
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>                     // MFC DAO database classes
#endif // _AFX_NO_DAO_SUPPORT

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC support for Internet Explorer 4 Common Controls
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>                     // MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT
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

extern CString strNAMEDLL;
extern CString szOffice;

extern _ApplicationPtr xl;
extern _WorkbookPtr pWb;

extern _WorksheetPtr pShActive;
extern RangePtr pRgActive;
extern CString sCodeActive;
extern CString sNameActive;
extern COLORREF bkgColorActive;
extern int iRowActive, iColumnActive;

extern bool blDragging;
extern int CLRow, CLColumn;
extern CColorEdit m_edit_rename, m_edit_properties, m_edit_autorename;

extern bool bCheckVirut;
extern bool bCheckVersion;
extern bool bRibbonSearch;

extern int jColumnHand;
extern CFont m_FontHeaderList;
extern CFont m_FontItalic14, m_FontItalic;

extern HWND hWndAutoRename;

struct MyStrSelection
{
	int row;
	int column;
	CString cell;

	CString sheet;
	COLORREF bkgcolor;

	CString value;
	CString formula;
};

extern vector<MyStrSelection> vecSelect;
extern bool blUndovalue;

// Hàm bổ trợ cho hàm sắp xếp qsort
// Ví dụ: qsort(vecItems.data(), (int)vecItems.size(), sizeof(<type of vector>), compareAZ);
// vecItems = Vector được sử dụng để sắp xếp dữ liệu từ A -> Z
// <type of vector> = Kiểu dữ liệu vector sử dụng
extern int compareAZ(const void * a, const void * b);
extern bool blCheckFileTemp();
