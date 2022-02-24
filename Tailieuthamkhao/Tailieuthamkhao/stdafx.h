#pragma once

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

#include "afxwinappex.h"
#include "afxdialogex.h"
#include "Resource.h"
#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <stdio.h>
#include <stdlib.h>
#include <fstream>
#include <codecvt>
#include <sys/stat.h>

#include <odbcinst.h>
//#pragma comment (lib, "Odbcinst")	// Link with Odbcinst.lib for 16-bit applications
//#pragma comment (lib, "Odbccp32")	// Link with Odbccp32.lib for 32-bit applications

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

#include "Color.h"
#include "Define.h"
#include "EasySize.h"
#include "Connection.h"

#include "msolib.h"
#include <afxcontrolbars.h>

#include "ColorEdit.h"
#include "SysImageList.h"

#include "BtnST.h"
#include "XPStyleButtonST.h"
#include "ListCtrlEx.h"
#include "ComboBoxExt.h"
#include "ComboBoxExtList.h"
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

// Excel application
extern _ApplicationPtr xl;
extern _WorkbookPtr pWb;

extern CConnection ObjConn;
extern CString SqlString;

extern CString spathDefault;
extern int RowEditList, ColumnEditList;

extern double jScaleLayout;

extern int jCheckRotot;
extern int iSaveCopyMulti;
extern int iSaveCreateFolder;
extern int irHeightDOWN, irWidthDOWN;

extern CFont myfontCheckbot, m_FontHeaderList;