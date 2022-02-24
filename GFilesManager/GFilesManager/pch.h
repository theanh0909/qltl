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

#include <psapi.h>
#pragma comment(lib, "psapi.lib")

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

#include "Define.h"
#include "Base64.h"
#include "EasySize.h"
#include "Registry.h"
#include "Util.h"
#include "FileFinder.h"
#include "ColorEdit.h"
#include "MenuIcon.h"
#include "SysImageList.h"
#include "ListCtrlEx.h"
#include "NumSpinCtrl.h"
#include "XPStyleButtonST.h"
#include "TextProgressCtrl.h"
#include "MultiFileOpenDialog.h"
#include "Function.h"

#include <afxcontrolbars.h>
#include <afxcontrolbars.h>

extern CFont m_FontHeader;

extern int jColumnHand;
extern int dPosX, dPosY, dWidthF, dHeightF;
extern int jXIcon, jYIcon;
extern int nIndexMouse;
extern bool blDragging;
extern CString szAppSmart;

struct MyStrList
{
	int iID;
	CString szFullpath;
	COLORREF clrBkg;
};

extern vector<MyStrList> vecAllFiles;

// Xác định số lượng file sử dụng
int _GetListMyFiles(vector<MyStrList> &vecItem);

#endif //PCH_H
