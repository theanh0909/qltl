
// sl.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols

#include "Base64.h"
#include "Registry.h"

// CslApp:
// See sl.cpp for the implementation of this class
//

class CslApp : public CWinApp
{
public:
	CslApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()

public:
	Base64Ex *bb;
	CRegistry *reg;

	CString _GetPathFolder(CString fileName);
	int _FileExistsUnicode(CString pathFile);
	char* _ConvertCstringToChars(CString szvalue);
	bool _IsProcessRunning(const wchar_t *processName);
	int _GetContextMenuParameter(vector<CString> &vecItems);
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);
	void _WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype = 0);
	CString _GetDirModules(DWORD processID, bool blFilter = false);
	CString _GetNameOfPath(CString pathFile, int ipath = 0);
	void _QuitAppSmartStart(CString szFileName);
};

extern CslApp theApp;
