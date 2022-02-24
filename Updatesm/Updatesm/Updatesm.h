
// Updatesm.h : main header file for the PROJECT_NAME application
//

#pragma once
#include "CDownloadFileSever.h"

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CUpdatesmApp:
// See Updatesm.cpp for the implementation of this class
//

class CUpdatesmApp : public CWinApp
{
public:
	CUpdatesmApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()

public:

	Base64Ex *bb;
	CRegistry *reg;
	CDownloadFileSever *sv;

	void _CheckUpdateApp();
	void _QuitAppSmartStart();
	CString _GetPathFolder(CString fileName);
	int _FileExistsUnicode(CString pathFile);
	char* _ConvertCstringToChars(CString szvalue);	
	bool _IsProcessRunning(const wchar_t *processName);
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);
	CString _GetDirModules(DWORD processID, bool blFilter = false);
	CString _GetNameOfPath(CString pathFile, int ipath = 0);
	bool _Is64BitWindows();
};

extern CUpdatesmApp theApp;
