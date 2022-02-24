
// QLTL GXD.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "msolib.h"
#include <vector>

using namespace std;

// CQLTLGXDApp:
// See QLTL GXD.cpp for the implementation of this class
//

class CQLTLGXDApp : public CWinApp
{
public:
	CQLTLGXDApp();

// Overrides
public:
	virtual BOOL InitInstance();
	DECLARE_MESSAGE_MAP()

public:
	CString szpathDefault;

	bool _OpenTemplate();
	char* _ConvertCstringToChars(CString szvalue);
	CString _ConvertCharsToCstring(char* chr);
	CString _ReadFileUnicode(CString spathFile);
	CString _GetPathDefault();	
	int _FileExistsUnicode(CString pathFile);
};

extern CQLTLGXDApp theApp;
