
// SetupGXD.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CSetupGXDApp:
// See SetupGXD.cpp for the implementation of this class
//

class CSetupGXDApp : public CWinApp
{
public:
	CSetupGXDApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()

public:
	CString _GetPathDefault();
	int _FileExistsUnicode(CString pathFile);
};

extern CSetupGXDApp theApp;
