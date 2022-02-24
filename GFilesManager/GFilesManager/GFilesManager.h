
// GFilesManager.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "Resource.h"		// main symbols


// CGFilesManagerApp:
// See GFilesManager.cpp for the implementation of this class
//

class CGFilesManagerApp : public CWinApp
{
public:
	CGFilesManagerApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()

protected:
	int _GetCountAppRun(CString szAppName);

};

extern CGFilesManagerApp theApp;
