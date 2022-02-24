
// GRegistryCOM.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols

// CGRegistryCOMApp:
// See GRegistryCOM.cpp for the implementation of this class
//

class CGRegistryCOMApp : public CWinApp
{
public:
	CGRegistryCOMApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()


public:
	CString GetWinDir();
	CString GetPathDefault();

	char* ConvertCstringToChars(CString szvalue);
	
	void RegisterGProperties();
};

extern CGRegistryCOMApp theApp;
