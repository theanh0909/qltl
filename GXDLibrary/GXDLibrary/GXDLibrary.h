// GXDLibrary.h : main header file for the GXDLibrary DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CGXDLibraryApp
// See GXDLibrary.cpp for the implementation of this class
//

class CGXDLibraryApp : public CWinApp
{
public:
	CGXDLibraryApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};

extern CGXDLibraryApp theApp;