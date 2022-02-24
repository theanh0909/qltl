// GXDUpdateInfo.h : main header file for the GXDUpdateInfo DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CGXDUpdateInfoApp
// See GXDUpdateInfo.cpp for the implementation of this class
//

class CGXDUpdateInfoApp : public CWinApp
{
public:
	CGXDUpdateInfoApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};
