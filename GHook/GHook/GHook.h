// GHook.h : main header file for the GHook DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CGHookApp
// See GHook.cpp for the implementation of this class
//

class CGHookApp : public CWinApp
{
public:
	CGHookApp();
	~CGHookApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};
