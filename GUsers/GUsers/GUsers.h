// GUsers.h : main header file for the GUsers DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CGUsersApp
// See GUsers.cpp for the implementation of this class
//

class CGUsersApp : public CWinApp
{
public:
	CGUsersApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};
