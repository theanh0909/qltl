// Tailieuthamkhao.h : main header file for the Tailieuthamkhao DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "Resource.h"		// main symbols


// CTailieuthamkhaoApp
// See Tailieuthamkhao.cpp for the implementation of this class
//

class CTailieuthamkhaoApp : public CWinApp
{
public:
	CTailieuthamkhaoApp();
	~CTailieuthamkhaoApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()

};

extern CTailieuthamkhaoApp theApp;
