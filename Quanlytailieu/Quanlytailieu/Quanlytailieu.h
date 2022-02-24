// Quanlytailieu.h : main header file for the Quanlytailieu DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols

// CQuanlytailieuApp
// See Quanlytailieu.cpp for the implementation of this class
//

class CQuanlytailieuApp : public CWinApp
{
public:
	CQuanlytailieuApp();
	~CQuanlytailieuApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()

};

extern CQuanlytailieuApp theApp;
