
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

// CGHookFormApp
// See GHookForm.cpp for the implementation of this class
//

class CGHookFormApp : public CWinApp
{
public:
	CGHookFormApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()

};

void CALLBACK MyTimerProc(HWND hwnd, UINT uMsg, UINT timerId, DWORD dwTime);