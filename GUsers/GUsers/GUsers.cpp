// GUsers.cpp : Defines the initialization routines for the DLL.
//

#include "pch.h"
#include "GUsers.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CGUsersApp

BEGIN_MESSAGE_MAP(CGUsersApp, CWinApp)
END_MESSAGE_MAP()

// CGUsersApp construction

CGUsersApp::CGUsersApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

// The one and only CGUsersApp object

CGUsersApp theApp;

// CGUsersApp initialization

BOOL CGUsersApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

void __stdcall GUSersNull()
{

}

void __stdcall GetDecentralization()
{
	try
	{
		Base64Ex* bb = new Base64Ex;		
		HardDisk* hd = new HardDisk;
		Function* ff = new Function;
		CRegistry* reg = new CRegistry;

		CString szHardDisk = bb->BaseEncodeEx(hd->_CheckHardWareID(), 0);

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings);
		reg->CreateKey(szCreate);
		reg->WriteChar("UserDefault", ff->_ConvertCstringToChars(szHardDisk));

		delete ff; delete bb; delete hd; delete reg;
	}
	catch (...) {}
}
