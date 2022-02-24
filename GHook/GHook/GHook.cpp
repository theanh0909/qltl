#include "pch.h"
#include "framework.h"
#include "GHook.h"

HHOOK hHook = NULL;
HMODULE hDll;

typedef VOID(*LOADPROC)(HHOOK hHook);

BOOL InstallHook();
BOOL UninstallHook();

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CGHookApp
BEGIN_MESSAGE_MAP(CGHookApp, CWinApp)
END_MESSAGE_MAP()

// CGHookApp construction
CGHookApp::CGHookApp()
{
	InstallHook();
}

CGHookApp::~CGHookApp()
{
	UninstallHook();
}

// The one and only CGHookApp object
CGHookApp theApp;

// CGHookApp initialization
BOOL CGHookApp::InitInstance()
{
	CWinApp::InitInstance();	
	return TRUE;
}

void __stdcall DLLStartHook()
{

}

/************************************************************************************************/
int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
	if (InstallHook() == FALSE) return -1;

	MSG msg;
	BOOL bResult;
	while ((bResult = GetMessage(&msg, NULL, 0, 0)) != 0)
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return 0;
}

BOOL InstallHook()
{
	
	if (LoadLibrary(L"GHookForm.dll") == NULL) return FALSE;

	HMODULE hDLL = GetModuleHandle(L"GHookForm");
	if (hDLL == NULL) return FALSE;

	hHook = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)GetProcAddress(hDLL, "MyHookKeyboard"), hDLL, 0);
	if (hHook == NULL) return FALSE;

	LOADPROC fPtrFcnt = (LOADPROC)GetProcAddress(hDLL, "SetGlobalHookHandle");
	if (fPtrFcnt == NULL) return FALSE;

	fPtrFcnt(hHook);

	return TRUE;
}

BOOL UninstallHook()
{
	if ((hHook != 0) && (UnhookWindowsHookEx(hHook) == TRUE))
	{
		hHook = NULL;
		return TRUE;
	}
	return FALSE;
}
