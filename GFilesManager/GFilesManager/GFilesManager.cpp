
// GFilesManager.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "framework.h"
#include "GFilesManager.h"
#include "GFilesManagerDlg.h"
#include "CRegistryContextMenu.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CGFilesManagerApp

BEGIN_MESSAGE_MAP(CGFilesManagerApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CGFilesManagerApp construction

CGFilesManagerApp::CGFilesManagerApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

// The one and only CGFilesManagerApp object

CGFilesManagerApp theApp;


// CGFilesManagerApp initialization

BOOL CGFilesManagerApp::InitInstance()
{
	// Xóa file trong thư mục cài đặt
	Function ff;
	CString pathFolder = ff._GetPathFolder(szAppSmart);
	if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";
	CString szFileOld = pathFolder + FileGFileManaOld + L".exe";
	if (ff._FileExistsUnicode(szFileOld, 0) == 1)
	{
		DeleteFile(pathFolder + FileGFileManaOld + L".exe");
	}

	if ((int)_GetCountAppRun(szAppSmart) >= 2)
	{
		Base64Ex bb;
		CRegistry reg;
		CString szCreate = bb.BaseDecodeEx(CreateKeyRun);
		reg.CreateKey(szCreate);

		// Xóa file cũ trong registry (Đối với các máy đã từng sử dụng tên cũ)
		reg.DeleteValue(FileGFileManaOld);

		// Ứng dụng đang chạy --> return ra luôn
		return FALSE;
	}

	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("Local AppWizard-Generated Applications"));

	m_FontHeader.DeleteObject();
	m_FontHeader.CreateFont(16, 0, 0, 0, FW_BOLD, false, false, 0,
		ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
		DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, FontArial);

	CRegistryContextMenu *pReg = new CRegistryContextMenu;
	pReg->_StartRegContextMenu();
	delete pReg;

	// Xác định số lượng
	_GetListMyFiles(vecAllFiles);

	CGFilesManagerDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "Warning: dialog creation failed, so application is terminating unexpectedly.\n");
		TRACE(traceAppMsg, 0, "Warning: if you are using MFC controls on the dialog, you cannot #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	}

	// Delete the shell manager created above.
	if (pShellManager != nullptr)
	{
		delete pShellManager;
	}

#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
	ControlBarCleanUp();
#endif

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

int CGFilesManagerApp::_GetCountAppRun(CString szAppName)
{
	try
	{
		int dem = 0, pos = 0;
		Function *ff = new Function;		
		DWORD cbNeeded, aProcesses[1024];
		if (EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
		{
			unsigned int ii;
			CString szNameProc = L"";
			DWORD iTotalProc = cbNeeded / sizeof(DWORD);
			for (ii = 0; ii < iTotalProc; ii++)
			{
				szNameProc = ff->_GetDirModules(aProcesses[ii]);
				if (szNameProc != L"")
				{
					szNameProc = ff->_GetNameOfPath(szNameProc, pos, 0);
					if (szNameProc == szAppName)
					{
						dem++;
						if (dem >= 2) break;
					}
				}
			}
		}
		return dem;
	}
	catch (...) {}
	return 0;
}


