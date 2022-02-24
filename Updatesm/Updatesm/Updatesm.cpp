
// Updatesm.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "framework.h"
#include "Updatesm.h"
#include "UpdatesmDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CUpdatesmApp

BEGIN_MESSAGE_MAP(CUpdatesmApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CUpdatesmApp construction

CUpdatesmApp::CUpdatesmApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CUpdatesmApp object

CUpdatesmApp theApp;


// CUpdatesmApp initialization

BOOL CUpdatesmApp::InitInstance()
{
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

	CUpdatesmDlg dlg;
	m_pMainWnd = &dlg;

	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;

	// Hàm xử lý quan trọng
	_CheckUpdateApp();

	delete sv;
	delete bb;
	delete reg;

	//INT_PTR nResponse = dlg.DoModal();
	//if (nResponse == IDOK)
	//{
	//	// TODO: Place code here to handle when the dialog is
	//	//  dismissed with OK
	//}
	//else if (nResponse == IDCANCEL)
	//{
	//	// TODO: Place code here to handle when the dialog is
	//	//  dismissed with Cancel
	//}
	//else if (nResponse == -1)
	//{
	//	TRACE(traceAppMsg, 0, "Warning: dialog creation failed, so application is terminating unexpectedly.\n");
	//	TRACE(traceAppMsg, 0, "Warning: if you are using MFC controls on the dialog, you cannot #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	//}

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

void CUpdatesmApp::_CheckUpdateApp()
{
	// Đóng tất cả các ứng dụng 'SmartStart' đang bật
	_QuitAppSmartStart();

	CString pathFolder = _GetPathFolder(FileUpdatesm);
	if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

	CString szWin = L"32\\";
	if (_Is64BitWindows()) szWin = L"64\\";

	// Tải dữ liệu
	sv->_LoadDecodeBase64();
	CString szFileApp = pathFolder + FileSmartStartApp;
	if (sv->_DownloadFile(sv->dbSeverDir + szWin + FileSmartStartApp, szFileApp))
	{
		// Ghi thông tin version vào Registry
		CString szFileVersion = pathFolder + FileVersionsm;
		if (sv->_DownloadFile(sv->dbSeverDir + FileVersionsm, szFileVersion))
		{
			vector<CString> vecLine;
			int ncount = _ReadMultiUnicode(szFileVersion, vecLine);
			if (ncount > 0)
			{
				CString szCreate = bb->BaseDecodeEx(CreateKeySettings);
				reg->CreateKey(szCreate);
				reg->WriteChar(
					_ConvertCstringToChars(Version),
					_ConvertCstringToChars(vecLine[0]));
			}
			vecLine.clear();
			DeleteFile(szFileVersion);
		}

		// Mở tiện ích 'SmartStart'
		ShellExecute(NULL, L"open", szFileApp, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

int CUpdatesmApp::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExistsUnicode(spathFile) != 1) return 0;

		wifstream readfile_(_ConvertCstringToChars(spathFile));
		readfile_.imbue(locale(locale::empty(), new codecvt_utf8<wchar_t, 0x10ffff, consume_header>));

		if (readfile_.is_open())
		{
			wstring username;
			while (getline(readfile_, username))
			{
				vecTextLine.push_back(username.c_str());
			}

			readfile_.close();
		}

		return (int)vecTextLine.size();
	}
	catch (...) {}

	return 0;
}

int CUpdatesmApp::_FileExistsUnicode(CString pathFile)
{
	try
	{
		LPTSTR lpPath = pathFile.GetBuffer(MAX_PATH);
		if (PathFileExists(lpPath)) return 1;
		return -1;
	}
	catch (...)
	{
		return -1;
	}
	return 0;
}

char* CUpdatesmApp::_ConvertCstringToChars(CString szvalue)
{
	wchar_t* wszString = new wchar_t[szvalue.GetLength() + 1];
	wcscpy(wszString, szvalue);
	int num = WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* chr = new char[num + 1];
	WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), chr, num, NULL, NULL);
	chr[num] = '\0';
	delete[] wszString;
	return chr;
}


CString CUpdatesmApp::_GetPathFolder(CString fileName)
{
	try
	{
		HMODULE hModule;
		hModule = GetModuleHandle(fileName);
		TCHAR szFileName[MAX_LEN];
		DWORD dSize = MAX_LEN;
		GetModuleFileName(hModule, szFileName, dSize);
		CString szSource = szFileName;
		for (int i = szSource.GetLength() - 1; i > 0; i--)
		{
			if (szSource.GetAt(i) == '\\')
			{
				szSource = szSource.Left(szSource.GetLength() - (szSource.GetLength() - i) + 1);
				break;
			}
		}
		return szSource;
	}
	catch (...) {}
	return L"";
}

bool CUpdatesmApp::_IsProcessRunning(const wchar_t *processName)
{
	bool exists = false;
	PROCESSENTRY32 entry;
	entry.dwSize = sizeof(PROCESSENTRY32);

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);

	if (Process32First(snapshot, &entry))
		while (Process32Next(snapshot, &entry))
			if (!wcsicmp(entry.szExeFile, processName))
				exists = true;

	CloseHandle(snapshot);
	return exists;
}

CString CUpdatesmApp::_GetDirModules(DWORD processID, bool blFilter)
{
	try
	{
		CString szProcessDir = L"";
		HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID);
		if (hProcess != NULL)
		{
			WCHAR processDir[MAX_PATH];
			if ((DWORD)GetProcessImageFileName(hProcess, processDir,
				sizeof(processDir) / sizeof(*processDir)) != 0)
			{
				if (blFilter)
				{
					DWORD cbNeeded;
					HMODULE hMods[1024];
					if (EnumProcessModulesEx(hProcess, hMods,
						sizeof(hMods), &cbNeeded, LIST_MODULES_ALL))
					{
						szProcessDir = (LPCTSTR)processDir;
					}					
				}
				else
				{
					szProcessDir = (LPCTSTR)processDir;
				}
			}
		}
		CloseHandle(hProcess);

		return szProcessDir;
	}
	catch (...) {}
	return L"";
}

CString CUpdatesmApp::_GetNameOfPath(CString pathFile, int ipath)
{
	try
	{
		if (pathFile != L"")
		{
			if (pathFile.Right(1) == L"\\") pathFile = pathFile.Left(pathFile.GetLength() - 1);

			int nLen = pathFile.GetLength();
			for (int i = nLen - 1; i >= 0; i--)
			{
				if (pathFile.Mid(i, 1) == L"\\" || pathFile.Mid(i, 1) == L"/")
				{
					if (ipath == -1)
					{
						// Trả về kết quả là tên file
						CString szfile = pathFile.Right(nLen - i - 1);
						int nLenName = szfile.GetLength();
						for (int j = nLenName - 1; j >= 0; j--)
						{
							if (szfile.Mid(j, 1) == L".")
							{
								return szfile.Left(j);
							}
						}
						return szfile;
					}
					else if (ipath == 0)
					{
						// Trả về kết quả là tên file + đuôi
						return pathFile.Right(nLen - i - 1);
					}

					// Trả về đường dẫn chứa file
					return pathFile.Left(i);
				}
			}
		}
	}
	catch (...) {}
	return L"";
}

bool CUpdatesmApp::_Is64BitWindows()
{
#if defined(_WIN64)
	return TRUE;  // 64-bit programs run only on Win64
#elif defined(_WIN32)
	// 32-bit programs run on both 32-bit and 64-bit Windows
	// so must sniff
	BOOL f64 = FALSE;
	return IsWow64Process(GetCurrentProcess(), &f64) && f64;
#else
	return FALSE; // Win64 does not support Win16
#endif
}

void CUpdatesmApp::_QuitAppSmartStart()
{
	try
	{
		if (_IsProcessRunning(FileSmartStartApp))
		{
			DWORD cbNeeded;
			DWORD aProcesses[1024];
			if (EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
			{
				HANDLE tmpH;
				unsigned int ii;
				CString szNameProc = L"";
				DWORD iTotalProc = cbNeeded / sizeof(DWORD);
				for (ii = 0; ii < iTotalProc; ii++)
				{
					szNameProc = _GetDirModules(aProcesses[ii]);
					if (szNameProc != L"")
					{
						szNameProc = _GetNameOfPath(szNameProc);
						if (szNameProc == FileSmartStartApp)
						{
							tmpH = OpenProcess(PROCESS_ALL_ACCESS, TRUE, aProcesses[ii]);
							if (tmpH != NULL) TerminateProcess(tmpH, 0);
						}
					}
				}
			}
		}
	}
	catch (...) {}
}