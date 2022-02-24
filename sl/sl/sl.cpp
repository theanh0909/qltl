
// sl.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "framework.h"
#include "sl.h"
#include "slDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define MAX_LEN					1024
#define FileSelected			L"GMSelect.exe"
#define FileSmartStartApp		L"SmartStart.exe"
#define FileMyFiles				L"ListMyFiles.ini"
#define CreateKeySMStart		L"YwWE5kWUowWEZObGRIUnBibWR6ZzRUlVW4=E6REGVRVTAJ51ZZa1ZjUjFoRVNsTkRYRlhZTj10SU"

// CslApp

BEGIN_MESSAGE_MAP(CslApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CslApp construction

CslApp::CslApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CslApp object

CslApp theApp;


// CslApp initialization

BOOL CslApp::InitInstance()
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

	CslDlg dlg;
	m_pMainWnd = &dlg;

	vector<CString> vecItems;
	int ncount = _GetContextMenuParameter(vecItems);
	if (ncount > 0)
	{
		CString pathFolder = _GetPathFolder(FileSelected);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		vector<CString> vecRead;
		CString szPathFile = pathFolder + FileMyFiles;

		ncount = _ReadMultiUnicode(szPathFile, vecRead);
		if (ncount > 0) for (int i = 0; i < ncount; i++) vecItems.push_back(vecRead[i]);
		vecRead.clear();

		_WriteMultiUnicode(vecItems, szPathFile, 0);

/**************************************************************************************/
		
		if (_IsProcessRunning(FileSmartStartApp))
		{
			bb = new Base64Ex;
			reg = new CRegistry;

			// Chú ý: Giá trị 'CheckAddFiles' --> Liên hệ với DLL GFilesManager (SmartStart.exe)
			reg->CreateKey(bb->BaseDecodeEx(CreateKeySMStart));
			reg->WriteChar("CheckAddFiles", "1");

			delete reg;
			delete bb;
		}
	}
	vecItems.clear();

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

int CslApp::_GetContextMenuParameter(vector<CString> &vecItems)
{
	try
	{
		vecItems.clear();

		int nArgs = 0;
		LPWSTR *szArglist = CommandLineToArgvW(GetCommandLineW(), &nArgs);
		if (szArglist != NULL)
		{
			if (nArgs > 1)
			{
				for (int i = 1; i < nArgs; i++)
				{
					if (szArglist[i] != L"")
					{
						vecItems.push_back((LPCTSTR)szArglist[i]);
					}					
				}
			}
		}

		LocalFree(szArglist);

		return (int)vecItems.size();
	}
	catch (...) {}
	return 0;
}

void CslApp::_WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (_FileExistsUnicode(spathFile) != 1) return;
		}

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(spathFile));
		if (myfile)
		{
			myfile.imbue(loc);
			long total = (long)vecTextLine.size();
			for (long i = 0; i < total; i++)
			{
				wstring ws(vecTextLine[i]);
				myfile << ws << "\n";
			}
			myfile.close();
		}
	}
	catch (...) {}
}

int CslApp::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
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

int CslApp::_FileExistsUnicode(CString pathFile)
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

char* CslApp::_ConvertCstringToChars(CString szvalue)
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

CString CslApp::_GetPathFolder(CString fileName)
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

bool CslApp::_IsProcessRunning(const wchar_t *processName)
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

CString CslApp::_GetDirModules(DWORD processID, bool blFilter)
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

CString CslApp::_GetNameOfPath(CString pathFile, int ipath)
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

void CslApp::_QuitAppSmartStart(CString szFileName)
{
	try
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
					if (szNameProc == szFileName)
					{
						tmpH = OpenProcess(PROCESS_ALL_ACCESS, TRUE, aProcesses[ii]);
						if (tmpH != NULL) TerminateProcess(tmpH, 0);
					}
				}
			}
		}
	}
	catch (...) {}
}
