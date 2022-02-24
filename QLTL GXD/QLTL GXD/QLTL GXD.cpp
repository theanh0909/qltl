
// QLTL GXD.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "QLTL GXD.h"
#include "QLTL GXDDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define fileNameApp		L"QLTL GXD.exe"

// CQLTLGXDApp
BEGIN_MESSAGE_MAP(CQLTLGXDApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CQLTLGXDApp construction

CQLTLGXDApp::CQLTLGXDApp()
{
	szpathDefault = L"";
}


// The one and only CQLTLGXDApp object

CQLTLGXDApp theApp;


// CQLTLGXDApp initialization

BOOL CQLTLGXDApp::InitInstance()
{
	CWinApp::InitInstance();


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

	CQLTLGXDDlg dlg;
	m_pMainWnd = &dlg;

	// Check bản quyền...
	// Tạm thời đang sử dụng miễn phí

	_OpenTemplate();

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

	//// Delete the shell manager created above.
	//if (pShellManager != nullptr)
	//{
	//	delete pShellManager;
	//}

#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
	ControlBarCleanUp();
#endif

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

CString CQLTLGXDApp::_GetPathDefault()
{
	try
	{
		HMODULE hModule;
		hModule = GetModuleHandle(fileNameApp);
		TCHAR szFileName[MAX_PATH];
		DWORD dSize = MAX_PATH;
		GetModuleFileName(hModule, szFileName, dSize);
		CString szSource = szFileName;
		for (int i = szSource.GetLength() - 1; i > 0; i--)
		{
			if (szSource.GetAt(i) == '\\')
			{
				szSource = szSource.Left(szSource.GetLength() - (szSource.GetLength() - i) + 1);
				if (szSource.Right(1) != L"\\") szSource += L"\\";
				break;
			}
		}

		return szSource;
	}
	catch (...) {}
	return L"";
}

char* CQLTLGXDApp::_ConvertCstringToChars(CString szvalue)
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

CString CQLTLGXDApp::_ConvertCharsToCstring(char* chr)
{
	int num = MultiByteToWideChar(CP_UTF8, 0, chr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	MultiByteToWideChar(CP_UTF8, 0, chr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	delete[] wstr;
	return szval;
}

CString CQLTLGXDApp::_ReadFileUnicode(CString spathFile)
{
	try
	{
		CString szTextLine = L"";
		if (_FileExistsUnicode(spathFile) != 1)
		{
			wifstream readfile_(_ConvertCstringToChars(spathFile));
			readfile_.imbue(locale(locale::empty(), new codecvt_utf8<wchar_t, 0x10ffff, consume_header>));

			if (readfile_.is_open())
			{
				wstring username;
				while (getline(readfile_, username))
				{
					szTextLine = username.c_str();
					break;
				}

				readfile_.close();
			}
		}		

		return szTextLine;
	}
	catch (...) {}
	return L"";
}

int CQLTLGXDApp::_FileExistsUnicode(CString pathFile)
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

bool CQLTLGXDApp::_OpenTemplate()
{
	try
	{
		CoInitialize(NULL);

		szpathDefault = _GetPathDefault();

		CString szFileExcel = _ReadFileUnicode(szpathDefault + L"temp.ini");
		if (szFileExcel == L"") szFileExcel = szpathDefault + L"Temp\\Quanlytailieu.xltm";
		if (_FileExistsUnicode(szFileExcel))
		{
			// Load the Excel application in the background.
			_ApplicationPtr xl;
			if (FAILED(xl.GetActiveObject(L"Excel.Application")))
			{
				if (FAILED(xl.CreateInstance(L"Excel.Application")))
				{
					MessageBoxW(NULL, L"Có lỗi trong quá trình khởi động Excel. Bạn hãy cài lại office hoặc thay thế một phiên bản office khác và thử lại", L"Hãy kiểm tra lại Office", MB_ICONERROR);
					return FALSE;
				}
			}

			xl->PutVisible(true);
			xl->Workbooks->Add((_bstr_t)szFileExcel);
			_WorkbookPtr pWb = xl->GetActiveWorkbook();

			_WorksheetPtr psh;
			_bstr_t bsConfig = "";
			int ncount = xl->ActiveWorkbook->Sheets->Count;
			
			for (int i = 1; i <= ncount; i++)
			{
				psh = pWb->Worksheets->Item[i];
				if (psh->CodeName == (_bstr_t)"shConfig")
				{
					bsConfig = psh->Name;
					break;
				}
			}

			if (bsConfig != (_bstr_t)"")
			{
				_WorksheetPtr psConfig = xl->Sheets->GetItem(bsConfig);
				psConfig->Cells->GetRange("CF_DIR_DEFAULT", "CF_DIR_DEFAULT")->PutValue2((_bstr_t)szpathDefault);
			}

			// Mở file 'GRibbon.xlam'
			CString szFileXLAM = szpathDefault + L"GRibbon.xlam";
			if (_FileExistsUnicode(szFileXLAM))
			{
				xl->Workbooks->Open((_bstr_t)szFileXLAM);

				_variant_t varUnused((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
				CString PATH_MODULEXLA = L"GRibbon.xlam!mFunction.VBADLLStartEXE";
				_variant_t szMacro(PATH_MODULEXLA);

				xl->PutScreenUpdating(VARIANT_FALSE);
				xl->Run(szMacro, 
					varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
					varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
					varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
					varUnused, varUnused, varUnused, varUnused, varUnused, varUnused);
				xl->PutScreenUpdating(VARIANT_TRUE);
			}
			else
			{
				MessageBox(NULL,
					L"Không tồn tại file 'GRibbon.xlam'. \nVui lòng kiểm tra lại.",
					L"Thông báo", MB_OK | MB_ICONERROR);
			}
		}
		else
		{
			MessageBox(NULL,
				L"Không tồn tại file template. \nVui lòng kiểm tra lại đường dẫn: " + szFileExcel,
				L"Thông báo", MB_OK | MB_ICONERROR);
		}
	}
	catch (...) {
		MessageBox(NULL,
			L"Xảy ra lỗi khi khởi động phần mềm. Vui lòng tắt đi và bật lại.",
			L"Lỗi phần mềm", MB_OK | MB_ICONERROR);
	}
	CoUninitialize();
}