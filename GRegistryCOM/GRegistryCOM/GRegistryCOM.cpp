
// GRegistryCOM.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "framework.h"
#include "GRegistryCOM.h"
#include "GRegistryCOMDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define fileNameApp				L"GRegistryCOM.exe"
#define fileDSO					L"dsofile.dll"
#define CreateKeySettings		L"WUZkVFJ4VFpYUjBhVzVuYzF4VgR=DJTV=VPOVYTbKlBZSFJURlhFZFlSRXBUUTFGVHg9UlFF"
#define EditProperties			L"EditProperties"

// CGRegistryCOMApp
BEGIN_MESSAGE_MAP(CGRegistryCOMApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CGRegistryCOMApp construction

CGRegistryCOMApp::CGRegistryCOMApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CGRegistryCOMApp object

CGRegistryCOMApp theApp;


// CGRegistryCOMApp initialization

BOOL CGRegistryCOMApp::InitInstance()
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

	CGRegistryCOMDlg dlg;
	m_pMainWnd = &dlg;

	// Đăng ký Registry bằng quyền Administrator
	RegisterGProperties();

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


CString CGRegistryCOMApp::GetWinDir()
{
	TCHAR windir[MAX_PATH];
	GetWindowsDirectory(windir, MAX_PATH);

	CString szSource = (CString)windir;
	if (szSource.Right(1) != L"\\") szSource += L"\\";

	return szSource;
}

CString CGRegistryCOMApp::GetPathDefault()
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

char* CGRegistryCOMApp::ConvertCstringToChars(CString szvalue)
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

void CGRegistryCOMApp::RegisterGProperties()
{
	try
	{
		CString szLibVer	= L"1.0";
		CString szVersion	= L"1.0.0.0";
		CString szFileName	= L"GProperties";
		CString szClassName	= L"CProperties";
		CString szPublicKey = L"68931570c5ca15ac";
		CString szCulture	= L"GXD JSC";

		CString szTkClass	= L"{95AC1ECB-199A-33E6-9091-E73BCD46BD67}";
		CString szTkIDPatch = L"{BF698AE6-B40B-3D0C-BE92-C4EBF4C821F3}";
		CString szTkDSOFile = L"{B607C6E4-F386-3473-9C63-404839A1894C}";
		CString szTkLIBID	= L"{F3752C3B-0014-46B7-B0EB-289275D94B06}";
		CString szTkQA		= L"{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}";

		char* chRuntimeVer	= "v4.0.30319";
		char* tokenProxy	= "{00020424-0000-0000-C000-000000000046}";

/****** Using HKEY_CURRENT_USER ************************************************************************/

		CString szpathWin = GetWinDir();
		CString szpathDefault = GetPathDefault();

		Base64Ex *bb = new Base64Ex;
		CRegistry *reg = new CRegistry;
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + EditProperties;
		reg->CreateKey(szCreate);
		reg->WriteChar("RegisterCOM", "1");
		delete bb;

/*****************************************************************************************************/

		char* tokenClass = ConvertCstringToChars(L"{" + szTkClass + L"}");
		char* tokenLIBID = ConvertCstringToChars(L"{" + szTkLIBID + L"}");
		char* chAssembly = ConvertCstringToChars(szFileName + L", Version=" + szVersion + L", Culture=" + szCulture + L", PublicKeyToken=" + szPublicKey);
		char* chLibVer = ConvertCstringToChars(szLibVer);

		CString szpath = szpathDefault;
		if (szpathDefault.Right(1) = L"\\") szpath = szpathDefault.Left(szpathDefault.GetLength() - 1);

		char* pathLibDIR = ConvertCstringToChars(szpath);
		char* pathLibTLB = ConvertCstringToChars(szpath + L"\\" + szFileName + L".tlb");

		szpath.Replace(L"\\", L"/");
		char* pathLibDLL = ConvertCstringToChars(L"file:///" + szpath + L"/" + szFileName + L".dll");
		char* chDescription = ConvertCstringToChars(L"Register for COM interop " + szFileName + L" by Leokiri");

		CString szFullClass = szFileName + L"." + szClassName;
		char* chClassName = ConvertCstringToChars(szFullClass);

/****** Using HKEY_CLASSES_ROOT ************************************************************************/

		CString szRegPath = L"";
		if (reg->SetRootKey(HKEY_CLASSES_ROOT))
		{
			reg->CreateKey(L"CLSID\\" + szTkClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(L"CLSID\\" + szTkClass + L"\\Implemented Categories\\" + szTkQA);

			reg->CreateKey(L"CLSID\\" + szTkClass + L"\\InprocServer32");

			reg->WriteChar("", "mscoree.dll");
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);
			reg->WriteChar("ThreadingModel", "Both");

			reg->CreateKey(L"CLSID\\" + szTkClass + L"\\InprocServer32\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

			reg->CreateKey(L"CLSID\\" + szTkClass + L"\\ProgId");
			reg->WriteChar("", chClassName);

/*******************************************************************************************************/

			reg->CreateKey(L"Interface\\" + szTkIDPatch);
			reg->WriteChar("", "IProperties");

			reg->CreateKey(L"Interface\\" + szTkIDPatch + L"\\ProxyStubClsid32");
			reg->WriteChar("", tokenProxy);

			reg->CreateKey(L"Interface\\" + szTkIDPatch + L"\\TypeLib");
			reg->WriteChar("", tokenLIBID);
			reg->WriteChar("Version", chLibVer);

			reg->CreateKey(szFullClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szFullClass + L"\\CLSID");
			reg->WriteChar("", tokenClass);

			reg->CreateKey(L"Record\\" + szTkDSOFile + L"\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", "DSOFile.dsoFileOpenOptions");
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

/*******************************************************************************************************/

			reg->CreateKey(L"TypeLib\\" + szTkLIBID + L"\\" + szLibVer);
			reg->WriteChar("", chDescription);

			reg->CreateKey(L"TypeLib\\" + szTkLIBID + L"\\" + szLibVer + L"\\0\\win32");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(L"TypeLib\\" + szTkLIBID + L"\\" + szLibVer + L"\\0\\win64");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(L"TypeLib\\" + szTkLIBID + L"\\" + szLibVer + L"\\FLAGS");
			reg->WriteChar("", "0");

			reg->CreateKey(L"TypeLib\\" + szTkLIBID + L"\\" + szLibVer + L"\\HELPDIR");
			reg->WriteChar("", pathLibDIR);

/*******************************************************************************************************/

			szRegPath = L"WOW6432Node\\Interface\\";

			reg->CreateKey(szRegPath + szTkClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szRegPath + szTkClass + L"\\Implemented Categories\\" + szTkQA);

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32");
			reg->WriteChar("", "mscoree.dll");
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);
			reg->WriteChar("ThreadingModel", "Both");

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

			reg->CreateKey(szRegPath + szTkClass + L"\\ProgId");
			reg->WriteChar("", chClassName);

/*******************************************************************************************************/

			szRegPath = L"WOW6432Node\\Interface\\";

			reg->CreateKey(szRegPath + szTkIDPatch);
			reg->WriteChar("", "IProperties");

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\ProxyStubClsid32");
			reg->WriteChar("", tokenProxy);

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\TypeLib");
			reg->WriteChar("", tokenLIBID);
			reg->WriteChar("Version", chLibVer);

/*******************************************************************************************************/

			szRegPath = L"WOW6432Node\\TypeLib\\";

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer);
			reg->WriteChar("", chDescription);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win32");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win64");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\FLAGS");
			reg->WriteChar("", "0");

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\HELPDIR");
			reg->WriteChar("", pathLibDIR);
		}

/****** Using HKEY_LOCAL_MACHINE ************************************************************************/

		if (reg->SetRootKey(HKEY_LOCAL_MACHINE))
		{
			szRegPath = L"SOFTWARE\\Classes\\CLSID\\";

			reg->CreateKey(szRegPath + szTkClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szRegPath + szTkClass + L"\\Implemented Categories\\" + szTkQA);

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32");
			reg->WriteChar("", "mscoree.dll");
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);
			reg->WriteChar("ThreadingModel", "Both");

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

			reg->CreateKey(szRegPath + szTkClass + L"\\ProgId");
			reg->WriteChar("", chClassName);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\Interface\\";

			reg->CreateKey(szRegPath + szTkIDPatch);
			reg->WriteChar("", "IProperties");

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\ProxyStubClsid32");
			reg->WriteChar("", tokenProxy);

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\TypeLib");
			reg->WriteChar("", tokenLIBID);
			reg->WriteChar("Version", chLibVer);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\";

			reg->CreateKey(szRegPath + szFullClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szRegPath + szFullClass + L"\\CLSID");
			reg->WriteChar("", tokenClass);

			reg->CreateKey(szRegPath + L"Record\\" + szTkDSOFile + L"\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", "DSOFile.dsoFileOpenOptions");
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\TypeLib\\";

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"");
			reg->WriteChar("", chDescription);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win32");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win64");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\FLAGS");
			reg->WriteChar("", "0");

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\HELPDIR");
			reg->WriteChar("", pathLibDIR);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\WOW6432Node\\CLSID\\";

			reg->CreateKey(szRegPath + szTkClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szRegPath + szTkClass + L"\\Implemented Categories\\" + szTkQA);

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32");
			reg->WriteChar("", "mscoree.dll");
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);
			reg->WriteChar("ThreadingModel", "Both");

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

			reg->CreateKey(szRegPath + szTkClass + L"\\ProgId");
			reg->WriteChar("", chClassName);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\WOW6432Node\\Interface\\";

			reg->CreateKey(szRegPath + szTkIDPatch);
			reg->WriteChar("", "IProperties");

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\ProxyStubClsid32");
			reg->WriteChar("", tokenProxy);

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\TypeLib");
			reg->WriteChar("", tokenLIBID);
			reg->WriteChar("Version", chLibVer);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\Classes\\WOW6432Node\\TypeLib\\";

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer);
			reg->WriteChar("", chDescription);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win32");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win64");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\FLAGS");
			reg->WriteChar("", "0");

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\HELPDIR");
			reg->WriteChar("", pathLibDIR);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\WOW6432Node\\Classes\\CLSID\\";

			reg->CreateKey(szRegPath + szTkClass);
			reg->WriteChar("", chClassName);

			reg->CreateKey(szRegPath + szTkClass + L"\\Implemented Categories\\" + szTkQA);

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32}\\InprocServer32");
			reg->WriteChar("", "mscoree.dll");
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);
			reg->WriteChar("ThreadingModel", "Both");

			reg->CreateKey(szRegPath + szTkClass + L"\\InprocServer32\\" + szVersion);
			reg->WriteChar("Assembly", chAssembly);
			reg->WriteChar("Class", chClassName);
			reg->WriteChar("CodeBase", pathLibDLL);
			reg->WriteChar("RuntimeVersion", chRuntimeVer);

			reg->CreateKey(szRegPath + szTkClass + L"\\ProgId");
			reg->WriteChar("", chClassName);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\WOW6432Node\\Classes\\Interface\\";

			reg->CreateKey(szRegPath + szTkIDPatch);
			reg->WriteChar("", "IProperties");

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\ProxyStubClsid32");
			reg->WriteChar("", tokenProxy);

			reg->CreateKey(szRegPath + szTkIDPatch + L"\\TypeLib");
			reg->WriteChar("Version", chLibVer);

/*******************************************************************************************************/

			szRegPath = L"SOFTWARE\\WOW6432Node\\Classes\\TypeLib\\";

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer);
			reg->WriteChar("", chDescription);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win32");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\0\\win64");
			reg->WriteChar("", pathLibTLB);

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\FLAGS");
			reg->WriteChar("", "0");

			reg->CreateKey(szRegPath + szTkLIBID + L"\\" + szLibVer + L"\\HELPDIR");
			reg->WriteChar("", pathLibDIR);
		}

		delete reg;
	}
	catch (...)
	{
		MessageBox(NULL, L"Lỗi đăng ký thư viện COM.", L"Thông báo", MB_OK | MB_ICONERROR);
	}
}

