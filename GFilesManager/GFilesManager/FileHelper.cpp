#include "pch.h"
#include "FileHelper.h"
#include "FileRegister.h"
#include "Shlwapi.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

unsigned long CFileHelper::GetFileSize(CString& file)
{
	const char * filename = (const char *)file.GetBuffer();
	FILE *pFile = fopen(filename, "rb");
	fseek(pFile, 0, SEEK_END);
	unsigned long size = ftell(pFile);
	fclose(pFile);
	file.ReleaseBuffer();
	return size;
}
BOOL CFileHelper::IsFileExist(CString& file)
{
	DWORD dwAttricutes = ::GetFileAttributes(file);
	return dwAttricutes != DWORD(-1) && ((dwAttricutes & FILE_ATTRIBUTE_DIRECTORY) == 0);
}

BOOL CFileHelper::IsFileOpen(CString& filePath)
{
	//Open In Exclusive mode will fail if file already opened
	HANDLE hFile = CreateFile(filePath, GENERIC_READ, 0 /*No Share*/
		, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == hFile)
		return TRUE;
	else
		::CloseHandle(hFile);
	return FALSE;
}
BOOL CFileHelper::IsDirExist(CString& dir)
{
	DWORD dwAttricutes = ::GetFileAttributes(dir);
	return (dwAttricutes != DWORD(-1)) && (dwAttricutes & FILE_ATTRIBUTE_DIRECTORY);
}

BOOL CFileHelper::IsShortcut(CString& fileName)
{
	//DWORD dwAttricutes = ::GetFileAttributes(fileName);
	//if(dwAttricutes & FILE_ATTRIBUTE_REPARSE_POINT)
	//{
	//	WIN32_FIND_DATA find ;
	//	HANDLE hFile = FindFirstFile(fileName, &find);
	//	{
	//		if(find.dwReserved0 & IO_REPARSE_TAG_SYMLINK)
	//			return TRUE;
	//	}
	//}
	//return FALSE;
	CString tmp = fileName.Right(4).MakeUpper();
	if (tmp.Compare(_T(".LNK")) == 0)
		return TRUE;
	return FALSE;
}

BOOL CFileHelper::CreateDirIfNeeded(CString& path)
{
	BOOL ret = ::CreateDirectory(path, NULL);
	if (!ret)
		return GetLastError() == ERROR_ALREADY_EXISTS ? TRUE : FALSE;
	return ret;
}


CString CFileHelper::ResolveShortcut(CString& szShortcut)
{
	// start by checking its a valid file
	if (!IsFileExist(szShortcut))
		return _T("");

	CoInitialize(NULL);

	HRESULT hResult;
	IShellLink*	psl;
	CString sTarget(szShortcut);

	hResult = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
		IID_IShellLink, (LPVOID*)&psl);

	if (SUCCEEDED(hResult))
	{
		LPPERSISTFILE ppf;

		hResult = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf);

		if (SUCCEEDED(hResult))
		{
			hResult = ppf->Load(szShortcut, STGM_READ);
			if (SUCCEEDED(hResult))
			{
				hResult = psl->Resolve(NULL, SLR_ANY_MATCH | SLR_NO_UI);
				if (SUCCEEDED(hResult))
				{
					TCHAR szPath[MAX_PATH];
					WIN32_FIND_DATA wfd;
					_tcscpy_s(szPath, szShortcut);
					hResult = psl->GetPath(szPath, MAX_PATH, (WIN32_FIND_DATA*)&wfd, SLGP_SHORTPATH);
					if (SUCCEEDED(hResult))
						sTarget = CString(szPath);
				}
			}
			ppf->Release();
		}
		psl->Release();
	}
	CoUninitialize();
	return sTarget;
}
CString CFileHelper::GetAppFileName()
{
	return GetModuleFileName(NULL);
}

CString CFileHelper::GetModuleFileName(HMODULE hMod)
{
	CString sModulePath;
	::GetModuleFileName(hMod, sModulePath.GetBuffer(MAX_PATH + 1), MAX_PATH);
	sModulePath.ReleaseBuffer();
	return sModulePath;
}

CString CFileHelper::GetAppFolder()
{
	CSplitPath path(GetAppFileName());
	return path.GetFullPath();
}
DWORD CFileHelper::Run(HWND hwnd, LPCTSTR lpFile, LPCTSTR lpDirectory, int nShowCmd)
{
	CString sFile(lpFile), sParams;
	int nHash = sFile.Find('#');

	if (nHash != -1)
	{
		sParams = sFile.Mid(nHash);
		sFile = sFile.Left(nHash);
		CSplitPath path(sFile);
		CString sExt = path.GetExt();
		CString sApp = CFileRegister::GetRegisteredAppPath(sExt);

		if (!sApp.IsEmpty())
		{
			sFile = sApp;
			sParams = lpFile;
		}
		else
		{
			sFile = lpFile;
			sParams.Empty();
		}
	}

	DWORD dwRes = (DWORD)ShellExecute(hwnd, NULL, sFile, sParams, lpDirectory, nShowCmd);

	if (dwRes <= 32) // failure
	{
		// try CreateProcess
		STARTUPINFO si;
		PROCESS_INFORMATION pi;

		ZeroMemory(&si, sizeof(si));
		si.cb = sizeof(si);
		si.wShowWindow = (WORD)nShowCmd;
		si.dwFlags = STARTF_USESHOWWINDOW;

		ZeroMemory(&pi, sizeof(pi));

		// Start the child process.
		if (CreateProcess(NULL,			// No module name (use command line).
			(LPTSTR)lpFile,	// Command line.
			NULL,			// Process handle not inheritable.
			NULL,			// Thread handle not inheritable.
			FALSE,			// Set handle inheritance to FALSE.
			0,				// No creation flags.
			NULL,			// Use parent's environment block.
			lpDirectory,	// starting directory.
			&si,			// Pointer to STARTUPINFO structure.
			&pi))			// Pointer to PROCESS_INFORMATION structure.
		{
			dwRes = 32; // success
		}

		// Close process and thread handles.
		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);
	}
	return dwRes;
}

CString CFileHelper::GetCurDir()
{
	CString sFolder;
	::GetCurrentDirectory(MAX_PATH, sFolder.GetBuffer(MAX_PATH + 1));
	sFolder.ReleaseBuffer();
	return sFolder;
}

CString CFileHelper::GetRelativePath(const CString& sFilePath, CString& sRelativeToFolder, BOOL bFolder)
{
	if (::PathIsRelative(sFilePath))
		return sFilePath;
	TCHAR szRelPath[MAX_PATH + 1];
	BOOL bRes = ::PathRelativePathTo(szRelPath,
		UnterminatePath(sRelativeToFolder), FILE_ATTRIBUTE_DIRECTORY,
		sFilePath, bFolder ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL);
	return bRes ? CString(szRelPath) : sFilePath;
}

CString& CFileHelper::TerminatePath(CString& sPath)
{
	sPath.TrimRight();
	if (sPath.ReverseFind('\\') != (sPath.GetLength() - 1))
		sPath += '\\';
	return sPath;
}
CString& CFileHelper::UnterminatePath(CString& sPath)
{
	sPath.TrimRight();
	int len = sPath.GetLength();
	if (sPath.ReverseFind('\\') == (len - 1))
		sPath = sPath.Left(len - 1);
	return sPath;
}
CString CFileHelper::GetFullPath(const CString& sFilePath, const CString& sRelativeToFolder)
{
	if (!::PathIsRelative(sFilePath))
	{
		// special case: filename has no drive letter and is not UNC
		if (sFilePath.Find(_T(":\\")) == -1 && sFilePath.Find(_T("\\\\")) == -1)
		{
			CString sDrive;
			CSplitPath path(sRelativeToFolder);
			sDrive = path.GetDrive();
			TCHAR sFullPath[MAX_PATH];
			_tmakepath_s(sFullPath, MAX_PATH, sDrive, NULL, sFilePath, NULL);
			return sFullPath;
		}
		// else
		return sFilePath;
	}

	CString sSrcPath(sRelativeToFolder);
	TerminatePath(sSrcPath);
	sSrcPath += sFilePath;

	TCHAR szFullPath[MAX_PATH + 1];
	BOOL bRes = ::PathCanonicalize(szFullPath, sSrcPath);
	return bRes ? CString(szFullPath) : sSrcPath;
}

/****************************************************************************/
CSplitPath::CSplitPath(LPCTSTR lpszPathBuffer/*=NULL*/)
{
	if (lpszPathBuffer != NULL)
	{
		m_path = lpszPathBuffer;
		SplitPath(lpszPathBuffer);
	}
}

CSplitPath::~CSplitPath()
{
}

void CSplitPath::SplitPath(LPCTSTR lpszPathBuffer)
{
	_tsplitpath_s(lpszPathBuffer, m_szDrive, m_szDrive ? _MAX_DRIVE : 0,
		m_szDir, m_szDir ? _MAX_DIR : 0, m_szFName,
		m_szFName ? _MAX_FNAME : 0, m_szExt, m_szExt ? _MAX_EXT : 0);
}

CString CSplitPath::GetDrive() const
{
	return CString(m_szDrive);
}
CString CSplitPath::GetDir() const
{
	return CString(m_szDir);
}
CString CSplitPath::GetFName() const
{
	return CString(m_szFName);
}
CString CSplitPath::GetExt() const
{
	return CString(m_szExt);
}
CString CSplitPath::GetFullPath() const
{
	return GetDrive() + GetDir();
}
CString CSplitPath::GetFullName() const
{
	return GetFName() + GetExt();
}

CString CSplitPath::MakeDBPath() const
{
	CString str = _T("DB:>\\");
	str += GetFullName();
	return str;
}

void CSplitPath::SuggestFileName(CString& str)
{
	TCHAR unique[MAX_PATH];
	PathYetAnotherMakeUniqueName(unique, str.GetBuffer(), NULL, NULL);
	str.ReleaseBuffer();
	str = unique;
}

BOOL CSplitPath::IsRemotePath(BOOL bAllowFileCheck)
{
	if (bAllowFileCheck)
	{
		DWORD dwAttr = ::GetFileAttributes(m_path);

		if (dwAttr == 0xffffffff)
			return -1;
	}

	return (IsUNCPath() || IsMappedPath());
}

BOOL CSplitPath::IsDBPath() const
{
	return m_path.Left(5).Compare(_T("DB:>\\")) == 0;
}

BOOL CSplitPath::IsUNCPath()
{
	return m_path.Left(2).Compare(_T("\\\\")) == 0;
}

BOOL CSplitPath::IsMappedPath()
{
	return ::GetDriveType(m_path) == DRIVE_REMOTE;
}
BOOL CSplitPath::IsFixedPath()
{
	return ::GetDriveType(m_path) == DRIVE_FIXED;
}
BOOL CSplitPath::IsReadonlyPath()
{
	DWORD dwAttr = ::GetFileAttributes(m_path);
	if (dwAttr == 0xffffffff)
		return -1;
	return (dwAttr & FILE_ATTRIBUTE_READONLY);
}