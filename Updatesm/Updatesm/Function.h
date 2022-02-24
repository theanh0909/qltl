#pragma once
#include "pch.h"

typedef int(__stdcall *MSGBOXAAPI)(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

typedef int(__stdcall *MSGBOXWAPI)(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

#ifdef UNICODE
#define MessageBoxTimeout  MessageBoxTimeoutW
#else
#define MessageBoxTimeout  MessageBoxTimeoutA
#endif

int MessageBoxTimeoutA(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

int MessageBoxTimeoutW(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

typedef struct file_info_struct
{
	DWORD    dwFileAttributes;
	FILETIME ftCreationTime;
	FILETIME ftLastAccessTime;
	FILETIME ftLastWriteTime;
	DWORD    nFileSizeHigh;
	DWORD    nFileSizeLow;
} FILE_INFO;

typedef BOOL(WINAPI *GET_FILE_ATTRIBUTES_EX)(LPCWSTR lpFileName, int fInfoLevelId, LPVOID lpFileInformation);

class Function
{

public:
	Function();
	~Function();

/*************** Message box ************************************/
	// itype: 0 = MB_ICONINFORMATION | 1 = MB_ICONWARNING | 2 = MB_ICONERROR
	void _msgbox(CString message = L"", int itype = 0, int itimeout = MINISEC_TIME_OUT);
	void _s(CString message = L"", int itype = 0, int itimeout = MINISEC_TIME_OUT);
	void _d(int num = 0, int itype = 0, int itimeout = MINISEC_TIME_OUT);
	void _n(CString message = L"Num", int num = 0, int itype = 0, int istr = 0, int itimeout = MINISEC_TIME_OUT);
	int _y(CString message = L"", int istyle = 0, int itype = 0, int itimeout = MINISEC_TIME_OUT);

/*************** Functions ****************************************/
	
	int _CheckVersionSoftware();

	CString _GetDirModules(DWORD processID, bool blFilter = false);
	bool _IsProcessRunning(const wchar_t *processName);

	void _LogFileWrite(CString szLog = L"", bool blClear = false, bool blOpenFile = false, CString szFileName = L"FileLog.txt");
	void _LogFileStart(CString szStart = L"<!>");
	void _LogFileEnd();

	CString _GetDesktopDir();
	CString _GetNameOfPath(CString pathFile, int ipath = 0);
	CString _GetPathFolder(CString fileName);

	char* _ConvertCstringToChars(CString szvalue);
	CString _ConvertCharsToCstring(char* chr);
	
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);
	int _FileExistsUnicode(CString pathFile);
};
