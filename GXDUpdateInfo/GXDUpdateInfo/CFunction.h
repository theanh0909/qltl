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

class CFunction
{

public:
	CFunction();
	~CFunction();

	CRegistry* reg;

	/*************** Message box ************************************/

	// itype: 0 = MB_ICONINFORMATION | 1 = MB_ICONWARNING | 2 = MB_ICONERROR
	void _msgbox(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _s(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _d(int num = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _n(CString message = L"Num", int num = 0, int itype = 0, int istr = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	int _y(CString message = L"", int istyle = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);

	/***********************************************************************************************/

	// Hàm khởi tạo Excel
	// blOpenFrm = true --> Trong hàm được sử dụng có mở Form
	bool _xlCreateExcel(bool blOpenFrm = false);

	// Hàm gọi hàm từ DLL khác
	// szDllName = Tên DLL sử dụng
	// szFunction = Tên hàm trong DLL được gọi ra sử dụng
	// blQuit = true --> Đóng DLL sau khi kết thúc sử dụng hàm
	void _xlCallFunctionDLL(CString szDllName, CString szFunction, bool blQuit = true);


	int _GetScaleLayoutWindow();

};