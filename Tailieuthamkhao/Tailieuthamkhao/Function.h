#pragma once
class Function
{
public:
	Function();
	~Function();
	

/************ Các hàm hay sử dụng *******************************************************/
	
	char * _ConvertCstringToChars(CString szvalue);
	CString _ConvertCharsToCstring(char * chr);

	void _SetPathDefault();
	BOOL _CheckConnection();
	BOOL _CreateNewFolder(CString directoryPath);
	CString _GetPathFolder(CString szFile);

	int _RandomTime();
	CString _GetTimeNow(int ifulltime = 0);
	CString _ConvertKhongDau(CString sKeyValue);
	int _FileExists(CString pathFile, int itype = 1);
	void _WriteFileText(CString sNoidung, CString sPath, int itype = 1);
	CString _ReadFileText(CString sPath);
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);

	int _TackTukhoa(vector<CString> &vecPkey, CString sKeyValue,
		CString symbol1 = L";", CString symbol2 = L"|", int itype1 = 0, int itype2 = 1);

	CString _TackTenFile(CString spathURL, int itype = 1);

	int _GetScaleLayoutWindow();

	int _FindKhongdau(CString szNoidung, CString szKey);

	void _CopyData(CString sNoidung = L"");

	CString _PasteData();

/************ Các hàm thông báo MessageBox **********************************************/

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

	void _msgbox(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _s(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _d(int num = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _n(CString message = L"Num", int num = 0, int itype = 0, int itype2 = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	int _y(CString message = L"", int istyle = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);

/************ Các hàm sử dụng trong Excel *******************************************/

	void _xlCreateExcel(int itype = 0);
	void _xlCloseExcel();	

/************ Các hàm đặc biệt khác *************************************************/





};

