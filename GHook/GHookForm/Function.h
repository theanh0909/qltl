#pragma once

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

void _msgbox(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
void _s(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
void _d(int num = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
void _n(CString message = L"Num", int num = 0, int itype = 0, int istr = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
int _y(CString message = L"", int istyle = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);

CString ConvertstringtoUTF8(string ret);
string ConvertUTF8tostring(CString szval);
char* ConvertCstringToChars(CString szvalue);
CString ConvertCharsToCstring(char* chr);
int IsTaskbarWndVisible(int &iTaskHeight);
int FileExistsUnicode(CString pathFile);
int CheckVersionSoftware();
int RandomTime();
CString GetPathFolder(CString fileName);
CString ConvertKhongDau(CString szNoidung);
CString GetTypeOfFile(CString pathFile = L"");
int TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1 = L";", CString symbol2 = L"|", int itypeFilter = 0, int itypeTrim = 1);
int CompareAZ(const void * a, const void * b);
CString NameColumnHeading(CListCtrl &clist, int column, int itype = 1, CString szName = L"");
void WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype = 0);
int ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);
bool FolderExistsUnicode(CString dirName_in);
BOOL CreateNewFolder(CString directoryPath);
bool CreateDirectories(CString pathName);
CString GetNameOfPath(CString pathFile, int &pos, int ipath = 0);
void ShellExecuteSelectFile(CString szDirFileSelect);
int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType = 0, int iStatus = 0);

bool _xlCreateExcel(bool blOpenFrm = false);
void _xlGetInfoSheetActive();
_bstr_t _xlGetNameSheet(_bstr_t codename, bool blError = false);
int _xlGetColumn(_WorksheetPtr pSheet, CString name, int itype);
int _xlGetRow(_WorksheetPtr pSheet, CString name, int itype);
CString _xlGetValue(_WorksheetPtr pSheet, CString name, int itype = 1, int inumberformat = 0);
void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS, CString pathFile, CString szName = L"Mở file",
	COLORREF clrBkg = xlNone, COLORREF clrTxt = rgbBlue, CString szFontName = L"Arial", int iSize = 12,
	bool bCenter = true, bool bFormula = true);
CString _xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError = L"");
int _xlFindComment(_WorksheetPtr pSheet, int column, int rowStart, int rowEnd, _bstr_t comment, int itype);
int _xlFindValue(_WorksheetPtr pSheet, int column, int rowbd, int rowkt, _bstr_t bvalue, int itype, bool match_content, bool match_case);
CString _xlGAR(RangePtr pRng, int nRow, int nColumn, int itype);
bool _xlGetPointWindowsOfCell(RangePtr pRng, int &pX, int &pY);

/**********************************************************************************/

bool IsCapsLockOn();
bool CheckKeyBoard(WPARAM wParam);
CString GetClassActiveWindow();
int KiemtraDieukienHook();
CString ConvertUnikey(CString szKey);

