#include "stdafx.h"
#include "Base64.h"
#include "Registry.h"
#include "Function.h"

Function::Function()
{
}

Function::~Function()
{
}

/************ Các hàm hay sử dụng ***************************************************/

char* Function::_ConvertCstringToChars(CString szvalue)
{
	wchar_t* wszString = new wchar_t[szvalue.GetLength() + 1];
	wcscpy(wszString, szvalue);
	int num = WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* chr = new char[num + 1];
	WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), chr, num, NULL, NULL);
	chr[num] = '\0';
	return chr;
}

CString Function::_ConvertCharsToCstring(char* chr)
{
	int num = MultiByteToWideChar(CP_UTF8, 0, chr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	MultiByteToWideChar(CP_UTF8, 0, chr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	return szval;
}

void Function::_SetPathDefault()
{
	if (spathDefault == L"") spathDefault = _GetPathFolder(FILENAMEDLL);
	if (spathDefault.Right(1) != L"\\") spathDefault += L"\\";
}

BOOL Function::_CheckConnection()
{
	Base64Ex bb;
	CString szBCheckConn = bb.BaseDecodeEx(CHECKCONN);
	if (InternetCheckConnection(szBCheckConn, FLAG_ICC_FORCE_CONNECTION, 0)) return TRUE;

	return FALSE;
}

int Function::_GetScaleLayoutWindow()
{
	try
	{
		Base64Ex bb;
		CRegistry reg;
		CString szCreate = bb.BaseDecodeEx(CreateKeyRun);
		reg.CreateKey(szCreate);

		int numScale = 100;
		CString szval = reg.ReadString(L"ScaleLayoutInWindows", L"");
		if (szval != L"")
		{
			numScale = _wtoi(szval);
			if (numScale < 100) numScale = 100;
			else if (numScale > 200) numScale = 200;
		}

		return numScale;
	}
	catch (...) {}
	return 100;
}

BOOL Function::_CreateNewFolder(CString directoryPath)
{
	CString directoryPathClone = directoryPath;
	LPTSTR lpPath = directoryPath.GetBuffer(MAX_PATH);
	if (!PathFileExists(lpPath))
	{
		if (!CreateDirectory(lpPath, NULL) && (GetLastError() == ERROR_PATH_NOT_FOUND))
		{
			if (PathRemoveFileSpec(lpPath))
			{
				CString newPath = lpPath;
				_CreateNewFolder(newPath);
			}
		}
	}

	directoryPath.ReleaseBuffer(-1);
	return PathFileExists(directoryPathClone);
}

CString Function::_GetPathFolder(CString szFile)
{
	try
	{
		HMODULE hModule;
		hModule = GetModuleHandle(szFile);
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
	catch (...)
	{
		CString szval = L"";
		szval.Format(L"Không tồn tại thư mục: %s.", szFile);
		_msgbox(szval, 2);
	}
	return L"";
}

int Function::_RandomTime()
{
	CString szval = L"";
	time_t now = time(0);
	tm *ltm = localtime(&now);
	szval.Format(L"%d%d", ltm->tm_min, ltm->tm_sec);
	int num = _wtoi(szval) + rand() % 1000;
	return num;
}

CString Function::_GetTimeNow(int ifulltime)
{
	time_t currentTime;
	time(&currentTime);
	struct tm *localTime = localtime(&currentTime);

	CString szday = L"", szmonth = L"";
	int iday = localTime->tm_mday;
	if (iday < 10) szday.Format(L"0%d", iday);
	else szday.Format(L"%d", iday);

	int imonth = 1 + (int)localTime->tm_mon;
	if (imonth < 10) szmonth.Format(L"0%d", imonth);
	else szmonth.Format(L"%d", imonth);

	CString szDaynow = L"";
	szDaynow.Format(L"%s/%s/%d", szday, szmonth, 1900 + (int)localTime->tm_year);

	if (ifulltime == 1)
	{
		CString szhour = L"", szmin = L"", szsec = L"";
		int ihour = localTime->tm_hour;
		if (ihour < 10) szhour.Format(L"0%d", ihour);
		else szhour.Format(L"%d", ihour);

		int imin = localTime->tm_min;
		if (imin < 10) szmin.Format(L"0%d", imin);
		else szmin.Format(L"%d", imin);

		int isec = localTime->tm_sec;
		if (isec < 10) szsec.Format(L"0%d", isec);
		else szsec.Format(L"%d", isec);

		CString sztimer = L"";
		sztimer.Format(L"%s:%s:%s", szhour, szmin, szsec);

		szDaynow += L" ";
		szDaynow += sztimer;
	}

	return szDaynow;
}

CString Function::_ConvertKhongDau(CString sKeyValue)
{
	CString scvtUTF[134] = {
		L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
		L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
		L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
		L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
		L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",
		L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
		L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",
		L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
		L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"đ",
		L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"Đ" };

	CString scvtKOD[134] = {
		L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",
		L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",
		L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",
		L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",
		L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",
		L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",
		L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",
		L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",
		L"i",L"i",L"i",L"i",L"i",L"y",L"y",L"y",L"y",L"y",L"d",
		L"I",L"I",L"I",L"I",L"I",L"Y",L"Y",L"Y",L"Y",L"Y",L"D" };

	for (int i = 0; i < 134; i++)
	{
		sKeyValue.Replace(scvtUTF[i], scvtKOD[i]);
	}

	return sKeyValue;
}

int Function::_FileExists(CString pathFile, int itype)
{
	CString szval = L"Đường dẫn không hợp lệ. Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		struct stat fileInfo;
		if (!stat(_ConvertCstringToChars(pathFile), &fileInfo)) return 1;
		if (itype == 1) _msgbox(szval, 2);
		return -1;
	}
	catch (...)
	{
		if (itype == 1) _msgbox(szval, 2);
		return -1;
	}
	return 0;
}

void Function::_WriteFileText(CString sNoidung, CString sPath, int itype)
{
	if (itype == 1)
	{
		// Kiểm tra xem tệp tin có tồn tại hay không
		if (_FileExists(sPath, 0) != 1) return;
	}

	ofstream myfile;
	myfile.open(_ConvertCstringToChars(sPath));
	myfile << _ConvertCstringToChars(sNoidung);
	myfile.close();
}

CString Function::_ReadFileText(CString sPath)
{
	if (_FileExists(sPath, 0) != 1) return L"";

	CString szNoidung = L"";
	wchar_t txtopen[MAX_LEN];
	FILE * fp = fopen(_ConvertCstringToChars(sPath), "r");
	fgetws(txtopen, MAX_LEN, fp);
	szNoidung.Format(L"%s", txtopen);
	fclose(fp);

	if (szNoidung != L"")
	{
		CString szChar = L"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
		CString szval = szNoidung.Left(1);
		szval.MakeUpper();
		
		int pos = szChar.Find(szval);
		if (pos == -1) szNoidung = L"";
	}

	return szNoidung;
}

int Function::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExists(spathFile, 0) != 1) return 0;

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

int Function::_TackTukhoa(vector<CString> &vecPkey, CString sKeyValue,
	CString symbol1, CString symbol2, int itype1, int itype2)
{
	try
	{
		vecPkey.clear();

		if (sKeyValue == L"") return 0;
		if (sKeyValue.Right(1) != symbol1 && sKeyValue.Right(1) != symbol2)
		{
			sKeyValue += symbol1;
		}

		int vtri = 0, icheck = 0;
		CString szval = L"", sktra = L"";
		for (int i = 0; i < sKeyValue.GetLength(); i++)
		{
			szval = sKeyValue.Mid(i, 1);
			if ((szval == symbol1 || szval == symbol2) && i >= vtri)
			{
				icheck = 0;
				if (i > vtri) sktra = sKeyValue.Mid(vtri, i - vtri);
				else sktra = L"";

				if (itype1 == 1)
				{
					if ((int)vecPkey.size() > 0)
					{
						for (int j = 0; j < (int)vecPkey.size(); j++)
						{
							if (sktra == vecPkey[j])
							{
								icheck = 1;
								break;
							}
						}
					}
				}

				if (icheck == 0)
				{
					if (itype2 == 0) sktra.Trim();
					vecPkey.push_back(sktra);
				}

				vtri = i + 1;
			}
		}

		return (int)vecPkey.size();
	}
	catch (...) {}
	return 0;
}


CString Function::_TackTenFile(CString spathURL, int itype)
{
	try
	{
		CString szval = L"";
		int ilen = spathURL.GetLength();
		for (int i = ilen - 1; i > 0; i--)
		{
			if (spathURL.Mid(i, 1) == L"\\" || spathURL.Mid(i, 1) == L"/")
			{
				if (itype == 1) return spathURL.Right(ilen - i - 1);
				else
				{
					spathURL = spathURL.Right(ilen - i - 1);
					ilen = spathURL.GetLength();
					for (int j = ilen - 1; j > 0; j--)
					{
						if (spathURL.Mid(j, 1) == L".")
						{
							if (itype == 0) return spathURL.Left(i);
							else spathURL.Right(ilen - i - 1);
						}
					}
				}
			}
		}

		return L"";
	}
	catch (...) {}
	return L"";
}

int Function::_FindKhongdau(CString szNoidung, CString szKey)
{
	// Chuyển có dấu sang không dấu & chuyển sang chữ thường
	szNoidung = _ConvertKhongDau(szNoidung);
	szKey = _ConvertKhongDau(szKey);
	szNoidung.MakeLower();
	szKey.MakeLower();

	vector<CString> vecLine;
	_TackTukhoa(vecLine, szKey, L" ", L"+", 1, 0);
	for (int i = 0; i < (int)vecLine.size(); i++)
	{
		if (szNoidung.Find(vecLine[i]) == -1) return 0;
	}
	return 1;
}

void Function::_CopyData(CString sNoidung)
{
	try
	{
		if (!OpenClipboard(NULL)) return;
		if (!EmptyClipboard()) return;

		size_t size = sizeof(TCHAR)*(1 + sNoidung.GetLength());
		HGLOBAL hResult = GlobalAlloc(GMEM_MOVEABLE, size);
		LPTSTR lptstrCopy = (LPTSTR)GlobalLock(hResult);
		memcpy(lptstrCopy, sNoidung.GetBuffer(), size);
		GlobalUnlock(hResult);
		#ifndef _UNICODE
			if (::SetClipboardData(CF_TEXT, hResult) == NULL)
		#else
			if (::SetClipboardData(CF_UNICODETEXT, hResult) == NULL)
		#endif
		{
			GlobalFree(hResult);
			CloseClipboard();
			return;
		}
		CloseClipboard();
	}
	catch (...) {}
}

CString Function::_PasteData()
{
	try
	{
		if (!OpenClipboard(NULL)) return L"";

		HANDLE hData;
		CString strData = L"";

		#ifndef _UNICODE
			hData = ::GetClipboardData(CF_TEXT);
		#else
			hData = ::GetClipboardData(CF_UNICODETEXT);
		#endif

		if (hData)
		{
			WCHAR *pchData = (WCHAR*)GlobalLock(hData);
			if (pchData)
			{
				strData = pchData;
				GlobalUnlock(hData);
			}
		}

		CloseClipboard();

		return strData;
	}
	catch (...) {}
	return L"";
}


/************ Các hàm thông báo MessageBox ******************************************/

int Function::MessageBoxTimeoutA(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption, 
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds)
{
	static MSGBOXAAPI MsgBoxTOA = NULL;
	if (!MsgBoxTOA)
	{
		HMODULE hUser32 = GetModuleHandle(_T("user32.dll"));
		if (hUser32) MsgBoxTOA = (MSGBOXAAPI)GetProcAddress(hUser32, "MessageBoxTimeoutA");
		else return 0;
	}
	if (MsgBoxTOA) return MsgBoxTOA(hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
	return 0;
}

int Function::MessageBoxTimeoutW(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption, 
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds)
{
	static MSGBOXWAPI MsgBoxTOW = NULL;
	if (!MsgBoxTOW)
	{
		HMODULE hUser32 = GetModuleHandle(_T("user32.dll"));
		if (hUser32) MsgBoxTOW = (MSGBOXWAPI)GetProcAddress(hUser32, "MessageBoxTimeoutW");
		else return 0;
	}
	if (MsgBoxTOW) return MsgBoxTOW(hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
	return 0;
}

void Function::_msgbox(CString message, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
	MessageBoxTimeout(NULL, message, THONGBAO, style[itype], 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
}

void Function::_s(CString message, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
	MessageBoxTimeout(NULL, message, THONGBAO, style[itype], 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
}

void Function::_d(int num, int itype, int iexcel, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, iexcel, itimeout);
}

void Function::_n(CString message, int num, int itype, int itype2, int iexcel, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (itype2 == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, iexcel, itimeout);
}


int Function::_y(CString message, int istyle, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT qs[2] = { MB_YESNO, MB_YESNOCANCEL };
	UINT style[2] = { MB_ICONQUESTION, MB_ICONWARNING };

	return (int)MessageBoxTimeout(NULL, message, THONGBAO,
		qs[istyle] | style[itype], 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);

	return 6;
}

/************ Các hàm sử dụng trong Excel *******************************************/

void Function::_xlCreateExcel(int itype)
{
	CoInitialize(NULL);
	xl.GetActiveObject(ExcelApp);
	xl->PutVisible(VARIANT_TRUE);
	if (itype == 1) xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	pWb = xl->GetActiveWorkbook();
}

void Function::_xlCloseExcel()
{
	CoUninitialize();
}

/************ Các hàm đặc biệt khác *************************************************/

