#include "pch.h"
#include "Function.h"

int MessageBoxTimeoutA(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
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

int MessageBoxTimeoutW(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
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

void _msgbox(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);

		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void _s(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);

		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void _d(int num, int itype, int iexcel, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, iexcel, itimeout);
}

void _n(CString message, int num, int itype, int istr, int iexcel, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (istr == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, iexcel, itimeout);
}


int _y(CString message, int istyle, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT qs[2] = { MB_YESNO, MB_YESNOCANCEL };
	UINT style[2] = { MB_ICONQUESTION, MB_ICONWARNING };

	return (int)MessageBoxTimeout(NULL, message, MS_THONGBAO,
		qs[istyle] | style[itype], 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);

	return 6;
}

CString ConvertstringtoUTF8(string ret)
{
	char *cstr = new char[ret.length() + 1];
	strcpy(cstr, ret.c_str());
	int num = ::MultiByteToWideChar(CP_UTF8, 0, cstr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	::MultiByteToWideChar(CP_UTF8, 0, cstr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	return szval;
}

string ConvertUTF8tostring(CString szval)
{
	wchar_t* wszString = new wchar_t[szval.GetLength() + 1];
	wcscpy(wszString, szval);
	int num = ::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* utf8 = new char[num + 1];
	::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), utf8, num, NULL, NULL);
	utf8[num] = '\0';
	string str(utf8);
	return str;
}

char* ConvertCstringToChars(CString szvalue)
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

CString ConvertCharsToCstring(char* chr)
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

int IsTaskbarWndVisible(int &iTaskHeight)
{
	// nPostion = -1 --> Thanh Taskbar bị ẩn
	// nPostion =  0 - Bottom | 1 - Left | 2 - Top | 3 - Right
	int nPostion = -1;
	HWND hTaskbarWnd = ::FindWindow(L"Shell_TrayWnd", NULL);
	HMONITOR hMonitor = MonitorFromWindow(hTaskbarWnd, MONITOR_DEFAULTTONEAREST);
	MONITORINFO info = { sizeof(MONITORINFO) };
	if (GetMonitorInfo(hMonitor, &info))
	{
		RECT rect;
		GetWindowRect(hTaskbarWnd, &rect);

		if (!((rect.top >= info.rcMonitor.bottom - 4) ||
			(rect.right <= 2) ||
			(rect.bottom <= 4) ||
			(rect.left >= info.rcMonitor.right - 2)))
		{
			if ((int)rect.left > 0)
			{
				// Task --> Pane Right
				nPostion = 3;
				iTaskHeight = (int)rect.left;
			}
			else
			{
				if ((int)rect.top > 0)
				{
					// Task --> Pane Bottom
					nPostion = 0;
					iTaskHeight = (int)rect.top;
				}
				else
				{
					if ((int)rect.right == 1366)
					{
						// Task --> Pane Top
						nPostion = 2;
						iTaskHeight = (int)rect.bottom;
					}
					else
					{
						// Task --> Pane Left
						nPostion = 1;
						iTaskHeight = (int)rect.right;
					}
				}
			}
		}
	}
	return nPostion;
}

int FileExistsUnicode(CString pathFile)
{
	LPTSTR lpPath = pathFile.GetBuffer(MAX_PATH);
	if (PathFileExists(lpPath)) return 1;
	return 0;
}

int CheckVersionSoftware()
{
	return _wtoi(VERSION_SOFT);
}

int RandomTime()
{
	time_t now = time(0);
	tm *ltm = localtime(&now);
	return ((100 * (int)ltm->tm_min + (int)ltm->tm_sec + rand()) % 1000);
}

void SetUseKeyPress()
{
	try
	{
		CRegistry *reg = new CRegistry;
		CString szCreate = BaseDecodeEx(CreateKeySettings);
		reg->CreateKey(szCreate);
		reg->WriteChar(ConvertCstringToChars(UseKeyPress), "1");
		delete reg;
	}
	catch (...) {}
}

bool GetUseKeyPress()
{
	try
	{
		CRegistry *reg = new CRegistry;

		bool blCheck = false;
		CString szCreate = BaseDecodeEx(CreateKeySettings);
		reg->CreateKey(szCreate);

		CString szBase = reg->ReadString(UseKeyPress, L"");
		if (_wtoi(szBase) == 1)
		{
			reg->DeleteValue(UseKeyPress);
			blCheck = true;
		}

		delete reg;
		return blCheck;
	}
	catch (...) {}
	return false;
}

CString GetPathFolder(CString fileName)
{
	try
	{
		_bstr_t bsConfig = (LPCTSTR)_xlGetNameSheet(shConfig, 0);
		if (bsConfig != (_bstr_t)L"")
		{
			_WorksheetPtr pShConfig = xl->Sheets->GetItem((_bstr_t)bsConfig);
			return _xlGetValue(pShConfig, L"CF_DIR_DEFAULT", 0, 0);
		}

		HMODULE hModule;
		hModule = GetModuleHandle(fileName);
		TCHAR szFileName[MAX_PATH];
		DWORD dSize = MAX_PATH;
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

CString ConvertKhongDau(CString szNoidung)
{
	CString scvtUTF[] = {
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

	CString scvtKOD[] = {
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

	int nsizearr = getSizeArray(scvtUTF);
	for (int i = 0; i < nsizearr; i++)
	{
		szNoidung.Replace(scvtUTF[i], scvtKOD[i]);
	}

	return szNoidung;
}

CString GetTypeOfFile(CString pathFile)
{
	try
	{
		if (pathFile != L"")
		{
			if ((int)pathFile.Find(L".") > 0)
			{
				CString szval = L"";
				int nLen = pathFile.GetLength();
				for (int i = nLen - 1; i >= 0; i--)
				{
					szval = pathFile.Mid(i, 1);
					if (szval == L".") return pathFile.Right(nLen - i - 1);
					else if (szval == L"\\" || szval == L"/") return L"";
				}
			}
		}
	}
	catch (...) {}
	return L"";
}

int TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1, CString symbol2, int itypeFilter, int itypeTrim)
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

				if (itypeFilter == 1)
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
					if (itypeTrim == 0) sktra.Trim();
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

int CompareAZ(const void * a, const void * b)
{
	if (*(CString*)a < *(CString*)b) return -1;
	else if (*(CString*)a > *(CString*)b) return  1;
	return 0;
}

CString NameColumnHeading(CListCtrl &clist, int column, int itype, CString szName)
{
	try
	{
		CString strNome = L"";
		if (itype == 1) strNome = szName;

		CHeaderCtrl *pHeader = clist.GetHeaderCtrl();
		HDITEM hdi;
		hdi.mask = HDI_TEXT;
		hdi.pszText = strNome.GetBuffer(256);
		hdi.cchTextMax = 256;

		if (itype == 0) pHeader->GetItem(column, &hdi);
		else pHeader->SetItem(column, &hdi);
		strNome.ReleaseBuffer();

		return strNome;
	}
	catch (...) {}
	return L"";
}

void WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (FileExistsUnicode(spathFile) != 1) return;
		}

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(ConvertCstringToChars(spathFile));
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

int ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (FileExistsUnicode(spathFile) != 1) return 0;

		wifstream readfile_(ConvertCstringToChars(spathFile));
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

bool FolderExistsUnicode(CString dirName_in)
{
	DWORD dwAttrib = GetFileAttributes((LPCTSTR)dirName_in);
	return (dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

BOOL CreateNewFolder(CString directoryPath)
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
				CreateNewFolder(newPath);
			}
		}
	}

	directoryPath.ReleaseBuffer(-1);
	return PathFileExists(directoryPathClone);
}

bool CreateDirectories(CString pathName)
{
	try
	{
		if (pathName.Right(1) != L"\\") pathName += L"\\";
		for (int i = 0; i < pathName.GetLength(); i++)
		{
			if (pathName.Mid(i, 1) == L"\\") CreateNewFolder(pathName.Left(i));
		}
		return true;
	}
	catch (...) {}
	return false;
}



CString GetNameOfPath(CString pathFile, int &pos, int ipath)
{
	try
	{
		pos = -1;
		if (pathFile != L"")
		{
			if (pathFile.Right(1) == L"\\") pathFile = pathFile.Left(pathFile.GetLength() - 1);

			int nLen = pathFile.GetLength();
			for (int i = nLen - 1; i >= 0; i--)
			{
				if (pathFile.Mid(i, 1) == L"\\" || pathFile.Mid(i, 1) == L"/")
				{
					pos = i;	// <-- Trả về vị trí phân tách

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


void ShellExecuteSelectFile(CString szDirFileSelect)
{
	SHELLEXECUTEINFO shExecInfo = { 0 };
	shExecInfo.cbSize = sizeof(shExecInfo);
	shExecInfo.lpFile = _T("Explorer.exe");
	shExecInfo.lpParameters = L"/Select, " + szDirFileSelect;
	shExecInfo.nShow = SW_SHOWNORMAL;
	shExecInfo.lpVerb = _T("Open");
	shExecInfo.fMask = SEE_MASK_INVOKEIDLIST | SEE_MASK_FLAG_DDEWAIT | SEE_MASK_FLAG_NO_UI;
	VERIFY(ShellExecuteEx(&shExecInfo));
}


int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType, int iStatus)
{
	try
	{
		vecHyper.clear();
		int ncount = (int)pRgSelect->Hyperlinks->GetCount();
		if (ncount > 0)
		{
			CString szval = L"", szlink = L"";
			CString szUpdateStatus = L"Đang tiến hành lấy thông tin dữ liệu. Vui lòng đợi trong giây lát...";

			for (int i = 1; i <= ncount; i++)
			{
				szlink = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(i)->GetAddress();

				if (iType == 1) szlink.Replace(L"/", L"\\");
				else if (iType == 2) szlink.Replace(L"\\", L"/");

				if (iStatus == 1)
				{
					szval.Format(L"%s (%d%s) %s", szUpdateStatus,
						(int)(100 * i / ncount), L"%", szlink);
					if (szval.GetLength() > 250) szval = szval.Left(250) + L"...";
					xl->PutStatusBar((_bstr_t)szval);
				}

				vecHyper.push_back(szlink);
			}

			if (iStatus == 1) xl->PutStatusBar((_bstr_t)L"Ready");
		}
		return (int)vecHyper.size();
	}
	catch (...) {}
	return 0;
}

bool _xlCreateExcel(bool blOpenFrm)
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(ExcelApplication);
		xl->PutVisible(VARIANT_TRUE);
		if (blOpenFrm) xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	}
	catch (...) {
		_msgbox(L"Lỗi khởi tạo ứng dụng Microsoft Excel. "
			L"Nguyên nhân có thể do có Excel chạy ngầm, "
			L"bạn vào TaskManager để tắt đi sau đó khởi động lại phần mềm.", 2);
		return false;
	}

	return true;
}

void _xlGetInfoSheetActive()
{
	_WorkbookPtr pWbObject = xl->GetActiveWorkbook();

	pShActive = pWbObject->ActiveSheet;
	pRgActive = pShActive->Cells;

	sCodeActive = (LPCTSTR)pShActive->CodeName;
	sNameActive = (LPCTSTR)pShActive->Name;

	iRowActive = pShActive->Application->ActiveCell->GetRow();
	iColumnActive = pShActive->Application->ActiveCell->GetColumn();
}


_bstr_t _xlGetNameSheet(_bstr_t codename, bool blError)
{
	try
	{
		_WorksheetPtr psh;
		_WorkbookPtr pWbObject = xl->GetActiveWorkbook();
		int ncountSheet = pWbObject->Sheets->Count;
		for (int i = 1; i <= ncountSheet; i++)
		{
			psh = pWbObject->Worksheets->Item[i];
			if (psh->CodeName == codename) return psh->Name;
		}
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy bảng tính có codename '%s'. Vui lòng kiểm tra lại.", (LPCTSTR)codename);
		if (blError) _msgbox(str, 2);
	}
	return L"";
}


int _xlGetColumn(_WorksheetPtr pSheet, CString name, int itype)
{
	try
	{
		RangePtr pCell = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		return (int)pCell->Cells->Column;
	}
	catch (...){}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn() + 1;
}

int _xlGetRow(_WorksheetPtr pSheet, CString name, int itype)
{
	try
	{
		RangePtr pCell = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		return (int)pCell->Cells->Row;
	}
	catch (...){}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
}

CString _xlGetValue(_WorksheetPtr pSheet, CString name, int itype, int inumberformat)
{
	try
	{
		if (inumberformat == 1)
		{
			RangePtr pCell = pSheet->Cells;
			return _xlGIOR(pCell, _xlGetRow(pSheet, name, 1), _xlGetColumn(pSheet, name, 0), L"");
		}
		else return (LPCTSTR)(_bstr_t)pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name)->GetValue2();
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
		if (itype == 1) _msgbox(str, 1);
	}

	return L"";
}

void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS, CString pathFile, CString szName,
	COLORREF clrBkg, COLORREF clrTxt, CString szFontName, int iSize, bool bCenter, bool bFormula)
{
	try
	{
		pSheet->Hyperlinks->Add(PRS, (_bstr_t)pathFile);
		PRS->PutValue2((_bstr_t)szName);
		PRS->Interior->PutColor(clrBkg);
		PRS->Font->PutColor(clrTxt);
		PRS->Font->PutName((_bstr_t)szFontName);
		PRS->Font->PutSize(iSize);
		if (bCenter) PRS->PutHorizontalAlignment(xlCenter);
	}
	catch (...) {}
}

CString _xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError)
{
	try
	{
		_bstr_t bstr = pRange->GetItem(nRow, nColumn);
		return (LPCTSTR)bstr;
	}
	catch (...) { return szError; }
}

int _xlFindComment(_WorksheetPtr pSheet, int column, int rowStart, int rowEnd, _bstr_t comment, int itype)
{
	try
	{
		int iRowComment = 0;
		CString szval = L"";		
		RangePtr PRS, pRange = pSheet->Cells;

		if (rowStart <= 0) rowStart = 1;
		if (rowEnd <= 0) rowEnd = 1000000;
		szval.Format(L"%s:%s", _xlGAR(pRange, rowStart, column, 0), _xlGAR(pRange, rowEnd, column, 0));
		if (PRS = pRange->GetRange(COleVariant(szval))->Find(comment, vtMissing, XlFindLookIn::xlComments, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false))
		{
			iRowComment = PRS->Cells->Row;
		}

		pRange->GetRange(COleVariant(L"A1:A1"))->Find("", vtMissing, XlFindLookIn::xlFormulas, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false);

		if (iRowComment > 0) return iRowComment;
	}
	catch (...)
	{
	}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
}


int _xlFindValue(_WorksheetPtr pSheet, int column, int rowbd, int rowkt,
	_bstr_t bvalue, int itype, bool match_content, bool match_case)
{
	try
	{
		int iRowValue = 0;
		CString szval = L"A1:AZ1000000";

		RangePtr PRS;
		RangePtr pRange = pSheet->Cells;

		if (rowbd <= 0) rowbd = 1;
		if (rowkt <= 0) rowkt = 1000000;
		if (column > 0) szval.Format(L"%s:%s", _xlGAR(pRange, rowbd, column, 0), _xlGAR(pRange, rowkt, column, 0));

		_variant_t vtCon = XlLookAt::xlPart;
		if (match_content == true) vtCon = XlLookAt::xlWhole;

		if (PRS = pRange->GetRange(COleVariant(szval))->Find(bvalue, vtMissing, XlFindLookIn::xlValues, vtCon,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, match_case, false, false))
		{
			iRowValue = PRS->Cells->Row;
		}

		pRange->GetRange(COleVariant(L"A1:A1"))->Find("", vtMissing, XlFindLookIn::xlFormulas, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false);

		if (iRowValue > 0) return iRowValue;
	}
	catch (...) {}
	return 0;
}


CString _xlGAR(RangePtr pRng, int nRow, int nColumn, int itype)
{
	RangePtr dRangeptr = pRng->GetItem(nRow, nColumn);
	_bstr_t szAdress = dRangeptr->GetAddress(false, false, XlReferenceStyle::xlA1);
	CString kq = (LPCTSTR)szAdress;
	if (itype == 0) return kq;

	int k = 0;
	int iLen = kq.GetLength();
	for (int i = 1; i < iLen; i++)
	{
		if (_wtoi(kq.Mid(i, 1)) > 0)
		{
			k = i;
			break;
		}
	}

	CString szval = L"";
	if (itype == 1) szval.Format(L"$%s", kq);
	else if (itype == 2) szval.Format(L"%s$%s", kq.Left(k), kq.Right(iLen - k));
	else if (itype == 3) szval.Format(L"$%s$%s", kq.Left(k), kq.Right(iLen - k));

	return szval;
}


bool _xlGetPointWindowsOfCell(RangePtr pRng, int &pX, int &pY)
{
	try
	{
		HDC hScreenDC = GetDC(0);
		int dpiX = static_cast<int>(GetDeviceCaps(hScreenDC, LOGPIXELSX));
		int dpiY = static_cast<int>(GetDeviceCaps(hScreenDC, LOGPIXELSY));
		ReleaseDC(0, hScreenDC);

		int zoom = xl->ActiveWindow->GetZoom();
		int zoomRatio = (int)(zoom/100);
		int pointsPerInch = xl->Application->InchesToPoints(1);

		pX = xl->ActiveWindow->PointsToScreenPixelsX(0);
		pY = xl->ActiveWindow->PointsToScreenPixelsY(0);

		pX += (int)(pRng->Left)*zoom*dpiX/pointsPerInch/100;
		pY += (int)(pRng->Top)*zoom*dpiY/pointsPerInch/100;

		pX -= 3;/* pY += 1;*/

		return true;
	}
	catch (...) {}
	return false;
}


/**********************************************************************************/

bool IsCapsLockOn()
{
	if ((GetKeyState(VK_CAPITAL) & 0x00001) != 0) return true;
	else return false;
}

CString GetClassActiveWindow()
{
	CString szClass = L"";
	wchar_t *wch = new wchar_t[255];

	DWORD winId = NULL;
	hWndMain = GetActiveWindow();
	GetWindowThreadProcessId(hWndMain, &winId);
	GetClassNameW(hWndMain, wch, 255);

	szClass.Format(L"%s", wch);
	return szClass;
}

bool CheckKeyBoard(WPARAM wParam)
{
	try
	{
		if (wParam == VK_F2) jKeyBeforeF2 = 0;
		else
		{
			jKeyBeforeF2++;
			if (jKeyBeforeF2 > 2) jKeyBeforeF2 = 2;
		}

		if (wParam == VK_RETURN) blKeyEnter = true;
		else blKeyEnter = false;

		switch (wParam)
		{
			// Special
			case VK_CAPITAL:
			case VK_LCONTROL:
			case VK_RCONTROL:
			case VK_INSERT:
			case VK_END:
			case VK_PRINT:
			case VK_DELETE:
			case VK_BACK:
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_ESCAPE:
			case VK_MENU:
			case VK_TAB:
			{
				return false;
			}
			break;
		}

		if (wParam >= 0x70 && wParam <= 0x7B) return false;		 // F1 ~ F12.

		// ENTER || SPACE || a~z + 0~9 + ...
		if (wParam == VK_RETURN || wParam == VK_SPACE || (wParam >= 0x2f) && (wParam <= 0x100))
		{
			return true;
		}
	}
	catch (...) {}
	return false;
}

int KiemtraDieukienHook()
{
	try
	{
		int iRowMaso = 0, iColMaso = 0;
		int nLen = ((CString)shChiphi).GetLength();

		if (sCodeActive == shQuytrinh)
		{
			iColMaso = _xlGetColumn(pShActive, L"QTR_MASO", -1);
			if (iColMaso > 0 && iColMaso == iColumnActive)
			{
				iRowMaso = _xlGetRow(pShActive, L"QTR_MASO", -1);
				if (iRowMaso > 0 && iRowMaso < iRowActive) return 1;
			}
		}
		else if (sCodeActive == shTieuchuan)
		{
			iColMaso = _xlGetColumn(pShActive, L"TCH_MATC", -1);
			if (iColMaso > 0 && iColMaso == iColumnActive)
			{
				iRowMaso = _xlGetRow(pShActive, L"TCH_MATC", -1);
				if (iRowMaso > 0 && iRowMaso < iRowActive) return 2;
			}
		}
		else if (sCodeActive.Left(nLen) == shChiphi)
		{
			if (sCodeActive.GetLength() > nLen)
			{
				_bstr_t bs5540 = (LPCTSTR)_xlGetNameSheet(shChiphi, 0);
				if (bs5540 != (_bstr_t)L"")
				{
					_WorksheetPtr pSh5540 = xl->Sheets->GetItem((_bstr_t)bs5540);

					int iColNDung = _xlGetColumn(pSh5540, L"M5540_NDUNG", -1);
					if (iColNDung > 0 && iColNDung == iColumnActive)
					{
						int iRowNDung = _xlGetRow(pSh5540, L"M5540_NDUNG", -1);
						if (iRowNDung > 0 && iRowNDung < iRowActive) return 3;
					}
				}
			}
		}
	}
	catch (...) {}
	return 0;
}

CString ConvertUnikey(CString szKey)
{
	CString szUni = L"", szConv = L"";
	if (!IsCapsLockOn())
	{
		szUni = L"ddaaawaưafasaraxajeeefeserexejifisirixijooowoưofosoroxojuwuưufusuruxujyfysyryxyj";
		szConv = L"đâăăàáảãạêèéẻẽẹìíỉĩịôơơòóỏõọưưùúủũụỳýỷỹỵ";
	}
	else
	{
		szUni = L"DDAAAWAƯAFASARAXAJEEEFESEREXEJIFISIRIXIJOOOWOƯOFOSOROXOJUWUƯUFUSURUXUJYFYSYRYXYJ";
		szConv = L"ĐÂĂĂÀÁẢÃẠÊÈÉẺẼẸÌÍỈĨỊÔƠƠÒÓỎÕỌƯƯÙÚỦŨỤỲÝỶỸỴ";
	}

	int nLen = (int)(szUni.GetLength() / 2);
	for (int i = 0; i < nLen; i++)
	{
		if (szUni.Mid(2 * i, 2) == szKey)
		{
			return szConv.Mid(i, 1);
		}
	}

	return szKey;
}