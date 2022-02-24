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

Function::Function()
{
}

Function::~Function()
{
}

void Function::_msgbox(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void Function::_s(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void Function::_d(int num, int itype, int iexcel, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, iexcel, itimeout);
}

void Function::_n(CString message, int num, int itype, int istr, int iexcel, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (istr == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, iexcel, itimeout);
}


int Function::_y(CString message, int istyle, int itype, int iexcel, int itimeout)
{
	UINT qs[2] = { MB_YESNO, MB_YESNOCANCEL };
	UINT style[2] = { MB_ICONQUESTION, MB_ICONWARNING };

	return (int)MessageBoxTimeout(NULL, message, MS_THONGBAO,
		qs[istyle] | style[itype], 0, itimeout);

	return 6;
}

/***********************************************************************************************/

int Function::_CheckVersionSoftware()
{
	return _wtoi(VERSION_SOFT);
}

char* Function::_ConvertCstringToChars(CString szvalue)
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

CString Function::_ConvertCharsToCstring(char* chr)
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

int Function::_FileExists(CString pathFile, int itype)
{
	CString str = L"Đường dẫn file không hợp lệ. "
		L"Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		struct stat fileInfo;
		if (!stat(_ConvertCstringToChars(pathFile), &fileInfo)) return 1;
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}
	catch (...)
	{
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}
	return 0;
}

int Function::_FileExistsUnicode(CString pathFile, int itype)
{
	CString str = L"Đường dẫn file không hợp lệ. Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		LPTSTR lpPath = pathFile.GetBuffer(MAX_PATH);
		if (PathFileExists(lpPath)) return 1;
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}
	catch (...)
	{
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}
	return 0;
}

bool Function::_FolderExistsUnicode(CString dirName_in)
{
	DWORD dwAttrib = GetFileAttributes((LPCTSTR)dirName_in);
	return (dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
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

bool Function::_CreateDirectories(CString pathName)
{
	try
	{
		if (pathName.Right(1) != L"\\") pathName += L"\\";
		for (int i = 0; i < pathName.GetLength(); i++)
		{
			if (pathName.Mid(i, 1) == L"\\") _CreateNewFolder(pathName.Left(i));
		}
		return true;
	}
	catch (...) {}
	return false;
}

int Function::_FindAllFile(CFileFinder &_finder, CString pathFolder, CString typeOfFile)
{
	try
	{
		CFileFinder::CFindOpts opts;
		opts.sBaseFolder = pathFolder;
		opts.sFileMask = L"*" + typeOfFile + L"*";
		opts.FindNormalFiles();

		_finder.RemoveAll();
		_finder.Find(opts);

		return _finder.GetFileCount();
	}
	catch (...) {}
	return 0;
}

CString Function::_GetPathFolder(CString fileName)
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

bool Function::_IsProcessRunning(const wchar_t *processName)
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

int Function::_TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1, CString symbol2, int itypeFilter, int itypeTrim)
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

int Function::_RandomTime()
{
	time_t now = time(0);
	tm *ltm = localtime(&now);
	return (100*(int)ltm->tm_min + (int)ltm->tm_sec + rand() % 1000);
}

CString Function::_GetTimeNow(int ifulltime)
{
	try
	{
		time_t currentTime;
		time(&currentTime);
		struct tm *localTime = localtime(&currentTime);
		int iYear = (int)localTime->tm_year;
		if (iYear < 1900) iYear += 1900;

		CString szDaynow = L"";
		szDaynow.Format(L"%02d/%02d/%d",
			(int)localTime->tm_mday, 1 + (int)localTime->tm_mon, iYear);

		if (ifulltime == 1)
		{
			CString sztimer = L"";
			sztimer.Format(L" %02d:%02d:%02d",
				(int)localTime->tm_hour, (int)localTime->tm_min, (int)localTime->tm_sec);
			szDaynow += sztimer;
		}
		return szDaynow;
	}
	catch (...) {}
	return L"";
}

bool Function::_CompareDate(CString dateBefore, CString dateAfter)
{
	int d1 = _wtoi(dateBefore.Left(2));
	int d2 = _wtoi(dateAfter.Left(2));

	int m1 = _wtoi(dateBefore.Mid(3, 2));
	int m2 = _wtoi(dateAfter.Mid(3, 2));

	int y1 = _wtoi(dateBefore.Right(2));
	int y2 = _wtoi(dateAfter.Right(2));

	if (y1 < y2 || (y1 == y2 && m1 < m2) || (y1 == y2 && m1 == m2 && d1 <= d2)) return true;

	return false;
}

CString	Function::_DayNextPrev(CString szdate, int num)
{
	if (szdate == L"") return L"";
	if (num == 0) return szdate;

	CString szDay = szdate;
	int dd = _wtoi(szDay.Left(2));
	int mm = _wtoi(szDay.Mid(3, 2));
	int yy = _wtoi(szDay.Right(2)) + 2000;
	int ddmax = 31;

	if (num > 0)
	{
		while (true)
		{
			ddmax = 31;
			if (mm == 2)
			{
				if (yy % 400 == 0 || (yy % 4 == 0 && yy % 100 != 0)) ddmax = 29;
				else ddmax = 28;
			}
			else if (mm == 4 || mm == 6 || mm == 9 || mm == 11) ddmax = 30;

			if (num + dd > ddmax)
			{
				num -= (ddmax - dd + 1);

				dd = 1;
				if (mm < 12) mm++;
				else
				{
					mm = 1;
					yy++;
				}
			}
			else
			{
				dd += num;
				break;
			}
		}
	}
	else
	{
		num = abs(num);
		while (true)
		{
			if (dd - num < 1)
			{
				mm--;
				if (mm <= 0)
				{
					mm = 12;
					yy--;
				}

				ddmax = 31;
				if (mm == 2)
				{
					if (yy % 400 == 0 || (yy % 4 == 0 && yy % 100 != 0)) ddmax = 29;
					else ddmax = 28;
				}
				else if (mm == 4 || mm == 6 || mm == 9 || mm == 11) ddmax = 30;

				num -= dd;
				dd = ddmax;
			}
			else
			{
				dd -= num;
				break;
			}
		}
	}

	szDay.Format(L"%02d/%02d/%04d", dd, mm, yy);
	return szDay;
}

CString Function::_GetTypeOfFile(CString pathFile)
{
	try
	{
		if (pathFile != L"")
		{
			if ((int)pathFile.Find(L".") > 0)
			{
				int nLen = pathFile.GetLength();
				for (int i = nLen - 1; i >= 0; i--)
				{
					if (pathFile.Mid(i, 1) == L".")
					{
						return pathFile.Right(nLen - i - 1);
					}
				}
			}			
		}
	}
	catch (...) {}
	return L"";
}

CString Function::_GetNameOfPath(CString pathFile, int &pos, int ipath)
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

CString Function::_GetDesktopDir()
{
	CString szDesktop = L"";
	char* szpath = new char[MAX_PATH + 1];
	if (SHGetSpecialFolderPathA(HWND_DESKTOP, szpath, CSIDL_DESKTOP, FALSE))
	{
		szDesktop = _ConvertCharsToCstring(szpath);
		if (szDesktop.Right(1) != L"\\") szDesktop += L"\\";
	}
	return szDesktop;
}

void Function::_WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (_FileExistsUnicode(spathFile, 0) != 1) return;
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

int Function::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExistsUnicode(spathFile, 0) != 1) return 0;

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

CString Function::_GetDirModules(DWORD processID, bool blFilter)
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

void Function::_LogFileWrite(CString szLog, bool blClear, bool blOpenFile, CString szFileName)
{
	try
	{
		CString szpathFileLog = _GetDesktopDir() + szFileName;
		
		vector<CString> vecLine;
		if (!blClear) _ReadMultiUnicode(szpathFileLog, vecLine);
		
		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(szpathFileLog));
		if (myfile)
		{
			myfile.imbue(loc);

			int itotal = vecLine.size();
			if (itotal > 0)
			{
				for (int i = 0; i < itotal; i++)
				{
					wstring ws(vecLine[i]);
					myfile << ws << "\n";
				}
			}

			if (szLog != "<!>")
			{
				wstring ws(szLog);
				myfile << ws << "\n";
			}

			myfile.close();
		}

		vecLine.clear();

		if (blOpenFile) ShellExecute(NULL, L"open", szpathFileLog, NULL, NULL, SW_SHOWMAXIMIZED);
	}
	catch (...) {}
}

void Function::_LogFileWrite(int iLog, bool blClear, bool blOpenFile, CString szFileName)
{
	CString str = L"";
	str.Format(L"%d", iLog);
	_LogFileWrite(str, blClear, blOpenFile, szFileName);
}

void Function::_LogFileStart(CString szStart) { _LogFileWrite(szStart, true); }
void Function::_LogFileEnd() { _LogFileWrite(L"<!>", false, true); }

CString Function::_ConvertKhongDau(CString szNoidung)
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

double Function::_RoundDouble(double doValue, int nPrecision)
{
	static const double doBase = 10.0;
	double doComplete5, doComplete5i;

	doComplete5 = doValue * pow(doBase, (double)(nPrecision + 1));

	if (doValue < 0.0) doComplete5 -= 5.0;
	else doComplete5 += 5.0;

	doComplete5 /= doBase;
	modf(doComplete5, &doComplete5i);

	return doComplete5i / pow(doBase, (double)nPrecision);
}

CString Function::_ConvertBytes(double ivalue)
{
	CString szval = L"";
	CString cs[4] = { L"KB", L"MB", L"GB", L"TB" };
	for (int i = 0; i < 4; i++)
	{
		ivalue = ivalue / 1024;
		if (ivalue < 1024)
		{
			ivalue = _RoundDouble(ivalue, 2);
			szval.Format(L"%4.2f %s", ivalue, cs[i]);
			return szval;
		}
	}

	szval.Format(L"%f bytes", ivalue);
	return szval;
}

void Function::_GetFileAttributesEx(CString szDir, CString &szKichthuoc, CString &szNgay)
{
	try
	{
		szKichthuoc = L"";
		szNgay = L"";

		HMODULE hLib;
		GET_FILE_ATTRIBUTES_EX func;
		FILE_INFO fInfo;

		hLib = LoadLibrary(L"Kernel32.dll");
		if (hLib != NULL)
		{
			func = (GET_FILE_ATTRIBUTES_EX)GetProcAddress(hLib, "GetFileAttributesExW");
			if (func != NULL) func(szDir, 0, &fInfo);

			SYSTEMTIME times, stLocal;
			FileTimeToSystemTime(&fInfo.ftLastWriteTime, &times);
			SystemTimeToTzSpecificLocalTime(NULL, &times, &stLocal);

			szKichthuoc = _ConvertBytes((double)fInfo.nFileSizeLow);

			if((int)stLocal.wYear >= 2000)
			{
				szNgay.Format(L"%02d/%02d/%04d %02d:%02d:%02d",
					stLocal.wDay, stLocal.wMonth, stLocal.wYear,
					stLocal.wHour, stLocal.wMinute, stLocal.wSecond);
			}

			FreeLibrary(hLib);
		}
	}
	catch (...) {}
}

CString Function::_FormatTienVND(CString szTien, CString szSteparator, CString szDecimal)
{
	if (szTien == L"") return L"";
	CString snguyen = szTien;
	CString stphan = L"";
	int pos = szTien.Find(szDecimal);
	if (pos > 0)
	{
		snguyen = szTien.Left(pos);
		stphan = szTien.Right(szTien.GetLength() - pos - 1);
	}

	CString szval = L"";
	int len = snguyen.GetLength();
	int pnguyen = (int)(len / 3);
	int ple = len % 3;

	if (ple > 0) szval = snguyen.Left(ple);
	if (pnguyen >= 1)
	{
		if (ple > 0 && szval != L"-") szval += szSteparator;
		for (int i = 0; i < pnguyen; i++)
		{
			szval += snguyen.Mid(ple + 3 * i, 3);
			if (i < pnguyen - 1) szval += szSteparator;
		}
	}

	if (szval == L"") szval = snguyen;
	if (stphan != L"")
	{
		szval += szDecimal;
		szval += stphan;
	}

	return szval;
}

CString Function::_ConvertRename(CString sname)
{
	sname.Replace(L" ", L"");
	if (sname == L"") return L"";

	sname.Replace(L"\\", L"");
	sname.Replace(L"/", L"");
	sname.Replace(L":", L"");
	sname.Replace(L"*", L"");
	sname.Replace(L"?", L"");
	sname.Replace(L"\"", L"");
	sname.Replace(L"<", L"");
	sname.Replace(L">", L"");
	sname.Replace(L"|", L"");

	return sname;
}

COLORREF Function::_SetColorRGB(CString szColor, int itype)
{
	szColor.Trim();
	szColor.Replace(L"RGB(", L"");
	szColor.Replace(L")", L"");

	if (szColor == L"")
	{
		if (itype == 0) return RGBWHITE;
		else return RGBBLACK;
	}

	int pos = szColor.Find(L",");
	int iRed = _wtoi(szColor.Left(pos));

	CString szval = szColor.Right(szColor.GetLength() - pos - 1);
	pos = szval.Find(L",");
	int iGreen = _wtoi(szval.Left(pos));
	int iBlue = _wtoi(szval.Right(szval.GetLength() - pos - 1));

	BYTE bRed = iRed;
	BYTE bGreen = iGreen;
	BYTE bBlue = iBlue;

	return RGB(bRed, bGreen, bBlue);
}

CString Function::_GetColorRGB(COLORREF clr)
{
	CString szval = L"";
	int iRed = (int)GetRValue(clr);
	int iGreen = (int)GetGValue(clr);
	int iBlue = (int)GetBValue(clr);
	szval.Format(L"RGB(%d,%d,%d)", iRed, iGreen, iBlue);
	return szval;
}

void Function::_ShellExecuteSelectFile(CString szDirFileSelect)
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

void Function::_CallFunctionDLL(CString szDllName, CString szFunction)
{
	try
	{
		typedef bool(__stdcall *p)();									// ...__stdcall *p)(int opt) nếu có đối là 'opt'
		HINSTANCE loadDLL = LoadLibrary((_bstr_t)szDllName);			// Ex: UtilityQLCL9.dll

		if (loadDLL != NULL)
		{
			p pc = (p)GetProcAddress(loadDLL, (_bstr_t)szFunction);		// Ex: OpenFrmInhoso
			if (pc != NULL) pc();
		}

		FreeLibrary(loadDLL);
	}
	catch (...) {}
}

bool Function::_Is64BitWindows()
{
#if defined(_WIN64)
	return TRUE;  // 64-bit programs run only on Win64
#elif defined(_WIN32)
	// 32-bit programs run on both 32-bit and 64-bit Windows
	// so must sniff
	BOOL f64 = FALSE;
	return IsWow64Process(GetCurrentProcess(), &f64) && f64;
#else
	return FALSE; // Win64 does not support Win16
#endif
}