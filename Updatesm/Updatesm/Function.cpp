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

void Function::_msgbox(CString message, int itype, int itimeout)
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

void Function::_s(CString message, int itype, int itimeout)
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

void Function::_d(int num, int itype, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, itimeout);
}

void Function::_n(CString message, int num, int itype, int istr, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (istr == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, itimeout);
}


int Function::_y(CString message, int istyle, int itype, int itimeout)
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
				if (!blFilter) szProcessDir = (LPCTSTR)processDir;
				else
				{
					DWORD cbNeeded;
					HMODULE hMods[1024];
					if (EnumProcessModulesEx(hProcess, hMods,
						sizeof(hMods), &cbNeeded, LIST_MODULES_ALL))
					{
						szProcessDir = (LPCTSTR)processDir;
					}
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

void Function::_LogFileStart(CString szStart) { _LogFileWrite(szStart, true); }
void Function::_LogFileEnd() { _LogFileWrite(L"<!>", false, true); }

CString Function::_GetNameOfPath(CString pathFile, int ipath)
{
	try
	{
		if (pathFile != L"")
		{
			if (pathFile.Right(1) == L"\\") pathFile = pathFile.Left(pathFile.GetLength() - 1);

			int nLen = pathFile.GetLength();
			for (int i = nLen - 1; i >= 0; i--)
			{
				if (pathFile.Mid(i, 1) == L"\\" || pathFile.Mid(i, 1) == L"/")
				{
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

int Function::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExistsUnicode(spathFile) != 1) return 0;

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

int Function::_FileExistsUnicode(CString pathFile)
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
