// pch.cpp: source file corresponding to the pre-compiled header

#include "pch.h"

// When you are using pre-compiled headers, this source file is necessary for compilation to succeed.

int _FileExistsUnicode(CString pathFile)
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

char* _ConvertCstringToChars(CString szvalue)
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

CString _GetTimeNow(int ifulltime)
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

bool _IsProcessRunning(const wchar_t *processName)
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

CString _GetDirModules(DWORD processID, bool blFilter)
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