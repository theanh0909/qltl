#include "pch.h"
#include "CCheckVersion.h"

CCheckVersion::CCheckVersion()
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;

	pathFolder = L"";
}

CCheckVersion::~CCheckVersion()
{
	delete ff;
	delete bb;
	delete sv;
	delete reg;
}

int CCheckVersion::_CompareVersion(int iQuestion)
{
	try
	{
		pathFolder = ff->_GetPathFolder(strNAMEDLL);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		// Kiểm tra version trong phần mềm
		vector<CString> vecLine;

		// Ban đầu để mặc định 2 giá trị bất kỳ nhưng phải khác nhau
		CString szVerUser = L"<1>", szVerUp = L"<2>";

		int ncount = ff->_ReadMultiUnicode(pathFolder + FileVersion, vecLine);
		if (ncount >= 5) szVerUser = vecLine[4];

		// Download file version trên sever
		sv->_LoadDecodeBase64();
		CString szFileSaved = pathFolder + FileHelp;
		if (sv->_DownloadFile(sv->dbSeverDir + FileHelp, szFileSaved))
		{
			vecLine.clear();
			ncount = ff->_ReadMultiUnicode(szFileSaved, vecLine);
			if (ncount > 0)
			{
				for (int i = ncount - 1; i >= 0; i--)
				{
					vecLine[i].Trim();
					if (vecLine[i] == L"" || vecLine[i].Left(2) == L"//") vecLine.erase(vecLine.begin() + i);
				}
			}

			ncount = (int)vecLine.size();
			if (ncount >= 5) szVerUp = vecLine[4];
		}

		DeleteFile(szFileSaved);

		vecLine.clear();

		if (szVerUp != szVerUser)
		{
			if (iQuestion == 1)
			{
				int result = ff->_y(L"Bạn đang sử dụng phiên bản cũ. Đã có bản cập nhật mới hơn. Bạn có muốn cập nhật?", 0, 0, 0);
				if (result != 6) return -1;
			}

			return 1;
		}
		else return 0;		
	}
	catch (...) {}
	return 1;
}

void CCheckVersion::_OpenFileManagerApp()
{
	try
	{
		pathFolder = ff->_GetPathFolder(strNAMEDLL);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		CString szCreate = bb->BaseDecodeEx(CreateKeySMStart) + UploadData;
		reg->CreateKey(szCreate);

		int iCheckStart = 1;
		CString szval = reg->ReadString(StartAppWithWindows, L"");
		if (szval != L"") iCheckStart = _wtoi(szval.Trim());
		if (iCheckStart != 1) iCheckStart = 0;

		szval.Format(L"%d", iCheckStart);
		reg->WriteChar(ff->_ConvertCstringToChars(StartAppWithWindows),
			ff->_ConvertCstringToChars(szval));

/********************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeyRun);
		reg->CreateKey(szCreate);

		if (iCheckStart == 1)
		{
			szval.Format(_T("%s%s.exe"), pathFolder, FileSmartStartApp);
			reg->WriteChar(
				ff->_ConvertCstringToChars(FileSmartStartApp),
				ff->_ConvertCstringToChars(szval));
		}
		else reg->DeleteValue(FileSmartStartApp);

/********************************************************************************/

		CString szAppRun = L"";
		szAppRun.Format(L"%s.exe", FileSmartStartApp);

		if (!ff->_IsProcessRunning(szAppRun))
		{
			CString szFileSaved = pathFolder + szAppRun;
			if (ff->_FileExistsUnicode(szFileSaved, 0) != 1)
			{
				sv->_DownloadFile(sv->dbSeverDir + szAppRun, szFileSaved);
			}

			if (ff->_FileExistsUnicode(szFileSaved, 0) == 1)
			{
				ShellExecute(NULL, L"open", szFileSaved, NULL, NULL, SW_NORMAL);
			}
		}
	}
	catch (...) {}
}

void CCheckVersion::_OpenFileUpdateGXD()
{
	pathFolder = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

	CString szFileSaved = pathFolder + FileUpdateGXDNew;
	if (sv->_DownloadFile(sv->dbSeverDir + szOffice + L"\\" + FileUpdateGXDNew, szFileSaved))
	{
		ShellExecute(NULL, L"open", szFileSaved, NULL, NULL, SW_NORMAL);
	}
}
