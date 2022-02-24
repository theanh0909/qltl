#include "pch.h"
#include "CCheckVersion.h"

CCheckVersion::CCheckVersion()
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;
}

CCheckVersion::~CCheckVersion()
{
	delete ff;
	delete bb;
	delete sv;
	delete reg;
}

int CCheckVersion::_CompareVersion(CString &szNew)
{
	try
	{
		CString szNewVersion = L"";
		CString pathFolder = ff->_GetPathFolder(szAppSmart);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		// Lấy thông tin version trong Registry
		CString szCreate = bb->BaseDecodeEx(CreateKeySMStart);
		reg->CreateKey(szCreate);
		CString szOldVersion = reg->ReadString(Version, L"");

		// Lấy thông tin version trên Sever để so sánh
		sv->_LoadDecodeBase64();
		CString szFileVersion = pathFolder + FileVersionsm;
		if (sv->_DownloadFile(sv->dbSeverDir + FileVersionsm, szFileVersion))
		{
			vector<CString> vecLine;
			int ncount = ff->_ReadMultiUnicode(szFileVersion, vecLine);
			if (ncount > 0) szNewVersion = vecLine[0];	
			vecLine.clear();
			DeleteFile(szFileVersion);
		}

		if (szNewVersion != L"" && szOldVersion != szNewVersion)
		{
			szNew = szNewVersion;
			return 1;
		}

		return 0;
	}
	catch (...) {}
	return 1;
}
