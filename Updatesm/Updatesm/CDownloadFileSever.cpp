#include "pch.h"
#include "CDownloadFileSever.h"

CDownloadFileSever::CDownloadFileSever()
{
	ff = new Function;
	bb = new Base64Ex;

	dbCheckConn = L"";
	dbSeverMain = L"", dbSeverDir = L"";
}

CDownloadFileSever::~CDownloadFileSever()
{
	delete ff;
	delete bb;
}

bool CDownloadFileSever::_CheckConnection()
{
	if (InternetCheckConnection(dbCheckConn, FLAG_ICC_FORCE_CONNECTION, 0)) return true;
	return false;
}

void CDownloadFileSever::_LoadDecodeBase64()
{
	try
	{
		dbCheckConn = bb->BaseDecodeEx(CHECKCONN);
		dbSeverMain = bb->BaseDecodeEx(SEVERMAIN);

		CString szVersionSoft = L"";
		szVersionSoft.Format(L"%02d", ff->_CheckVersionSoftware());
		
		dbSeverDir = dbSeverMain + bb->BaseDecodeEx(SEVERNEXT);
		dbSeverDir.Replace(L"{xx}", szVersionSoft);
		if (dbSeverDir.Right(1) != L"\\") dbSeverDir += L"\\";
	}
	catch (...) {}
}

bool CDownloadFileSever::_DownloadFile(CString szURL, CString pathSaved)
{
	try
	{
		if (_CheckConnection())
		{
			szURL.Replace(L"\\", L"/");
			DeleteUrlCacheEntry(szURL);
			URLDownloadToFileW(0, szURL, pathSaved, 0, 0);
			if (ff->_FileExistsUnicode(pathSaved) == 1) return true;
		}
	}
	catch (...) {}
	return false;
}

