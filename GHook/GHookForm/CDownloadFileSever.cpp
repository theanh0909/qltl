#include "pch.h"
#include "CDownloadFileSever.h"

CDownloadFileSever::CDownloadFileSever()
{
	dbCheckConn = L"";
	dbSeverMain = L"", dbSeverDir = L"";
}

CDownloadFileSever::~CDownloadFileSever()
{

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
		dbCheckConn = BaseDecodeEx(CHECKCONN);
		dbSeverMain = BaseDecodeEx(SEVERMAIN);

		CString szVersionSoft = L"";
		szVersionSoft.Format(L"%02d", CheckVersionSoftware());
		
		dbSeverDir = dbSeverMain + BaseDecodeEx(SEVERNEXT);
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
			if (FileExistsUnicode(pathSaved) == 1) return true;
		}
	}
	catch (...) {}
	return false;
}

