#include "pch.h"
#include "CDownloadFileSever.h"

#define VERSION_SOFT	L"1"
#define CHECKCONN		L"JsbXhaTE52ZzRUlVY4=l6RTIYMMUhZSGNTeTJiOT1uOD"
#define SEVERMAIN		L"51MlpMTFIxYkdsbGRTTUMcWBOO=D1RH2aS5DBkkEETa2RYUnZZV0dlNT1uUW"
//#define SEVERNEXT		L"R2xTbGJkOTdlUEV9EgS5EGWVlobRWHJdhzlYdEwzR2RGcHNG"
#define SEVERNEXT		L"M3SDlZZWg5bH9IEVVW=EEUnXYaMUcR5GJVM0FaRj1zPV"
#define SEVERNEXTSM		L"eUNGZGQ5N2VIaDlUcgR=0ZRQ=VsTnXYaMUcR5GJVM050WVh3Y0o9MD1H"
#define SEVERDATA		L"aEdSY2RWVgR=DJTV=VPOESUMzmxUj1JVTFI9QlVH"

CDownloadFileSever::CDownloadFileSever()
{
	bb = new Base64Ex;

	dbCheckConn = L"";
	dbSeverMain = L"", dbSeverDir = L"", dbSeverData = L"", dbSeverDirSM = L"";
}

CDownloadFileSever::~CDownloadFileSever()
{
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
		szVersionSoft.Format(L"%02d", _CheckVersionSoftware());
		
		dbSeverDir = dbSeverMain + bb->BaseDecodeEx(SEVERNEXT);
		/*MessageBox(NULL, dbSeverMain,
			L"Hướng dẫn", NULL);*/

		//dbSeverDir.Replace(L"{xx}", szVersionSoft);
		dbSeverDir.Replace(L"{xx}", L"");
		if (dbSeverDir.Right(1) != L"/") dbSeverDir += L"/";

		dbSeverData = dbSeverMain + bb->BaseDecodeEx(SEVERDATA);
		if (dbSeverData.Right(1) != L"/") dbSeverData += L"/";

		dbSeverDirSM = dbSeverMain + bb->BaseDecodeEx(SEVERNEXTSM);
		dbSeverDirSM.Replace(L"{xx}", szVersionSoft);
		if (dbSeverDirSM.Right(1) != L"/") dbSeverDirSM += L"/";
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
			if (_FileExistsUnicode(pathSaved) == 1) return true;
		}		
	}
	catch (...) {}
	return false;
}

int CDownloadFileSever::_CheckVersionSoftware()
{
	return _wtoi(VERSION_SOFT);
}