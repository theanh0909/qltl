#pragma once
#include "StdAfx.h"

class CDownloadFileSever
{

public:
	CDownloadFileSever();
	~CDownloadFileSever();

	Function *ff;
	Base64Ex *bb;

	CString dbCheckConn;
	CString dbSeverMain, dbSeverDir, dbSeverData;

	void _LoadDecodeBase64();

	bool _CheckConnection();
	bool _DownloadFile(CString szURL, CString pathSaved);

};

