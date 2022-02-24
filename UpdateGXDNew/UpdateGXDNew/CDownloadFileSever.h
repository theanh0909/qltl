#pragma once
#include "Base64.h"

class CDownloadFileSever
{

public:
	CDownloadFileSever();
	~CDownloadFileSever();

	Base64Ex *bb;

	CString dbCheckConn;
	CString dbSeverMain, dbSeverDir, dbSeverData, dbSeverDirSM;

	void _LoadDecodeBase64();

	bool _CheckConnection();
	bool _DownloadFile(CString szURL, CString pathSaved);
	int _CheckVersionSoftware();

};

