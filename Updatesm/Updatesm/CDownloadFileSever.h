#pragma once
#include "pch.h"

class CDownloadFileSever
{

public:
	CDownloadFileSever();
	~CDownloadFileSever();

	Function *ff;
	Base64Ex *bb;

	CString dbCheckConn;
	CString dbSeverMain, dbSeverDir;

	void _LoadDecodeBase64();

	bool _CheckConnection();
	bool _DownloadFile(CString szURL, CString pathSaved);

};

