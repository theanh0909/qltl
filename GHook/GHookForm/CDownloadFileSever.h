#pragma once

class CDownloadFileSever
{

public:
	CDownloadFileSever();
	~CDownloadFileSever();

	CString dbCheckConn;
	CString dbSeverMain, dbSeverDir;

	void _LoadDecodeBase64();

	bool _CheckConnection();
	bool _DownloadFile(CString szURL, CString pathSaved);

};

