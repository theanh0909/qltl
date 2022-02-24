#pragma once
#include "StdAfx.h"
#include "CDownloadFileSever.h"

class CCheckVersion
{
public:
	CCheckVersion();
	~CCheckVersion();

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;
	CDownloadFileSever *sv;

	CString pathFolder;

public:
	int _CompareVersion(int iQuestion = 0);	
	void _OpenFileManagerApp();
	void _OpenFileUpdateGXD();
};

