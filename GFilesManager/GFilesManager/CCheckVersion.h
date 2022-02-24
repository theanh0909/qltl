#pragma once
#include "pch.h"
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

public:
	int _CompareVersion(CString &szNew);

};

