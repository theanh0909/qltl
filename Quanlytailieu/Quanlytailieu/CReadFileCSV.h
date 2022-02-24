#pragma once
#include "StdAfx.h"

class CReadFileCSV
{

public:
	CReadFileCSV();
	~CReadFileCSV();

	Function *ff;

	int _ReadLineFile(CString pathFileCSV, vector<CString> &vecItem);
	int _GetLineFile(CString pathFileCSV);
	int _GetItemStr(CString szItem, vector<CString> &vecLine);
};
