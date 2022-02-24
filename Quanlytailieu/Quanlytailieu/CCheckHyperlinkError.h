#pragma once
#include "StdAfx.h"

class CCheckHyperlinkError
{
public:
	CCheckHyperlinkError();
	~CCheckHyperlinkError();

	Function *ff;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;
	bool blCheckMaxLen;

public:

	void _KiemtraLienketError();
	void _XacdinhSheetLienquan();
	int _CheckError(int iRow, int iColumn);
};

