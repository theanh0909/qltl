#pragma once
#include "StdAfx.h"

class CQSortData
{
public:
	CQSortData();
	~CQSortData();

	Function *ff;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

public:
	void _SapxepTatcaDulieu();
	void _XacdinhSheetLienquan();
	bool _CheckFileMove(int nRow);
};

