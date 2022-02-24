#pragma once
#include "StdAfx.h"

class CMoveFiles
{
public:
	CMoveFiles();
	~CMoveFiles();

	Function *ff;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

public:
	void _DichuyenDulieu();
	void _XacdinhSheetLienquan();	
	void _AutoSothutu(_WorksheetPtr pSheet, int numStart = 1, int nRowStart = 0, int nRowEnd = 0);

	int _XacdinhVitriChen(CString szPathFind);
	bool _DichuyenFile(int iRow, int iRChen, CString szPathFind, CString szLinhvuc);
	bool _CheckMoveFiles(int iRow);
};

