#pragma once
#include "StdAfx.h"

class CRenameFolders
{

public:
	CRenameFolders();
	~CRenameFolders();

	Function *ff;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

public:
	void _GetDanhsachFolders();	
	void _XacdinhSheetLienquan();
	void _RenameAllChildFolder(CString szNameFd, CString szDirOld, CString szChange);

};

