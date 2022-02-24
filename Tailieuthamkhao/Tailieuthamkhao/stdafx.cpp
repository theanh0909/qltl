// stdafx.cpp : source file that includes just the standard includes
// Tailieuthamkhao.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

_ApplicationPtr xl;
_WorkbookPtr pWb;

CConnection ObjConn;
CString SqlString = L"";

CString spathDefault = L"";
int RowEditList = -1, ColumnEditList = -1;

double jScaleLayout = 1;

int jCheckRotot = 0;
int iSaveCopyMulti = 0;
int iSaveCreateFolder = 1;
int irHeightDOWN = 0, irWidthDOWN = 0;

CFont myfontCheckbot, m_FontHeaderList;