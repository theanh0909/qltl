#include "pch.h"

_ApplicationPtr xl;
_WorkbookPtr pWb;

_WorksheetPtr pShActive;
RangePtr pRgActive;
CString szCodeActive = L"", szNameActive = L"";
COLORREF bkgColorActive = RGB(255, 255, 255);
int iRowActive = 1, iColumnActive = 1;

CConnection ObjConn;
CString SqlString = L"";

int compareStringAZ(const void * a, const void * b)
{
	if (*(CString*)a < *(CString*)b) return -1;
	else if (*(CString*)a > *(CString*)b) return  1;
	return 0;
}