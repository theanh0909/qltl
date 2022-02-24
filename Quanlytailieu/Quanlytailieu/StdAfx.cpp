#include "pch.h"
#include "StdAfx.h"

CString strNAMEDLL = FILENAMEDLL;
CString szOffice = L"32";

_ApplicationPtr xl;
_WorkbookPtr pWb;

_WorksheetPtr pShActive;
RangePtr pRgActive;
CString sCodeActive;
CString sNameActive;
COLORREF bkgColorActive = rgbWhite;
int iRowActive = 1, iColumnActive = 1;

bool blDragging = false;
int CLRow = -1, CLColumn = -1;
CColorEdit m_edit_rename, m_edit_properties, m_edit_autorename;

bool bCheckVirut = false;
bool bCheckVersion = false;
bool bRibbonSearch = false;

int jColumnHand = -1;
CFont m_FontHeaderList;
CFont m_FontItalic;

HWND hWndAutoRename;

vector<MyStrSelection> vecSelect;
bool blUndovalue = false;

int compareAZ(const void * a, const void * b)
{
	if (*(CString*)a < *(CString*)b) return -1;
	else if (*(CString*)a > *(CString*)b) return  1;
	return 0;
}

bool blCheckFileTemp()
{
	try
	{
		bool bl = false;
		Function *ff = new Function;
		CString bsConfig = (LPCTSTR)ff->_xlGetNameSheet("shConfig", 0);
		if (bsConfig != L"")
		{
			_WorksheetPtr pShConfig = xl->Sheets->GetItem((_bstr_t)bsConfig);
			RangePtr PRS = pShConfig->Cells->GetItem(1, 1);
			if (ff->_xlGetcomment(PRS) == TEMP_NAME) bl = true;
		}
		delete ff;

		return bl;
	}
	catch (...) {}
	return false;
}