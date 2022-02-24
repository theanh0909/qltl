#include "pch.h"

_ApplicationPtr xl;

_WorksheetPtr pShActive;
RangePtr pRgActive;
CString sCodeActive;
CString sNameActive;
int iRowActive, iColumnActive;

CConnection ObjConn;
CString SqlString;

HWND hWndChiphi;
HWND hWndQuytrinh;
HWND hWndTieuchuan;

HWND hWndMain;
CString szCellBefore = L"", szcellAfter = L"";
CString szKeyboard = L"";
int nIndexForm = 0;
int nIndexMouse = 0;

int jKeyBeforeF2 = 0;
bool blKeyEnter = false;

int dPosX = 0, dPosY = 0, dWidthF = 0, dHeightF = 0;

CFont m_FontHeaderList;

int CLRow = -1, CLColumn = -1;
int jColumnHand = -1;